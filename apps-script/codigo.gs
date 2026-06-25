// ============================================================
// CONFIGURAÇÕES — edite apenas esta seção
// ============================================================
const CONFIG = {
  SHEET_ID:          "11hNd1faQ9N1LLtK-2gwkL0_kYGw18w8IUBXnUTfA8ZA",
  ABA_PACIENTES:     "pacientes",
  ABA_RESPOSTAS:     "respostas",
  EMAIL_ALERTA:      "dra.ivanamelo25@gmail.com",
  DIAS_SEM_RESPOSTA: 3,
};

const URL_BASE = "https://ivanameloendocrino.github.io/acompanhamento-pacientes/";

// Colunas da aba Pacientes (1-based, para uso com getRange)
const P = {
  CODIGO:        1,  // A
  NOME:          2,  // B
  DIAGNOSTICO:   3,  // C
  WHATSAPP:      4,  // D  — número sem DDI, ex: 85998103840
  TOKEN:         5,  // E
  TOKEN_USADO:   6,  // F
  LINK_CHECKIN:  7,  // G
  DATA_DISPARO:  8,  // H
  LINK_WHATSAPP: 9,  // I
};

// ============================================================
// ENDPOINT GET — único ponto de entrada para requisições GET
// ============================================================
function doGet(e) {
  const params = e.parameter || {};
  const acao   = params.acao;

  if (acao === "dados") {
    return retornarDados();
  }

  if (acao === "enviarLembretes") {
    return enviarLembretesHttp(params);
  }

  return resposta({ ok: true, status: "Sistema de acompanhamento ativo." });
}

// ============================================================
// ENDPOINT POST — recebe respostas do formulário
// ============================================================
function doPost(e) {
  try {
    const dados    = JSON.parse(e.postData.contents);
    const token    = dados.token || "";
    const paciente = buscarPacientePorToken(token);

    if (!paciente) {
      return resposta({ ok: false, erro: "Token inválido ou expirado." });
    }

    invalidarToken(paciente.linha);
    gravarResposta(dados, paciente);
    verificarAlertas();

    return resposta({ ok: true, mensagem: "Resposta registrada com sucesso." });
  } catch (err) {
    return resposta({ ok: false, erro: err.message });
  }
}

// ============================================================
// DADOS PARA O DASHBOARD
// ============================================================
function retornarDados() {
  const ss        = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const pacientes = ss.getSheetByName(CONFIG.ABA_PACIENTES).getDataRange().getValues();
  const respostas = ss.getSheetByName(CONFIG.ABA_RESPOSTAS).getDataRange().getValues();

  return ContentService
    .createTextOutput(JSON.stringify({ pacientes, respostas }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// LEMBRETES QUINZENAIS — endpoint HTTP (chamado pelo GitHub Actions)
// ============================================================
function enviarLembretesHttp(params) {
  const segredoSalvo = PropertiesService.getScriptProperties().getProperty("LEMBRETE_SECRET");

  if (!segredoSalvo || params.segredo !== segredoSalvo) {
    return resposta({ ok: false, erro: "Não autorizado." });
  }

  const isDryRun = (params.dryRun === "true");
  const resultado = cicloLembretes(isDryRun);
  return resposta(resultado);
}

// ============================================================
// CICLO QUINZENAL — gera tokens, atualiza planilha e envia
// ============================================================
function cicloLembretes(dryRun) {
  dryRun = dryRun || false;

  const props      = PropertiesService.getScriptProperties();
  const baseUrl    = props.getProperty("UAZAPI_BASE_URL");   // https://bflabs.uazapi.com
  const instanceId = props.getProperty("UAZAPI_INSTANCE_ID"); // r564b2c65290e0b
  const apiToken   = props.getProperty("UAZAPI_TOKEN");       // a75cc4d7-...

  if (!dryRun && (!baseUrl || !instanceId || !apiToken)) {
    return { ok: false, erro: "Credenciais uazapi ausentes (UAZAPI_BASE_URL, UAZAPI_INSTANCE_ID, UAZAPI_TOKEN)." };
  }

  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const aba   = ss.getSheetByName(CONFIG.ABA_PACIENTES);
  const dados = aba.getDataRange().getValues();
  const hoje  = new Date();

  let enviados = 0, pulados = 0;
  const erros = [];

  for (let i = 1; i < dados.length; i++) {
    const codigo   = String(dados[i][P.CODIGO - 1]    || "").trim();
    const nome     = String(dados[i][P.NOME - 1]      || "").trim();
    const whatsapp = String(dados[i][P.WHATSAPP - 1]  || "").replace(/\D/g, "");

    if (!codigo || !nome || !whatsapp) { pulados++; continue; }

    const novoToken      = Utilities.getUuid();
    const telefone       = "55" + whatsapp;
    const linkCheckin    = URL_BASE + "?token=" + novoToken + "&codigo=" + encodeURIComponent(codigo);
    const primeiroNome   = nome.split(" ")[0];

    const mensagem =
      "Olá, " + primeiroNome + "! 👋\n\n" +
      "É hora do seu check-in quinzenal com a Dra. Ivana Melo.\n\n" +
      "Preencha o formulário de acompanhamento — leva menos de 2 minutos:\n\n" +
      "📋 " + linkCheckin + "\n\n" +
      "Sua participação é essencial para o sucesso do seu tratamento. 💚\n\n" +
      "— Consultório Dra. Ivana Melo · Endocrinologia";

    const linkWhatsapp = "https://wa.me/" + telefone + "?text=" + encodeURIComponent(mensagem);

    // Atualiza planilha com novo token e links
    const linha = i + 1;
    aba.getRange(linha, P.TOKEN).setValue(novoToken);
    aba.getRange(linha, P.TOKEN_USADO).setValue("nao");
    aba.getRange(linha, P.LINK_CHECKIN).setValue(linkCheckin);
    aba.getRange(linha, P.DATA_DISPARO).setValue(hoje);
    aba.getRange(linha, P.LINK_WHATSAPP).setValue(linkWhatsapp);

    if (!dryRun) {
      const res = enviarWhatsApp(telefone, mensagem, baseUrl, instanceId, apiToken);
      if (res.ok) {
        enviados++;
      } else {
        erros.push({ paciente: nome, codigo, erro: res.erro });
      }
      Utilities.sleep(1500);
    } else {
      Logger.log("[DRY RUN] " + nome + " → " + telefone + "\n" + linkCheckin);
      enviados++;
    }
  }

  return {
    ok: true,
    dryRun,
    enviados,
    pulados,
    total: dados.length - 1,
    erros,
    data: Utilities.formatDate(hoje, "America/Fortaleza", "dd/MM/yyyy"),
  };
}

// ============================================================
// INTEGRAÇÃO UAZAPI — envia mensagem WhatsApp
// Docs: https://docs.uazapi.com
// ============================================================
function enviarWhatsApp(telefone, mensagem, baseUrl, instanceId, token) {
  // Endpoint uazapi para envio de texto
  const url = baseUrl.replace(/\/$/, "") + "/message/sendText/" + instanceId;

  try {
    const resp = UrlFetchApp.fetch(url, {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        "token": token,
      },
      payload: JSON.stringify({
        number:  telefone,   // formato: 5585998103840
        text:    mensagem,
      }),
      muteHttpExceptions: true,
    });

    const code = resp.getResponseCode();
    let body   = {};
    try { body = JSON.parse(resp.getContentText()); } catch (_) {}

    return (code >= 200 && code < 300)
      ? { ok: true }
      : { ok: false, erro: body.message || body.error || "HTTP " + code };
  } catch (err) {
    return { ok: false, erro: err.message };
  }
}

// ============================================================
// GRAVA RESPOSTA NA ABA "respostas"
// ============================================================
function gravarResposta(dados, paciente) {
  const ss  = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const aba = ss.getSheetByName(CONFIG.ABA_RESPOSTAS);

  aba.appendRow([
    new Date(),
    paciente.codigo,
    paciente.nome,
    paciente.diagnostico,
    dados.peso              || "",
    dados.sentimento_peso   || "",
    dados.dieta             || "",
    (dados.dificuldades_dieta || []).join(", "),
    dados.dias_exercicio    || "",
    dados.intensidade       || "",
    dados.tipo_exercicio    || "",
    dados.medicacao         || "",
    dados.efeitos_colaterais || "",
    (dados.sintomas         || []).join(", "),
    dados.bem_estar         || "",
    dados.duvidas           || "",
  ]);
}

// ============================================================
// BUSCA PACIENTE PELO TOKEN
// ============================================================
function buscarPacientePorToken(token) {
  if (!token) return null;
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const aba   = ss.getSheetByName(CONFIG.ABA_PACIENTES);
  const dados = aba.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][P.TOKEN - 1] === token && dados[i][P.TOKEN_USADO - 1] !== "sim") {
      return {
        linha:       i + 1,
        codigo:      dados[i][P.CODIGO - 1],
        nome:        dados[i][P.NOME - 1],
        diagnostico: dados[i][P.DIAGNOSTICO - 1],
        whatsapp:    dados[i][P.WHATSAPP - 1],
      };
    }
  }
  return null;
}

// ============================================================
// INVALIDA TOKEN APÓS USO
// ============================================================
function invalidarToken(linha) {
  const ss  = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const aba = ss.getSheetByName(CONFIG.ABA_PACIENTES);
  aba.getRange(linha, P.TOKEN_USADO).setValue("sim");
}

// ============================================================
// VERIFICA ALERTAS — pacientes sem resposta há mais de N dias
// ============================================================
function verificarAlertas() {
  const ss            = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const abaPacientes  = ss.getSheetByName(CONFIG.ABA_PACIENTES);
  const abaRespostas  = ss.getSheetByName(CONFIG.ABA_RESPOSTAS);
  const pacientes     = abaPacientes.getDataRange().getValues();
  const respostas     = abaRespostas.getDataRange().getValues();
  const agora         = new Date();
  const alertas       = [];

  for (let i = 1; i < pacientes.length; i++) {
    const codigo      = pacientes[i][P.CODIGO - 1];
    const nome        = pacientes[i][P.NOME - 1];
    const dataDisparo = pacientes[i][P.DATA_DISPARO - 1];

    if (!dataDisparo || !codigo) continue;

    let ultimaResposta = null;
    for (let j = 1; j < respostas.length; j++) {
      if (respostas[j][1] === codigo) {
        const dataResp = new Date(respostas[j][0]);
        if (!ultimaResposta || dataResp > ultimaResposta) ultimaResposta = dataResp;
      }
    }

    const diasSemResposta = (agora - (ultimaResposta || new Date(dataDisparo))) / 86400000;

    if (diasSemResposta >= CONFIG.DIAS_SEM_RESPOSTA && !ultimaResposta) {
      alertas.push({ codigo, nome, diasSemResposta: Math.floor(diasSemResposta) });
    }
  }

  if (alertas.length > 0) enviarEmailAlerta(alertas);
}

// ============================================================
// ENVIA E-MAIL DE ALERTA
// ============================================================
function enviarEmailAlerta(alertas) {
  const linhas = alertas.map(a =>
    `• ${a.nome} (${a.codigo}) — ${a.diasSemResposta} dia(s) sem resposta`
  ).join("\n");

  GmailApp.sendEmail(
    CONFIG.EMAIL_ALERTA,
    `[Acompanhamento] ${alertas.length} paciente(s) sem resposta`,
    `Dra. Ivana,\n\nOs seguintes pacientes ainda não responderam ao check-in:\n\n${linhas}\n\nAcesse o dashboard para mais detalhes.\n\nSistema de Acompanhamento Automático`
  );
}

// ============================================================
// TESTES MANUAIS — execute diretamente no Apps Script
// ============================================================
function testarEnvio() {
  // Dry run: só loga, não envia e não atualiza planilha
  const r = cicloLembretes(true);
  Logger.log(JSON.stringify(r, null, 2));
}

function dispararEnvioManual() {
  // Envia de verdade — confirme antes de executar
  const r = cicloLembretes(false);
  Logger.log(JSON.stringify(r, null, 2));
}

// ============================================================
// UTILITÁRIO
// ============================================================
function resposta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
