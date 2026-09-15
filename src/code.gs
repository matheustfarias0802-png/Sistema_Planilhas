// ======================================================
// 🚀 CONFIGURAÇÃO GLOBAL - JH TELECOM
// ======================================================

const SS_ID = '1aoQbDX1AAxrRCx8uaU24mkZaloG62zqbFiH182R58uo'; 
const BASE_ID = '1S7gJzRixU2IlroN4nbIMXsq671QQUDb400y8hR8qC1s';
const AGENDAMENTOS_ID = '1xP7j5pT8HYaQZ9mLoW-59bnpgUbKsnpYuAyiZOF238s';
const REMANEJAMENTO_ID = '15Ai24tCW1biXgxe9LVJuZunbVLByQuzOKyjnvAQZvO8';

const ADMINS_PORTAL = [
  "wallafteste@gmail.com"
];

// Tempo de cache (em segundos) para dados que não precisam ser lidos
// da planilha em toda requisição. 6 minutos é seguro porque o cache
// é invalidado manualmente sempre que a planilha é editada (ver onEdit).
const CACHE_TTL = 360;

function isAdminPortal() {
  const email = getUser().toLowerCase().trim();
  return ADMINS_PORTAL
    .map(e => e.toLowerCase().trim())
    .includes(email);
}

// Mantida por compatibilidade (não é mais chamada separadamente pelo
// frontend, pois getPortalData() já devolve "admin" — ver seção RPC).
function verificarAdminPortal() {
  return isAdminPortal();
}

const COR_JH = "#09124f";

function doGet() {
  const faviconUrl =
    'https://drive.google.com/uc?id=1mhpwI0HqS2M5uehF3tDQVHajL0RmM9Dy&export=download&format=png';

  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('JH Telecom - Portal')
    .setFaviconUrl(faviconUrl)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getUser() {
  try { return Session.getActiveUser().getEmail() || "Visitante"; } catch (e) { return "E-mail Oculto"; }
}

// ======================================================
// 📊 LINKS (com cache)
// ======================================================
function getLinks() {
  const cache = CacheService.getScriptCache();
  const cacheado = cache.get("CACHE_LINKS");
  if (cacheado) {
    try { return JSON.parse(cacheado); } catch (e) { /* cache corrompido, recalcula abaixo */ }
  }

  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sh = ss.getSheetByName("PORTAL");
    if (!sh) return [];
    const data = sh.getDataRange().getValues();
    if (data.length <= 1) return [];
    data.shift();
    const resultado = data.filter(r => r[1]).map(r => ({
      categoria: r[0] || "Geral",
      nome: r[1],
      link: r[2],
      icone: r[3] || "🌐",
      obs: r[4] || "",
      status: r[5] || "Online",
      ordem: r[6] || 99,
      acesso: (r[7] || "Livre").toString().trim()
    }));

    cache.put("CACHE_LINKS", JSON.stringify(resultado), CACHE_TTL);
    return resultado;
  } catch (e) { return []; }
}

// ======================================================
// 🔐 PERMISSÕES (com cache por usuário)
// ======================================================
function getMinhasPermissoes() {
  const email = getUser().toLowerCase().trim();
  const cache = CacheService.getScriptCache();
  const chave = "CACHE_PERM_" + email;
  const cacheado = cache.get(chave);
  if (cacheado) {
    try { return JSON.parse(cacheado); } catch (e) { /* recalcula abaixo */ }
  }

  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sh = ss.getSheetByName("PERMISSOES");
    if (!sh) return [];
    const dados = sh.getDataRange().getValues();
    const resultado = dados
      .filter(r => r[0] && r[0].toString().toLowerCase().trim() === email)
      .map(r => r[1]);

    cache.put(chave, JSON.stringify(resultado), CACHE_TTL);
    return resultado;
  } catch (e) { return []; }
}

function registrarAcessoPortal(nome, statusAcesso) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    let sh = ss.getSheetByName("ACESSOS") || ss.insertSheet("ACESSOS");
    sh.appendRow([getUser(), nome, new Date(), statusAcesso]);
  } catch (e) {}
}

// ======================================================
// 🔄 ENGINE SYNC — ÚNICO PONTO DE ENTRADA DO FRONTEND
// ------------------------------------------------------
// Antes existiam várias chamadas google.script.run separadas
// (getPortalData, verificarAdminPortal, verificarNotificacoesSolicitante,
// verificarNotificacoesAdmin) feitas em paralelo a cada carregamento e a
// cada 60s. Cada uma delas é um round-trip de rede + execução de servidor.
// Agora TUDO é devolvido numa chamada só, reduzindo latência e a
// quantidade de vezes que a planilha é aberta.
// ======================================================
function getPortalData() {
  const admin = isAdminPortal();
  return {
    links: getLinks(),
    aviso: getAviso(),
    user: getUser(),
    admin: admin,
    permissoes: getMinhasPermissoes(),
    version: getPortalVersion(),
    atualizacaoBase: getAtualizacaoBase(),
    atualizacaoAgendamentos: getAtualizacaoAgendamentos(),
    atualizacaoRemanejamento: getAtualizacaoRemanejamento(),
    clima: getClima(),
    notificacoesUsuario: verificarNotificacoesSolicitante(),
    notificacoesAdmin: admin ? verificarNotificacoesAdmin() : 0
  };
}

function getAviso() {
  const cache = CacheService.getScriptCache();
  const cacheado = cache.get("CACHE_AVISO");
  if (cacheado !== null) return cacheado; // pode ser string vazia válida

  let aviso = "";
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sh = ss.getSheetByName("AVISOS");
    const data = sh ? sh.getDataRange().getValues() : [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === true || data[i][1] === "TRUE") { aviso = data[i][0]; break; }
    }
  } catch (e) {}

  cache.put("CACHE_AVISO", aviso, CACHE_TTL);
  return aviso;
}

function getPortalVersion(){
  const cache = CacheService.getScriptCache();
  return cache.get("PORTAL_VERSION") || Date.now().toString();
}

function atualizarVersaoPortal(){
  CacheService.getScriptCache().put("PORTAL_VERSION", Date.now().toString(), 21600);
}

function checkUpdate(){ return getPortalVersion(); }

// Invalida os caches relevantes sempre que a planilha muda, em vez de
// depender apenas do TTL. Assim os dados ficam atualizados na hora,
// mas sem precisar reler a sheet em toda requisição normal.
function limparCachePortal() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(["CACHE_LINKS", "CACHE_AVISO"]);
  // Obs: CACHE_PERM_<email> não é limpo aqui pois não sabemos todos os
  // e-mails; ele expira sozinho em CACHE_TTL segundos, o que é aceitável
  // para uma aba que muda com pouca frequência.
}

function onEdit(e){
  const aba = e.range.getSheet().getName();
  if (["PORTAL", "AVISOS", "PERMISSOES"].includes(aba)) {
    atualizarVersaoPortal();
    limparCachePortal();
    if(aba === "PORTAL") criarDropdownPermissoes();
  }
}

// ======================================================
// 🛠️ SETUP AUTOMÁTICO
// ======================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🚀 JH TELECOM')
    .addItem('Sincronizar e Configurar', 'configurarPortal')
    .addToUi();
}

function configurarPortal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPortal = ss.getSheetByName("PORTAL");
  if (shPortal && shPortal.getLastColumn() < 8) {
    shPortal.getRange(1, 8).setValue("Acesso").setBackground(COR_JH).setFontColor("white");
  }
  if (!ss.getSheetByName("ACESSOS")) {
    let sh = ss.insertSheet("ACESSOS");
    sh.appendRow(["Email", "Sistema", "Data", "Status"]).getRange("A1:D1").setBackground(COR_JH).setFontColor("white");
  }
  if (!ss.getSheetByName("PERMISSOES")) {
    let sh = ss.insertSheet("PERMISSOES");
    sh.appendRow(["Email do Usuário", "Sistema Restrito"]).getRange("A1:B1").setBackground(COR_JH).setFontColor("white");
    sh.setColumnWidth(1, 250); sh.setColumnWidth(2, 250);
  }
  criarDropdownPermissoes();
  SpreadsheetApp.getUi().alert("✅ Configuração concluída!");
}

function criarDropdownPermissoes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPortal = ss.getSheetByName("PORTAL");
  const shPerm = ss.getSheetByName("PERMISSOES");
  if (!shPortal || !shPerm) return;
  const dados = shPortal.getRange(2, 2, shPortal.getLastRow(), 7).getValues();
  const restritos = dados.filter(r => r[6] === "Restrito").map(r => r[0]);
  if(restritos.length > 0) {
    const regra = SpreadsheetApp.newDataValidation().requireValueInList(restritos).build();
    shPerm.getRange("B2:B100").setDataValidation(regra);
  }
}

// ======================================================
// 📅 ATUALIZAÇÕES DE BASES EXTERNAS (com cache)
// ------------------------------------------------------
// Antes de cada uma abria uma planilha externa a cada 60s por usuário
// ativo. Agora ficam em cache por CACHE_TTL segundos.
// ======================================================
function getAtualizacaoBase() {
  const cache = CacheService.getScriptCache();
  const cacheado = cache.get("CACHE_ATU_BASE");
  if (cacheado !== null) return cacheado;

  let valor = "";
  try {
    const ss = SpreadsheetApp.openById(BASE_ID);
    const sh = ss.getSheetByName("⚙️ BASE");
    if (sh) valor = sh.getRange("CT1").getDisplayValue();
  } catch (e) {
    Logger.log("Erro getAtualizacaoBase: " + e.message);
  }

  cache.put("CACHE_ATU_BASE", valor, CACHE_TTL);
  return valor;
}

function testeAtualizacao() {
  Logger.log(getAtualizacaoBase());
}

function getAtualizacaoAgendamentos() {
  const cache = CacheService.getScriptCache();
  const cacheado = cache.get("CACHE_ATU_AGENDAMENTOS");
  if (cacheado !== null) return cacheado;

  let valor = "";
  try {
    const ss = SpreadsheetApp.openById(AGENDAMENTOS_ID);
    const sh = ss.getSheetByName("PAINEL");
    if (sh) valor = sh.getRange("V2").getDisplayValue();
  } catch (e) {
    Logger.log("Erro getAtualizacaoAgendamentos: " + e.message);
  }

  cache.put("CACHE_ATU_AGENDAMENTOS", valor, CACHE_TTL);
  return valor;
}

function getAtualizacaoRemanejamento() {
  const cache = CacheService.getScriptCache();
  const cacheado = cache.get("CACHE_ATU_REMANEJAMENTO");
  if (cacheado !== null) return cacheado;

  let valor = "";
  try {
    const ss = SpreadsheetApp.openById(REMANEJAMENTO_ID);
    const sh = ss.getSheetByName("BASE DE DADOS");
    if (sh) valor = sh.getRange("N2").getDisplayValue();
  } catch (e) {
    Logger.log("Erro getAtualizacaoRemanejamento: " + e.message);
  }

  cache.put("CACHE_ATU_REMANEJAMENTO", valor, CACHE_TTL);
  return valor;
}

// ======================================================
// 📈 DASHBOARD DE PRODUÇÃO — dados da "Planilha de Prazos"
// ======================================================

const RECURSOS_IGNORADOS = ["ITAIM", "SAO PAULO", "SAO PAULO_1", "SAO PAULO_2", ""];
const TIPOS_ATIVIDADE = ["Ativação", "Mudança de Cômodo", "Mudança de Endereço", "Reparo Corretivo", "Upgrade/Downgrade"];
const STATUS_LISTA = ["pendente", "em rota", "iniciado", "concluído", "não concluído", "suspenso"];
const TIPOS_MOTIVO = ["Ativação", "Mudança de Endereço", "Reparo Corretivo"];

function getDashboardProducao(forcar) {
  const cache = CacheService.getScriptCache();

  if (!forcar) {
    const cacheado = cache.get("CACHE_DASHBOARD_PRODUCAO");
    if (cacheado) {
      try { return JSON.parse(cacheado); } catch (e) { /* recalcula abaixo */ }
    }
  }

  const resultado = calcularDashboardProducao();
  if (resultado) cache.put("CACHE_DASHBOARD_PRODUCAO", JSON.stringify(resultado), CACHE_TTL);
  return resultado;
}

// Normaliza o texto do status da planilha (que vem em minúsculo) para
// bater com os nomes "oficiais" em STATUS_LISTA, sem diferenciar
// maiúsculas/minúsculas.
function normalizarStatus(valor) {
  const texto = String(valor || "").trim().toLowerCase();
  return STATUS_LISTA.find(s => s.toLowerCase() === texto) || null;
}

function calcularDashboardProducao() {
  try {
    const ss = SpreadsheetApp.openById(BASE_ID);
    const sh = ss.getSheetByName("⚙️ BASE");
    if (!sh) return null;

    const ultimaLinha = sh.getLastRow();
    if (ultimaLinha < 2) return null;

    // A=1, C=3, L=12, X=24, AD=30, AX=50
    const colA = sh.getRange(2, 1, ultimaLinha - 1, 1).getValues().map(r => r[0]);
    const colC = sh.getRange(2, 3, ultimaLinha - 1, 1).getValues().map(r => r[0]);
    const colL = sh.getRange(2, 12, ultimaLinha - 1, 1).getValues().map(r => r[0]);
    const colX = sh.getRange(2, 24, ultimaLinha - 1, 1).getValues().map(r => r[0]);
    const colAD = sh.getRange(2, 30, ultimaLinha - 1, 1).getValues().map(r => r[0]);
    const colAX = sh.getRange(2, 50, ultimaLinha - 1, 1).getValues().map(r => r[0]);

    const total = colA.length;

    const tabelaTipoStatus = {};
    TIPOS_ATIVIDADE.forEach(tipo => {
      tabelaTipoStatus[tipo] = {};
      STATUS_LISTA.forEach(status => tabelaTipoStatus[tipo][status] = 0);
      tabelaTipoStatus[tipo].Total = 0;
    });
    const somaLinha = {};
    STATUS_LISTA.forEach(s => somaLinha[s] = 0);
    let somaTotalGeral = 0;

    for (let i = 0; i < total; i++) {
      const tipo = String(colX[i] || "").trim();
      const status = normalizarStatus(colC[i]);
      if (tabelaTipoStatus[tipo] && status) {
        tabelaTipoStatus[tipo][status]++;
        tabelaTipoStatus[tipo].Total++;
        somaLinha[status]++;
        somaTotalGeral++;
      }
    }

    const motivosMap = {};
    for (let i = 0; i < total; i++) {
      const status = normalizarStatus(colC[i]);
      if (status !== "não concluído") continue;
      const tipo = String(colX[i] || "").trim();
      const motivo = String(colAD[i] || "").trim();
      if (!motivo || !TIPOS_MOTIVO.includes(tipo)) continue;
      if (!motivosMap[motivo]) {
        motivosMap[motivo] = {};
        TIPOS_MOTIVO.forEach(t => motivosMap[motivo][t] = 0);
      }
      motivosMap[motivo][tipo]++;
    }

    const pendentesPeriodo = {
      "Ativação": { "Manhã": 0, "Tarde": 0 },
      "Mudança de Endereço": { "Manhã": 0, "Tarde": 0 }
    };
    for (let i = 0; i < total; i++) {
      const tipo = String(colX[i] || "").trim();
      const status = normalizarStatus(colC[i]);
      const periodo = String(colL[i] || "").trim();
      if (status !== "pendente") continue;
      if (pendentesPeriodo[tipo] && (periodo === "Manhã" || periodo === "Tarde")) {
        pendentesPeriodo[tipo][periodo]++;
      }
    }

    const faixas = { "Menor que 24hrs": 0, "Menor que 36hrs": 0, "Menor que 48hrs": 0, "Outlier": 0 };
    let totalReparos = 0;
    for (let i = 0; i < total; i++) {
      const tipo = String(colX[i] || "").trim();
      if (tipo !== "Reparo Corretivo") continue;
      const faixa = String(colAX[i] || "").trim();
      if (faixas.hasOwnProperty(faixa)) {
        faixas[faixa]++;
        totalReparos++;
      } else if (faixa === "Maior que 48hrs" || faixa === "Outliers") {
        faixas["Outlier"]++;
        totalReparos++;
      }
    }

    const recursosValidos = new Set();
    colA.forEach(v => {
      const nome = String(v || "").trim();
      if (nome && !RECURSOS_IGNORADOS.includes(nome)) recursosValidos.add(nome);
    });
    const qtdRecursos = recursosValidos.size || 1;

    const qtdConcluidos = somaLinha["concluído"] || 0;
    const qtdSuspensos = somaLinha["suspenso"] || 0;

    return {
      tabelaTipoStatus: tabelaTipoStatus,
      somaLinha: somaLinha,
      somaTotalGeral: somaTotalGeral,
      motivos: motivosMap,
      pendentesPeriodo: pendentesPeriodo,
      reparos: { faixas: faixas, total: totalReparos },
      mediaConcluidos: Math.round((qtdConcluidos / qtdRecursos) * 10) / 10,
      mediaDespachada: Math.round(((somaTotalGeral - qtdSuspensos) / qtdRecursos) * 10) / 10
    };

  } catch (e) {
    Logger.log("Erro calcularDashboardProducao: " + e.message);
    return null;
  }
}

function testeAtualizacaoRemanejamento() {
  try {
    const ss = SpreadsheetApp.openById(REMANEJAMENTO_ID);
    const sh = ss.getSheetByName("BASE DE DADOS");
    if (!sh) {
      Logger.log("❌ Aba BASE DE DADOS não encontrada");
      return;
    }
    const valor = sh.getRange("N2").getDisplayValue();
    Logger.log("📅 REMANEJAMENTO N2: " + valor);
  } catch (e) {
    Logger.log("❌ ERRO: " + e.message);
  }
}

// ======================================================
// 🌦️ CLIMA — Rua Arnaldo Bonaventura, 633 (Cidade Tiradentes, SP)
// ------------------------------------------------------
// Usa a API pública Open-Meteo (gratuita, sem necessidade de chave/API
// key). Coordenadas fixas do endereço informado. Resultado cacheado
// junto com os outros dados (mesmo CACHE_TTL) para não bater na API
// toda vez que alguém carregar a página.
// ======================================================
const CLIMA_LAT = -23.590036;
const CLIMA_LON = -46.407726;

function getClima() {
  const cache = CacheService.getScriptCache();
  const cacheado = cache.get("CACHE_CLIMA");
  if (cacheado) {
    try { return JSON.parse(cacheado); } catch (e) { /* recalcula abaixo */ }
  }

  let resultado = { temperatura: null, condicao: "" };
  let sucesso = false;

  try {
    const url = "https://api.open-meteo.com/v1/forecast"
      + "?latitude=" + CLIMA_LAT
      + "&longitude=" + CLIMA_LON
      + "&current=temperature_2m,weather_code"
      + "&timezone=America%2FSao_Paulo";

    const resposta = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

    if (resposta.getResponseCode() === 200) {
      const dados = JSON.parse(resposta.getContentText());

      if (dados && dados.current) {
        resultado = {
          temperatura: Math.round(dados.current.temperature_2m),
          condicao: traduzirCodigoClima(dados.current.weather_code)
        };
        sucesso = true;
      }
    } else {
      Logger.log("getClima: API respondeu com código " + resposta.getResponseCode());
    }
  } catch (e) {
    Logger.log("Erro getClima: " + e.message);
  }

  // Só guarda em cache se a chamada deu certo. Se falhou, a próxima
  // requisição tenta de novo imediatamente, em vez de ficar 6 minutos
  // mostrando o widget vazio.
  if (sucesso) {
    cache.put("CACHE_CLIMA", JSON.stringify(resultado), CACHE_TTL);
  }

  return resultado;
}

// Converte os códigos numéricos da Open-Meteo (padrão WMO) em texto
// simples + emoji para exibir no portal.
function traduzirCodigoClima(codigo) {
  const mapa = {
    0: "☀️ Céu limpo",
    1: "🌤️ Poucas nuvens",
    2: "⛅ Parcialmente nublado",
    3: "☁️ Nublado",
    45: "🌫️ Neblina",
    48: "🌫️ Neblina",
    51: "🌦️ Garoa fraca",
    53: "🌦️ Garoa",
    55: "🌦️ Garoa forte",
    61: "🌧️ Chuva fraca",
    63: "🌧️ Chuva",
    65: "🌧️ Chuva forte",
    71: "🌨️ Neve fraca",
    73: "🌨️ Neve",
    75: "🌨️ Neve forte",
    80: "🌦️ Pancadas de chuva",
    81: "🌦️ Pancadas de chuva",
    82: "⛈️ Pancadas fortes",
    95: "⛈️ Trovoadas",
    96: "⛈️ Trovoadas com granizo",
    99: "⛈️ Trovoadas fortes"
  };
  return mapa[codigo] || "🌡️ Indisponível";
}

// ======================================================
// 💡 SUGESTÕES & ERROS
// ======================================================

// Gera um ID sequencial confiável usando PropertiesService, em vez de
// contar linhas da planilha (que quebra se uma linha for excluída no
// meio, podendo gerar IDs duplicados).
function gerarProximoIdSolicitacao() {
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const atual = Number(props.getProperty("ULTIMO_ID_SOLICITACAO") || "0");
    const proximo = atual + 1;
    props.setProperty("ULTIMO_ID_SOLICITACAO", String(proximo));
    return "#" + String(proximo).padStart(3, "0");
  } finally {
    lock.releaseLock();
  }
}

function enviarSolicitacao(tipo, descricao) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sh = ss.getSheetByName("SUGESTOES_ERROS");

    if (!sh) {
      throw new Error("A aba SUGESTOES_ERROS não foi encontrada.");
    }

    if (!descricao || !descricao.toString().trim()) {
      throw new Error("A descrição não pode estar vazia.");
    }

    const email = getUser();
    const id = gerarProximoIdSolicitacao();

    sh.appendRow([
      id,
      new Date(),
      email,
      tipo,
      descricao.toString().trim(),
      "Pendente",
      "",
      "",
      "",
      "NÃO"
    ]);

    return {
      sucesso: true,
      id: id
    };

  } catch (e) {
    Logger.log("Erro enviarSolicitacao: " + e.message);
    return {
      sucesso: false,
      erro: e.message
    };
  }
}

function getMinhasSolicitacoes() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sh = ss.getSheetByName("SUGESTOES_ERROS");
    if (!sh) return [];

    const emailUsuario = getUser().toString().trim().toLowerCase();
    const dados = sh.getDataRange().getValues();
    if (dados.length <= 1) return [];

    const fuso = Session.getScriptTimeZone();

    return dados.slice(1)
      .filter(r => {
        const emailPlanilha = r[2] ? r[2].toString().trim().toLowerCase() : "";
        return emailPlanilha === emailUsuario;
      })
      .map(r => ({
        id: r[0],
        data: r[1] instanceof Date
          ? Utilities.formatDate(r[1], fuso, "dd/MM/yyyy HH:mm:ss")
          : r[1] || "",
        email: r[2] || "",
        tipo: r[3] || "",
        descricao: r[4] || "",
        status: r[5] || "Pendente",
        resposta: r[6] || "",
        dataResposta: r[7] instanceof Date
          ? Utilities.formatDate(r[7], fuso, "dd/MM/yyyy HH:mm:ss")
          : r[7] || ""
      }));

  } catch (e) {
    Logger.log("Erro getMinhasSolicitacoes: " + e.message);
    return [];
  }
}

function testeSolicitacoesWeb() {
  return {
    ok: true,
    email: getUser(),
    quantidade: 1,
    mensagem: "Teste funcionando"
  };
}

function getTodasSolicitacoes() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sh = ss.getSheetByName("SUGESTOES_ERROS");
    if (!sh) return [];

    const dados = sh.getDataRange().getValues();
    if (dados.length <= 1) return [];

    const fuso = Session.getScriptTimeZone();

    return dados.slice(1).map(r => ({
      id: r[0] || "",
      data: r[1] instanceof Date
        ? Utilities.formatDate(r[1], fuso, "dd/MM/yyyy HH:mm:ss")
        : r[1] || "",
      email: r[2] || "",
      tipo: r[3] || "",
      descricao: r[4] || "",
      status: r[5] || "Pendente",
      resposta: r[6] || "",
      dataResposta: r[7] instanceof Date
        ? Utilities.formatDate(r[7], fuso, "dd/MM/yyyy HH:mm:ss")
        : r[7] || ""
    }));

  } catch (e) {
    Logger.log("Erro getTodasSolicitacoes: " + e.message);
    return [];
  }
}

function atualizarSolicitacaoAdmin(id, status, resposta) {
  try {
    if (!isAdminPortal()) {
      throw new Error("Acesso não autorizado.");
    }

    const ss = SpreadsheetApp.openById(SS_ID);
    const sh = ss.getSheetByName("SUGESTOES_ERROS");
    if (!sh) {
      throw new Error("Aba SUGESTOES_ERROS não encontrada.");
    }

    const dados = sh.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(id)) {
        sh.getRange(i + 1, 6).setValue(status);          // F = Status
        sh.getRange(i + 1, 7).setValue(resposta || "");  // G = Resposta
        sh.getRange(i + 1, 8).setValue(new Date());      // H = Data Resposta
        sh.getRange(i + 1, 9).setValue("NÃO");            // I = Lida Usuário

        return {
          sucesso: true,
          mensagem: "Solicitação atualizada com sucesso."
        };
      }
    }

    throw new Error("Solicitação não encontrada.");

  } catch (e) {
    Logger.log("Erro atualizarSolicitacaoAdmin: " + e.message);
    throw e;
  }
}

function verificarNotificacoesSolicitante() {
  const email = getUser().toLowerCase().trim();

  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName("SUGESTOES_ERROS");
  if (!sh) return 0;

  const dados = sh.getDataRange().getValues();
  let quantidade = 0;

  for (let i = 1; i < dados.length; i++) {
    const emailSolicitante = String(dados[i][2] || "").toLowerCase().trim();
    const resposta = String(dados[i][6] || "").trim();
    const lidaUsuario = String(dados[i][8] || "").trim().toUpperCase();

    if (emailSolicitante === email && resposta !== "" && lidaUsuario !== "SIM") {
      quantidade++;
    }
  }

  return quantidade;
}

function verificarNotificacoesAdmin() {
  if (!isAdminPortal()) return 0;

  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName("SUGESTOES_ERROS");
  if (!sh) return 0;

  const dados = sh.getDataRange().getValues();
  let quantidade = 0;

  for (let i = 1; i < dados.length; i++) {
    const id = String(dados[i][0] || "").trim();
    const lidaAdmin = String(dados[i][9] || "").trim().toUpperCase();

    if (id && lidaAdmin !== "SIM") {
      quantidade++;
    }
  }

  return quantidade;
}

function marcarSolicitacoesLidasUsuario() {
  const email = getUser().toLowerCase().trim();

  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName("SUGESTOES_ERROS");
  if (!sh) {
    throw new Error("Aba SUGESTOES_ERROS não encontrada.");
  }

  const dados = sh.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    const emailSolicitante = String(dados[i][2] || "").toLowerCase().trim();
    const resposta = String(dados[i][6] || "").trim();

    if (emailSolicitante === email && resposta !== "") {
      sh.getRange(i + 1, 9).setValue("SIM"); // I = Lida Usuário
    }
  }

  return true;
}

// OBS: esta função existia DUPLICADA no arquivo original (uma versão
// aqui no Code.gs e outra, idêntica, dentro do <script> do index.html
// — essa segunda cópia nem funcionaria no client-side, pois usa
// SpreadsheetApp, que só existe no servidor). Mantida apenas esta versão.
function marcarSolicitacaoLidaAdmin(id) {
  if (!isAdminPortal()) {
    throw new Error("Acesso não autorizado.");
  }

  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName("SUGESTOES_ERROS");
  if (!sh) {
    throw new Error("Aba SUGESTOES_ERROS não encontrada.");
  }

  const dados = sh.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(id)) {
      sh.getRange(i + 1, 10).setValue("SIM"); // J = Lida Admin
      return true;
    }
  }

  return false;
}

function solicitarAutorizacao() {
  SpreadsheetApp.getActiveSpreadsheet().getName();
  UrlFetchApp.fetch("https://www.google.com");
}

function debugStatusDashboard() {
  const ss = SpreadsheetApp.openById(BASE_ID);
  const sh = ss.getSheetByName("⚙️ BASE");
  const dados = sh.getRange(2, 3, 10, 1).getValues(); // 10 primeiras linhas da coluna C

  dados.forEach((linha, i) => {
    const valor = linha[0];
    Logger.log(`Linha ${i + 2}: "${valor}" (tipo: ${typeof valor})`);
  });
}

function limparCacheDashboard() {
  CacheService.getScriptCache().remove("CACHE_DASHBOARD_PRODUCAO");
  Logger.log("Cache do dashboard limpo!");
}