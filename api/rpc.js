import { OAuth2Client } from "google-auth-library";
import { google } from "googleapis";

const IDS = {
  portal: "1aoQbDX1AAxrRCx8uaU24mkZaloG62zqbFiH182R58uo",
  base: "1S7gJzRixU2IlroN4nbIMXsq671QQUDb400y8hR8qC1s",
  agendamentos: "1xP7j5pT8HYaQZ9mLoW-59bnpgUbKsnpYuAyiZOF238s",
  remanejamento: "15Ai24tCW1biXgxe9LVJuZunbVLByQuzOKyjnvAQZvO8"
};

const ADMINS = ["wallafteste@gmail.com", "planilhasjh@gmail.com"];
const STATUS = ["pendente", "em rota", "iniciado", "concluído", "não concluído", "suspenso"];
const TIPOS = ["Ativação", "Mudança de Cômodo", "Mudança de Endereço", "Reparo Corretivo", "Upgrade/Downgrade"];
const MOTIVOS = ["Ativação", "Mudança de Endereço", "Reparo Corretivo"];

function sheetsClient() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON || "{}");
  if (!credentials.client_email || !credentials.private_key) {
    throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON não configurada.");
  }
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
  });
  return google.sheets({ version: "v4", auth });
}

async function userFromRequest(request) {
  const authorization = request.headers.authorization || "";
  const token = authorization.startsWith("Bearer ") ? authorization.slice(7) : "";
  if (!token) throw new Error("Autenticação Google necessária.");

  const ticket = await new OAuth2Client(process.env.GOOGLE_CLIENT_ID).verifyIdToken({
    idToken: token,
    audience: process.env.GOOGLE_CLIENT_ID
  });
  const payload = ticket.getPayload();
  if (!payload?.email || payload.email_verified === false) throw new Error("Conta Google inválida.");
  return payload.email.toLowerCase().trim();
}

function admin(email) { return ADMINS.includes(email); }
function colName(index) {
  let result = "";
  for (let value = index + 1; value; value = Math.floor((value - 1) / 26)) {
    result = String.fromCharCode(65 + ((value - 1) % 26)) + result;
  }
  return result;
}

async function readRange(sheets, spreadsheetId, range, render = "UNFORMATTED_VALUE") {
  const result = await sheets.spreadsheets.values.get({ spreadsheetId, range, valueRenderOption: render });
  return result.data.values || [];
}

async function appendRow(sheets, spreadsheetId, range, values) {
  await sheets.spreadsheets.values.append({
    spreadsheetId, range, valueInputOption: "USER_ENTERED", insertDataOption: "INSERT_ROWS",
    requestBody: { values: [values] }
  });
}

async function updateCell(sheets, spreadsheetId, range, value) {
  await sheets.spreadsheets.values.update({
    spreadsheetId, range, valueInputOption: "USER_ENTERED", requestBody: { values: [[value]] }
  });
}

function formatDate(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return new Intl.DateTimeFormat("pt-BR", { dateStyle: "short", timeStyle: "medium", timeZone: "America/Sao_Paulo" }).format(date);
}

async function links(sheets) {
  const rows = await readRange(sheets, IDS.portal, "PORTAL!A:H");
  return rows.slice(1).filter(row => row[1]).map(row => ({
    categoria: row[0] || "Geral", nome: row[1], link: row[2], icone: row[3] || "🌐",
    obs: row[4] || "", status: row[5] || "Online", ordem: row[6] || 99, acesso: String(row[7] || "Livre").trim()
  }));
}

async function permissions(sheets, email) {
  const rows = await readRange(sheets, IDS.portal, "PERMISSOES!A:B");
  return rows.filter(row => String(row[0] || "").toLowerCase().trim() === email).map(row => row[1]);
}

async function notice(sheets) {
  const rows = await readRange(sheets, IDS.portal, "AVISOS!A:B");
  const active = rows.slice(1).find(row => row[1] === true || String(row[1]).toUpperCase() === "TRUE");
  return active?.[0] || "";
}

async function lastUpdate(sheets, spreadsheetId, range) {
  const rows = await readRange(sheets, spreadsheetId, range, "FORMATTED_VALUE");
  return rows[0]?.[0] || "";
}

function weatherText(code) {
  const map = { 0: "☀️ Céu limpo", 1: "🌤️ Poucas nuvens", 2: "⛅ Parcialmente nublado", 3: "☁️ Nublado", 45: "🌫️ Neblina", 48: "🌫️ Neblina", 51: "🌦️ Garoa fraca", 53: "🌦️ Garoa", 55: "🌦️ Garoa forte", 61: "🌧️ Chuva fraca", 63: "🌧️ Chuva", 65: "🌧️ Chuva forte", 71: "🌨️ Neve fraca", 73: "🌨️ Neve", 75: "🌨️ Neve forte", 80: "🌦️ Pancadas de chuva", 81: "🌦️ Pancadas de chuva", 82: "⛈️ Pancadas fortes", 95: "⛈️ Trovoadas", 96: "⛈️ Trovoadas com granizo", 99: "⛈️ Trovoadas fortes" };
  return map[code] || "🌡️ Indisponível";
}

async function weather() {
  const url = "https://api.open-meteo.com/v1/forecast?latitude=-23.590036&longitude=-46.407726&current=temperature_2m,weather_code&timezone=America%2FSao_Paulo";
  const response = await fetch(url);
  if (!response.ok) return { temperatura: null, condicao: "" };
  const data = await response.json();
  return { temperatura: Math.round(data.current.temperature_2m), condicao: weatherText(data.current.weather_code) };
}

async function requests(sheets) {
  const rows = await readRange(sheets, IDS.portal, "SUGESTOES_ERROS!A:J", "FORMATTED_VALUE");
  return rows.slice(1).map(row => ({
    id: row[0] || "", data: formatDate(row[1]), email: row[2] || "", tipo: row[3] || "",
    descricao: row[4] || "", status: row[5] || "Pendente", resposta: row[6] || "", dataResposta: formatDate(row[7])
  }));
}

async function dashboard(sheets) {
  const rows = await readRange(sheets, IDS.base, "⚙️ BASE!A:AX");
  const table = Object.fromEntries(TIPOS.map(type => [type, Object.fromEntries([...STATUS, "Total"].map(key => [key, 0]))]));
  const totals = Object.fromEntries(STATUS.map(key => [key, 0]));
  const reasons = {};
  const periods = { "Ativação": { "Manhã": 0, "Tarde": 0 }, "Mudança de Endereço": { "Manhã": 0, "Tarde": 0 } };
  const bands = { "Menor que 24hrs": 0, "Menor que 36hrs": 0, "Menor que 48hrs": 0, Outlier: 0 };
  const resources = new Set();
  const normalize = value => STATUS.find(item => item.toLowerCase() === String(value || "").trim().toLowerCase());

  rows.slice(1).forEach(row => {
    const resource = String(row[0] || "").trim();
    if (resource && !["ITAIM", "SAO PAULO", "SAO PAULO_1", "SAO PAULO_2"].includes(resource)) resources.add(resource);
    const status = normalize(row[2]);
    const type = String(row[23] || "").trim();
    if (table[type] && status) { table[type][status]++; table[type].Total++; totals[status]++; }
    if (status === "não concluído" && MOTIVOS.includes(type) && String(row[29] || "").trim()) {
      const reason = String(row[29]).trim();
      reasons[reason] ||= Object.fromEntries(MOTIVOS.map(item => [item, 0]));
      reasons[reason][type]++;
    }
    if (status === "pendente" && periods[type]?.[row[11]] !== undefined) periods[type][row[11]]++;
    if (type === "Reparo Corretivo") {
      const band = String(row[49] || "").trim();
      if (bands[band] !== undefined) bands[band]++;
      else if (["Maior que 48hrs", "Outliers"].includes(band)) bands.Outlier++;
    }
  });
  const count = resources.size || 1;
  return { tabelaTipoStatus: table, somaLinha: totals, somaTotalGeral: Object.values(totals).reduce((sum, value) => sum + value, 0), motivos: reasons, pendentesPeriodo: periods, reparos: { faixas: bands, total: Object.values(bands).reduce((sum, value) => sum + value, 0) }, mediaConcluidos: Math.round((totals["concluído"] / count) * 10) / 10, mediaDespachada: Math.round(((Object.values(totals).reduce((sum, value) => sum + value, 0) - totals.suspenso) / count) * 10) / 10 };
}

async function execute(method, args, email) {
  const sheets = sheetsClient();
  if (method === "getPortalData") {
    const isAdmin = admin(email);
    const all = await requests(sheets);
    return { links: await links(sheets), aviso: await notice(sheets), user: email, admin: isAdmin, permissoes: await permissions(sheets, email), version: String(Date.now()), atualizacaoBase: await lastUpdate(sheets, IDS.base, "⚙️ BASE!CT1"), atualizacaoAgendamentos: await lastUpdate(sheets, IDS.agendamentos, "PAINEL!V2"), atualizacaoRemanejamento: await lastUpdate(sheets, IDS.remanejamento, "BASE DE DADOS!N2"), clima: await weather(), notificacoesUsuario: all.filter(item => item.email.toLowerCase() === email && item.resposta).length, notificacoesAdmin: isAdmin ? all.filter(item => item.id).length : 0 };
  }
  if (method === "registrarAcessoPortal") { await appendRow(sheets, IDS.portal, "ACESSOS!A:D", [email, args[0], new Date().toISOString(), args[1]]); return true; }
  if (method === "getMinhasSolicitacoes") return (await requests(sheets)).filter(item => item.email.toLowerCase() === email);
  if (method === "getTodasSolicitacoes") { if (!admin(email)) throw new Error("Acesso não autorizado."); return requests(sheets); }
  if (method === "enviarSolicitacao") { const rows = await requests(sheets); const id = `#${String(rows.length + 1).padStart(3, "0")}`; await appendRow(sheets, IDS.portal, "SUGESTOES_ERROS!A:J", [id, new Date().toISOString(), email, args[0], String(args[1] || "").trim(), "Pendente", "", "", "", "NÃO"]); return { sucesso: true, id }; }
  if (method === "marcarSolicitacoesLidasUsuario") { const rows = await readRange(sheets, IDS.portal, "SUGESTOES_ERROS!A:J"); for (let i = 1; i < rows.length; i++) if (String(rows[i][2] || "").toLowerCase().trim() === email && rows[i][6]) await updateCell(sheets, IDS.portal, `SUGESTOES_ERROS!I${i + 1}`, "SIM"); return true; }
  if (method === "marcarSolicitacaoLidaAdmin") { if (!admin(email)) throw new Error("Acesso não autorizado."); const rows = await readRange(sheets, IDS.portal, "SUGESTOES_ERROS!A:A"); const row = rows.findIndex(item => String(item[0]) === String(args[0])); if (row > 0) await updateCell(sheets, IDS.portal, `SUGESTOES_ERROS!J${row + 1}`, "SIM"); return true; }
  if (method === "atualizarSolicitacaoAdmin") { if (!admin(email)) throw new Error("Acesso não autorizado."); const rows = await readRange(sheets, IDS.portal, "SUGESTOES_ERROS!A:A"); const row = rows.findIndex(item => String(item[0]) === String(args[0])); if (row < 1) throw new Error("Solicitação não encontrada."); await updateCell(sheets, IDS.portal, `SUGESTOES_ERROS!F${row + 1}`, args[1]); await updateCell(sheets, IDS.portal, `SUGESTOES_ERROS!G${row + 1}`, args[2] || ""); await updateCell(sheets, IDS.portal, `SUGESTOES_ERROS!H${row + 1}`, new Date().toISOString()); await updateCell(sheets, IDS.portal, `SUGESTOES_ERROS!I${row + 1}`, "NÃO"); return { sucesso: true, mensagem: "Solicitação atualizada com sucesso." }; }
  if (method === "getDashboardProducao") return dashboard(sheets);
  throw new Error(`Método não suportado: ${method}`);
}

export default async function handler(request, response) {
  if (request.method !== "POST") return response.status(405).json({ error: "Método não permitido." });
  try {
    const email = await userFromRequest(request);
    const body = typeof request.body === "string" ? JSON.parse(request.body) : request.body;
    const result = await execute(body?.method, body?.args || [], email);
    return response.status(200).json({ result });
  } catch (error) {
    return response.status(error.message.includes("Autenticação") ? 401 : 500).json({ error: error.message });
  }
}