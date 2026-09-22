// enviar_operacoes.js — Relatório de Operações (LIBERACAO_VEICULO)
const {
  AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET,
  EVOLUTION_API_URL, EVOLUTION_INSTANCE, EVOLUTION_API_KEY, WHATSAPP_NUMERO
} = process.env;

const SITE_DOMAIN = "willtech7.sharepoint.com";
const SITE_PATH = "/sites/CARGOVAN";
const NOME_LISTA = "LIBERACAO_VEICULO";
const URL_IMAGEM_RAW = "https://raw.githubusercontent.com/Willgner777/Cargo-van-html/main/imagens/banner_bot_tech_solutions.png";

const sleep = (ms) => new Promise(r => setTimeout(r, ms));
const limpar = (v) => String(v ?? "").trim();

async function getGraphAccessToken() {
  const res = await fetch(`https://login.microsoftonline.com/${limpar(AZURE_TENANT_ID)}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: limpar(AZURE_CLIENT_ID), client_secret: limpar(AZURE_CLIENT_SECRET),
      scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials"
    })
  });
  const data = await res.json();
  if (!res.ok || !data.access_token) throw new Error(`Erro Azure: ${data.error_description || JSON.stringify(data)}`);
  return data.access_token;
}

async function graphGet(url, token) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Erro Graph ${res.status}: ${await res.text()}`);
  return res.json();
}

async function obterSiteId(token) {
  const site = await graphGet(`https://graph.microsoft.com/v1.0/sites/${SITE_DOMAIN}:${SITE_PATH}`, token);
  return site.id;
}

async function buscarItens(token) {
  const siteId = await obterSiteId(token);
  let url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${NOME_LISTA}/items?expand=fields&$top=1000`;
  const itens = [];
  while (url) {
    const data = await graphGet(url, token);
    itens.push(...(data.value || []));
    url = data["@odata.nextLink"] || null;
  }
  return itens;
}

(async () => {
  try {
    const token = await getGraphAccessToken();
    const itens = await buscarItens(token);
    
    let emAndamento = 0, pernoite = 0, concluidos = 0;
    for (const it of itens) {
      const f = it.fields || {};
      const st = limpar(f.STATUS ?? f.Status ?? "").toUpperCase();
      if (st.includes("ANDAMENTO")) emAndamento++;
      else if (st.includes("PERNOITE")) pernoite++;
      else if (st.includes("CONCLUÍDO") || st.includes("CONCLUIDO")) concluidos++;
    }

    const dataAtual = new Intl.DateTimeFormat('pt-BR', { dateStyle: 'short', timeStyle: 'short', timeZone: 'America/Sao_Paulo' }).format(new Date());

    const texto =
      `🚚 *ACOMPANHAMENTO DE OPERAÇÕES* — Cargo Van\n\n` +
      `📊 *Status dos Veículos e Viagens:*\n` +
      `🔄 *${emAndamento}* Em Andamento\n` +
      `🌙 *${pernoite}* Em Pernoite\n` +
      `✅ *${concluidos}* Concluído(s)\n\n` +
      `──────────\n` +
      `🤖 _By Tech Solutions Bot_\n` +
      `🕒 _Atualizado em: ${dataAtual}_`;

    const base = limpar(EVOLUTION_API_URL).replace(/\/$/, "");
    const numeros = limpar(WHATSAPP_NUMERO).split(",").map(n => n.trim()).filter(Boolean);

    for (const num of numeros) {
      await fetch(`${base}/message/sendMedia/${limpar(EVOLUTION_INSTANCE)}`, {
        method: "POST",
        headers: { "Content-Type": "application/json", apikey: limpar(EVOLUTION_API_KEY) },
        body: JSON.stringify({ number: num, media: URL_IMAGEM_RAW, mediatype: "image", caption: texto })
      });
      await sleep(3000);
    }
    console.log("Relatório de Operações enviado com sucesso.");
  } catch (e) {
    console.error("Erro:", e.message);
    process.exit(1);
  }
})();
