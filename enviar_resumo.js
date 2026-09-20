// enviar_resumo.js — Resumo diário de BD_DESPESAS via WhatsApp (Evolution API)
// TODAS as credenciais vêm de variáveis de ambiente (GitHub Secrets). Nada sensível no código.

const {
  AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET,
  EVOLUTION_API_URL, EVOLUTION_INSTANCE, EVOLUTION_API_KEY, WHATSAPP_NUMERO
} = process.env;

const SITE_DOMAIN = "willtech7.sharepoint.com";
const SITE_PATH = "/sites/CARGOVAN";
const LISTA_DESPESAS = "BD_DESPESAS";

const sleep = (ms) => new Promise(r => setTimeout(r, ms));

function exigirEnv() {
  const obrigatorias = { AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET,
    EVOLUTION_API_URL, EVOLUTION_INSTANCE, EVOLUTION_API_KEY, WHATSAPP_NUMERO };
  const faltando = Object.entries(obrigatorias).filter(([, v]) => !v).map(([k]) => k);
  if (faltando.length) throw new Error(`Secrets ausentes no GitHub: ${faltando.join(", ")}`);
}

async function getGraphAccessToken() {
  const res = await fetch(`https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: AZURE_CLIENT_ID,
      client_secret: AZURE_CLIENT_SECRET,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials"
    })
  });
  const data = await res.json();
  if (!res.ok || !data.access_token) {
    throw new Error(`Falha no login Azure (${res.status}): ${data.error_description || JSON.stringify(data)}`);
  }
  return data.access_token;
}

async function graphGet(url, token) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Graph ${res.status} em ${url}\n${await res.text()}`);
  return res.json();
}

async function buscarDespesas(token) {
  const site = await graphGet(`https://graph.microsoft.com/v1.0/sites/${SITE_DOMAIN}:${SITE_PATH}`, token);
  let url = `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/${LISTA_DESPESAS}/items?expand=fields&$top=1000`;
  const itens = [];
  while (url) {                                   // paginação: a lista pode ter mais de 1000 itens
    const data = await graphGet(url, token);
    itens.push(...(data.value || []));
    url = data["@odata.nextLink"] || null;
  }
  return itens;
}

const pick = (f, ...keys) => { for (const k of keys) if (f[k] !== undefined && f[k] !== null && f[k] !== "") return f[k]; return ""; };
const num = (v) => {
  if (typeof v === "number") return v;
  const s = String(v ?? "").replace(/[^\d,.-]/g, "");
  if (!s) return 0;
  const n = s.includes(",") ? Number(s.replace(/\./g, "").replace(",", ".")) : Number(s);
  return Number.isFinite(n) ? n : 0;
};
const brl = (n) => n.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });

function resumir(itens) {
  const r = { pendentes: 0, recusados: 0, aprovados: 0, valorPendente: 0, maisAntigo: null };
  for (const it of itens) {
    const f = it.fields || {};
    const st = String(pick(f, "STATUS", "Status") || "Pendente").toLowerCase().trim();
    if (st.includes("recusado") || st.includes("rejeitado") || st.includes("cancelado")) r.recusados++;
    else if (st.includes("aprovado")) r.aprovados++;
    else {                                         // tudo que não é aprovado/recusado = pendente
      r.pendentes++;
      r.valorPendente += num(pick(f, "VALORTOTAL", "VALOR_x0020_TOTAL", "ValorTotal"));
      const d = new Date(pick(f, "DATA_x0020_DE_x0020_COMPRA", "DATACOMPRA") || f.Created || it.createdDateTime);
      if (!isNaN(d) && (!r.maisAntigo || d < r.maisAntigo)) r.maisAntigo = d;
    }
  }
  return r;
}

// Render (plano free) hiberna: acorda a API e confere se o WhatsApp está conectado
async function prepararEvolution() {
  const base = EVOLUTION_API_URL.replace(/\/$/, "");
  for (let i = 1; i <= 4; i++) {
    try {
      const res = await fetch(`${base}/instance/connectionState/${EVOLUTION_INSTANCE}`, {
        headers: { apikey: EVOLUTION_API_KEY }, signal: AbortSignal.timeout(90_000)
      });
      const body = await res.text();
      console.log(`Evolution connectionState (${res.status}): ${body}`);
      if (res.ok) {
        if (!/"state"\s*:\s*"open"/.test(body)) {
          throw new Error("Instância NÃO está conectada ao WhatsApp (state != open). Reescaneie o QR Code no Evolution.");
        }
        return base;
      }
      if (res.status === 401 || res.status === 404) throw new Error(`Evolution recusou (${res.status}): ${body}`);
    } catch (e) {
      if (/NÃO está conectada|recusou/.test(e.message) || i === 4) throw e;
      console.log(`Tentativa ${i} falhou (${e.message}). Aguardando 20s...`);
      await sleep(20_000);
    }
  }
}

async function enviarWhatsApp(base, texto) {
  const res = await fetch(`${base}/message/sendText/${EVOLUTION_INSTANCE}`, {
    method: "POST",
    headers: { "Content-Type": "application/json", apikey: EVOLUTION_API_KEY },
    body: JSON.stringify({ number: WHATSAPP_NUMERO, text: texto }),
    signal: AbortSignal.timeout(60_000)
  });
  const corpo = await res.text();
  if (!res.ok) throw new Error(`Falha ao enviar WhatsApp (${res.status}): ${corpo}`);
  console.log("WhatsApp enviado:", corpo);
}

(async () => {
  try {
    exigirEnv();
    console.log("1/4 Login Azure...");
    const token = await getGraphAccessToken();

    console.log("2/4 Lendo BD_DESPESAS...");
    const itens = await buscarDespesas(token);
    const r = resumir(itens);
    console.log(`${itens.length} itens | ${r.pendentes} pendentes, ${r.recusados} recusados, ${r.aprovados} aprovados`);

    const dias = r.maisAntigo ? Math.floor((Date.now() - r.maisAntigo) / 86_400_000) : null;
    const texto =
      `*Resumo Diário de Despesas - Cargo Van* 🚛\n\n` +
      `📌 *${r.pendentes}* pendentes (${brl(r.valorPendente)})\n` +
      (dias !== null ? `⏳ Mais antiga pendente há *${dias}* dia(s)\n` : "") +
      `❌ *${r.recusados}* recusadas\n` +
      `✅ *${r.aprovados}* aprovadas\n\n` +
      `Acesse o painel da Cargo Van > *Aprovar Despesas* para analisar.`;

    console.log("3/4 Verificando Evolution API...");
    const base = await prepararEvolution();

    console.log("4/4 Enviando...");
    await enviarWhatsApp(base, texto);
  } catch (e) {
    console.error("ERRO:", e.message);
    process.exit(1);
  }
})();
