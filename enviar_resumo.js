// enviar_resumo.js — Resumo diário de BD_DESPESAS (STATUS) via WhatsApp (Evolution API)
// Credenciais seguras vindas dos Secrets do GitHub Actions.

const fs = require("fs");
const path = require("path");

const {
  AZURE_TENANT_ID,
  AZURE_CLIENT_ID,
  AZURE_CLIENT_SECRET,
  EVOLUTION_API_URL,
  EVOLUTION_INSTANCE,
  EVOLUTION_API_KEY,
  WHATSAPP_NUMERO
} = process.env;

/* ---------------- CONFIGURAÇÕES DO SHAREPOINT & POWER APPS ---------------- */
const SITE_DOMAIN = "willtech7.sharepoint.com";
const SITE_PATH = "/sites/CARGOVAN";
const LISTA_DESPESAS = "BD_DESPESAS";

// Link direto do seu aplicativo Power Apps
const URL_POWER_APPS = "https://apps.powerapps.com/play/e/default-669ab6c9-4a10-4796-a3dc-90f5e7d6ff40/a/6d3a4210-eb0c-4b05-833d-4dd7b9954bb0?tenantId=669ab6c9-4a10-4796-a3dc-90f5e7d6ff40&hint=ca56ffa4-162b-4cd7-bb88-a155a07dc94d&sourcetime=1788901458248&source=portal#";

const sleep = (ms) => new Promise(r => setTimeout(r, ms));
const limpar = (v) => String(v ?? "").trim();

function exigirEnv() {
  const obrigatorias = {
    AZURE_TENANT_ID,
    AZURE_CLIENT_ID,
    AZURE_CLIENT_SECRET,
    EVOLUTION_API_URL,
    EVOLUTION_INSTANCE,
    EVOLUTION_API_KEY,
    WHATSAPP_NUMERO
  };
  const faltando = Object.entries(obrigatorias).filter(([, v]) => !limpar(v)).map(([k]) => k);
  if (faltando.length) {
    throw new Error(`Secrets ausentes no GitHub: ${faltando.join(", ")}`);
  }
}

// Converte a imagem da pasta "imagens/banner_bot_tech_solutions.png" para Base64
function obterImagemBase64() {
  const caminhoImagem = path.join(__dirname, "imagens", "banner_bot_tech_solutions.png");
  if (!fs.existsSync(caminhoImagem)) {
    throw new Error(`A imagem não foi encontrada em: ${caminhoImagem}`);
  }
  const bitmap = fs.readFileSync(caminhoImagem);
  return `data:image/png;base64,${bitmap.toString("base64")}`;
}

/* ---------------- Azure / Graph ---------------- */
async function getGraphAccessToken() {
  const res = await fetch(`https://login.microsoftonline.com/${limpar(AZURE_TENANT_ID)}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: limpar(AZURE_CLIENT_ID),
      client_secret: limpar(AZURE_CLIENT_SECRET),
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

function diagnosticarToken(token) {
  const p = JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString());
  console.log("DIAGNÓSTICO TOKEN -> aud:", p.aud, "| tid:", p.tid, "| appid:", p.appid, "| roles:", p.roles);
  if (!p.roles || !p.roles.length) {
    console.log("ATENÇÃO: token SEM roles. A permissão está como Delegada ou falta 'Conceder consentimento do administrador'.");
  }
}

async function graphGet(url, token) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Graph ${res.status} em ${url}\n${await res.text()}`);
  return res.json();
}

async function obterSiteId(token) {
  try {
    const site = await graphGet(`https://graph.microsoft.com/v1.0/sites/${SITE_DOMAIN}:${SITE_PATH}`, token);
    return site.id;
  } catch (e) {
    console.log("Consulta direta do site falhou:", e.message.split("\n")[0]);
    console.log("Tentando via busca (sites?search=CARGOVAN)...");
    const busca = await graphGet(`https://graph.microsoft.com/v1.0/sites?search=CARGOVAN`, token);
    const achado = (busca.value || []).find(s => (s.webUrl || "").toLowerCase().includes("/sites/cargovan"));
    if (!achado) throw new Error("Site CARGOVAN não encontrado pelo Graph. Verifique permissão Sites.Read.All (Aplicativo) + consentimento.");
    return achado.id;
  }
}

async function buscarDespesas(token) {
  const siteId = await obterSiteId(token);
  let url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${LISTA_DESPESAS}/items?expand=fields&$top=1000`;
  const itens = [];
  while (url) {
    const data = await graphGet(url, token);
    itens.push(...(data.value || []));
    url = data["@odata.nextLink"] || null;
  }
  return itens;
}

/* ---------------- Contagem por STATUS ---------------- */
function contarStatus(itens) {
  const r = { pendentes: 0, recusados: 0, aprovados: 0 };
  for (const it of itens) {
    const f = it.fields || {};
    const st = limpar(f.STATUS ?? f.Status ?? "Pendente").toLowerCase();
    if (st.includes("recusado") || st.includes("rejeitado") || st.includes("cancelado")) r.recusados++;
    else if (st.includes("aprovado")) r.aprovados++;
    else r.pendentes++;
  }
  return r;
}

/* ---------------- Evolution API ---------------- */
async function prepararEvolution() {
  const base = limpar(EVOLUTION_API_URL).replace(/\/$/, "");
  const inst = limpar(EVOLUTION_INSTANCE);
  const key = limpar(EVOLUTION_API_KEY);
  for (let i = 1; i <= 4; i++) {
    try {
      const res = await fetch(`${base}/instance/connectionState/${inst}`, {
        headers: { apikey: key }, signal: AbortSignal.timeout(90_000)
      });
      const body = await res.text();
      console.log(`Evolution connectionState (${res.status}): ${body}`);
      if (res.ok) {
        if (!/"state"\s*:\s*"open"/.test(body)) {
          throw new Error("Instância NÃO está conectada ao WhatsApp. Reescaneie o QR Code no Evolution.");
        }
        return base;
      }
      if (res.status === 401 || res.status === 404) throw new Error(`Evolution recusou (${res.status}): ${body}`);
    } catch (e) {
      if (/NÃO está conectada|recusou/.test(e.message) || i === 4) throw e;
      console.log(`Tentativa ${i} falhou (${e.message}). Aguardando 20s (Render acordando)...`);
      await sleep(20_000);
    }
  }
}

async function enviarWhatsApp(base, texto) {
  const numeros = limpar(WHATSAPP_NUMERO)
    .split(",")
    .map(n => n.trim())
    .filter(Boolean);

  const imagemBase64 = obterImagemBase64();

  console.log(`Disparando envio para ${numeros.length} destinatário(s)...`);

  for (const num of numeros) {
    try {
      const res = await fetch(`${base}/message/sendMedia/${limpar(EVOLUTION_INSTANCE)}`, {
        method: "POST",
        headers: { 
          "Content-Type": "application/json", 
          apikey: limpar(EVOLUTION_API_KEY) 
        },
        body: JSON.stringify({
          number: num,
          media: imagemBase64,
          mediatype: "image",
          caption: texto
        }),
        signal: AbortSignal.timeout(60_000)
      });

      const corpo = await res.text();
      if (!res.ok) {
        console.error(`Falha ao enviar para ${num} (${res.status}): ${corpo}`);
      } else {
        console.log(`WhatsApp com imagem enviado com sucesso para ${num}!`);
      }
    } catch (err) {
      console.error(`Erro no envio para ${num}:`, err.message);
    }
    await sleep(2000);
  }
}

/* ---------------- Execução ---------------- */
(async () => {
  try {
    exigirEnv();

    console.log("1/4 Login Azure...");
    const token = await getGraphAccessToken();
    diagnosticarToken(token);

    console.log("2/4 Lendo BD_DESPESAS...");
    const itens = await buscarDespesas(token);
    const r = contarStatus(itens);
    console.log(`${itens.length} itens | ${r.pendentes} pendentes, ${r.recusados} recusados, ${r.aprovados} aprovados`);

    const texto =
      `*RELATÓRIO DE DESPESAS* 🚛\n` +
      `Resumo diário — Cargo Van\n\n` +
      `📊 *Status das Solicitações:*\n` +
      `⏳ *${r.pendentes}* Pendente(s) de Aprovação\n` +
      `✅ *${r.aprovados}* Aprovada(s)\n` +
      `❌ *${r.recusados}* Recusada(s)\n\n` +
      `📲 *Acesse o aplicativo pelo link:*\n` +
      `🔗 ${URL_POWER_APPS}\n\n` +
      `───────────────\n` +
      `🤖 _By Tech Solutions Bot_`;

    console.log("3/4 Verificando Evolution API...");
    const base = await prepararEvolution();

    console.log("4/4 Enviando WhatsApp...");
    await enviarWhatsApp(base, texto);
  } catch (e) {
    console.error("ERRO:", e.message);
    process.exit(1);
  }
})();
