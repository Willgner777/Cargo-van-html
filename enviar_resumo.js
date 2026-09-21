// enviar_resumo.js — Resumo diário de BD_DESPESAS (STATUS) via WhatsApp (Evolution API)
// Credenciais seguras vindas dos Secrets do GitHub Actions.

const {
  AZURE_TENANT_ID,
  AZURE_CLIENT_ID,
  AZURE_CLIENT_SECRET,
  EVOLUTION_API_URL,
  EVOLUTION_INSTANCE,
  EVOLUTION_API_KEY,
  WHATSAPP_NUMERO
} = process.env;

/* ---------------- CONFIGURAÇÕES DO SHAREPOINT ---------------- */
const SITE_DOMAIN = "willtech7.sharepoint.com";
const SITE_PATH = "/sites/CARGOVAN";
const LISTA_DESPESAS = "BD_DESPESAS";

// Link RAW direto da imagem no seu repositório GitHub
const URL_IMAGEM_RAW = "https://raw.githubusercontent.com/Willgner777/Cargo-van-html/main/imagens/banner_bot_tech_solutions.png";

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
    throw new Error(`[ERROR] Secrets ausentes no GitHub: ${faltando.join(", ")}`);
  }
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
    throw new Error(`[ERROR] Falha no login Azure (${res.status}): ${data.error_description || JSON.stringify(data)}`);
  }
  return data.access_token;
}

function diagnosticarToken(token) {
  const p = JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString());
  console.log(`[INFO] DIAGNÓSTICO TOKEN -> aud: ${p.aud} | tid: ${p.tid} | appid: ${p.appid} | roles: ${p.roles}`);
  if (!p.roles || !p.roles.length) {
    console.warn("[WARN] ATENÇÃO: token SEM roles. A permissão está como Delegada ou falta 'Conceder consentimento do administrador'.");
  }
}

async function graphGet(url, token) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`[ERROR] Graph ${res.status} em ${url}\n${await res.text()}`);
  return res.json();
}

async function obterSiteId(token) {
  try {
    const site = await graphGet(`https://graph.microsoft.com/v1.0/sites/${SITE_DOMAIN}:${SITE_PATH}`, token);
    return site.id;
  } catch (e) {
    console.warn(`[WARN] Consulta direta do site falhou: ${e.message.split("\n")[0]}`);
    console.log("[INFO] Tentando via busca (sites?search=CARGOVAN)...");
    const busca = await graphGet(`https://graph.microsoft.com/v1.0/sites?search=CARGOVAN`, token);
    const achado = (busca.value || []).find(s => (s.webUrl || "").toLowerCase().includes("/sites/cargovan"));
    if (!achado) throw new Error("[ERROR] Site CARGOVAN não encontrado pelo Graph. Verifique permissão Sites.Read.All (Aplicativo) + consentimento.");
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
      console.log(`[INFO] Evolution connectionState (${res.status}): ${body}`);
      if (res.ok) {
        if (!/"state"\s*:\s*"open"/.test(body)) {
          throw new Error("[ERROR] Instância NÃO está conectada ao WhatsApp. Reescaneie o QR Code no Evolution.");
        }
        return base;
      }
      if (res.status === 401 || res.status === 404) throw new Error(`[ERROR] Evolution recusou (${res.status}): ${body}`);
    } catch (e) {
      if (/NÃO está conectada|recusou/.test(e.message) || i === 4) throw e;
      console.warn(`[WARN] Tentativa ${i} falhou (${e.message}). Aguardando 20s (Render acordando)...`);
      await sleep(20_000);
    }
  }
}

/* ---------------- Envio Otimizado (Lotes e Concorrência Segura) ---------------- */
async function enviarWhatsApp(base, texto) {
  const numeros = limpar(WHATSAPP_NUMERO)
    .split(",")
    .map(n => n.trim())
    .filter(Boolean);

  if (!numeros.length) {
    throw new Error("[ERROR] Nenhum número válido foi encontrado na variável WHATSAPP_NUMERO.");
  }

  console.log(`[INFO] Disparando envio para ${numeros.length} destinatário(s)...`);

  // Define as opções padrão da requisição
  const requestOptions = (num) => ({
    method: "POST",
    headers: { 
      "Content-Type": "application/json", 
      apikey: limpar(EVOLUTION_API_KEY) 
    },
    body: JSON.stringify({
      number: num,
      media: URL_IMAGEM_RAW,
      mediatype: "image",
      caption: texto
    }),
    signal: AbortSignal.timeout(60_000)
  });

  // Cria um array de execuções com intervalos escalonados (Cascata)
  const promessasEnvio = numeros.map(async (num, index) => {
    // Escalonamento: O primeiro vai no tempo 0s, o segundo aos 3s, terceiro aos 6s, etc.
    await sleep(index * 3000);
    
    console.log(`[INFO] Processando envio para: ${num}`);
    const res = await fetch(`${base}/message/sendMedia/${limpar(EVOLUTION_INSTANCE)}`, requestOptions(num));
    
    const corpo = await res.text();
    if (!res.ok) {
      throw new Error(`Falha HTTP ${res.status}: ${corpo}`);
    }
    return num;
  });

  // Aguarda todos terminarem (mesmo se alguns falharem, o script não quebra)
  const resultados = await Promise.allSettled(promessasEnvio);

  // Exibe o relatório de disparos
  const sucessos = resultados.filter(r => r.status === "fulfilled");
  const falhas = resultados.filter(r => r.status === "rejected");

  console.log(`\n[INFO] --- RESUMO DO ENVIO ---`);
  console.log(`✅ Sucesso: ${sucessos.length}`);
  console.log(`❌ Falhas:  ${falhas.length}`);

  if (falhas.length > 0) {
    console.warn("\n[WARN] Detalhes das falhas de envio:");
    falhas.forEach((f, i) => console.warn(`   Erro ${i + 1}: ${f.reason.message}`));
  } else {
    console.log(`[INFO] Todos os relatórios foram enviados com sucesso!`);
  }
}

/* ---------------- Execução Principal ---------------- */
(async () => {
  try {
    exigirEnv();

    console.log("[INFO] 1/4 Login Azure...");
    const token = await getGraphAccessToken();
    diagnosticarToken(token);

    console.log("[INFO] 2/4 Lendo BD_DESPESAS...");
    const itens = await buscarDespesas(token);
    const r = contarStatus(itens);
    console.log(`[INFO] ${itens.length} itens | ${r.pendentes} pendentes, ${r.recusados} recusados, ${r.aprovados} aprovados`);

    // Captura a data e hora atual no padrão brasileiro
    const dataAtual = new Intl.DateTimeFormat('pt-BR', { dateStyle: 'short', timeStyle: 'short' }).format(new Date());

    const texto =
      `*RELATÓRIO DE DESPESAS* 🚛\n` +
      `Resumo diário — Cargo Van\n\n` +
      `📊 *Status das Solicitações:*\n` +
      `⏳ *${r.pendentes}* Pendente(s)\n` +
      `✅ *${r.aprovados}* Aprovada(s)\n` +
      `❌ *${r.recusados}* Recusada(s)\n\n` +
      `──────────\n` +
      `🤖 _By Tech Solutions Bot_\n` +
      `🕒 _Atualizado em: ${dataAtual}_`; // Data adicionada dinamicamente

    console.log("[INFO] 3/4 Verificando Evolution API...");
    const base = await prepararEvolution();

    console.log("[INFO] 4/4 Processando Disparos do WhatsApp...");
    await enviarWhatsApp(base, texto);
    
    console.log("[INFO] Execução finalizada com sucesso.");
  } catch (e) {
    console.error("\n[ERROR] FALHA CRÍTICA NA EXECUÇÃO:", e.message);
    process.exit(1);
  }
})();
