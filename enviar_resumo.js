const API_URL = "https://evolution-api-yga4.onrender.com";
const INSTANCE_NAME = "AutomacaoCargoVanv1";
const API_KEY = "SuaChaveSeguraAqui9405";
const NUMERO_DESTINO = "5585994050393";

// Configurações do SharePoint via Graph API
const SITE_DOMAIN = "willtech7.sharepoint.com";
const SITE_PATH = "/sites/CARGOVAN";
const LISTA_DESPESAS = "BD_DESPESAS";

/**
 * Obtém o Token de Acesso da Microsoft Graph API usando OAuth2 Client Credentials
 */
async function getGraphAccessToken() {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Variáveis de ambiente do Azure/Graph não foram configuradas nos Secrets do GitHub.");
  }

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const params = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials"
  });

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: params
  });

  if (!res.ok) {
    const errorText = await res.text();
    throw new Error(`Erro ao obter token Azure: ${errorText}`);
  }

  const data = await res.json();
  return data.access_token;
}

/**
 * Consulta os itens da lista BD_DESPESAS no SharePoint via Graph API
 */
async function buscarDadosDespesas(accessToken) {
  // 1. Obter o ID do Site
  const siteUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_DOMAIN}:${SITE_PATH}`;
  const siteRes = await fetch(siteUrl, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  if (!siteRes.ok) throw new Error("Erro ao obter ID do Site do SharePoint.");
  const siteData = await siteRes.json();
  const siteId = siteData.id;

  // 2. Buscar itens da lista BD_DESPESAS
  const listUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${LISTA_DESPESAS}/items?expand=fields&$top=1000`;
  const listRes = await fetch(listUrl, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  if (!listRes.ok) throw new Error("Erro ao buscar dados da lista BD_DESPESAS.");
  
  const listData = await listRes.json();
  return listData.value || [];
}

/**
 * Processa os itens e dispara a mensagem via Evolution API
 */
async function dispararResumo() {
  try {
    console.log("Iniciando busca de dados reais no SharePoint...");
    const accessToken = await getGraphAccessToken();
    const items = await buscarDadosDespesas(accessToken);

    // Contadores de Status
    let pendentes = 0;
    let recusados = 0;
    let aprovados = 0;

    items.forEach(it => {
      const f = it.fields || {};
      const status = String(f.STATUS || f.Status || "Pendente").toLowerCase().trim();

      if (status.includes("pendente")) {
        pendentes++;
      } else if (status.includes("recusado") || status.includes("rejeitado")) {
        recusados++;
      } else if (status.includes("aprovado")) {
        aprovados++;
      }
    });

    console.log(`Dados processados: ${pendentes} pendentes, ${recusados} recusados, ${aprovados} aprovados.`);

    const mensagem = `*Resumo Diário de Despesas - Cargo Van* 🚛\n\n` +
      `📌 *${pendentes}* itens pendentes para aprovação\n` +
      `❌ *${recusados}* itens recusados\n` +
      `✅ *${aprovados}* itens aprovados\n\n` +
      `Para aprovar ou analisar, aceda ao painel do sistema da Cargo Van e navegue até à aba *Aprovar Despesas*.`;

    const response = await fetch(`${API_URL}/message/sendText/${INSTANCE_NAME}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'apikey': API_KEY
      },
      body: JSON.stringify({
        number: NUMERO_DESTINO,
        text: mensagem
      })
    });

    const result = await response.json();
    console.log("Notificação enviada com sucesso:", JSON.stringify(result, null, 2));
  } catch (error) {
    console.error("Erro na execução do script:", error);
    process.exit(1);
  }
}

dispararResumo();
