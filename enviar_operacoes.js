// enviar_operacoes.js — Relatório de Operações (LIBERACAO_VEICULO)

const {
  AZURE_TENANT_ID,
  AZURE_CLIENT_ID,
  AZURE_CLIENT_SECRET,
  EVOLUTION_API_URL,
  EVOLUTION_INSTANCE,
  EVOLUTION_API_KEY,
  WHATSAPP_NUMERO
} = process.env;

const SITE_DOMAIN = "willtech7.sharepoint.com";
const SITE_PATH = "/sites/CARGOVAN";
const NOME_LISTA = "LIBERACAO_VEICULO";

// >>> Ajuste aqui se o nome interno do campo for diferente <<<
const CAMPO_CLIENTE_OPERACAO = "CLIENTE_x007c_OPERA_x00c7__x00c3"; // Coluna "CLIENTE | OPERAÇÃO"
const CAMPO_MOTORISTA = "NOME"; // Coluna "Motorista"
// A data usada no filtro é a coluna padrão "Criado" do SharePoint (metadado createdDateTime do item)

const sleep = (ms) => new Promise(r => setTimeout(r, ms));
const limpar = (v) => String(v ?? "").trim();

// "JOÃO FRANCISCO" -> "João Francisco"
function titleCase(nome) {
  return limpar(nome)
    .toLowerCase()
    .split(/\s+/)
    .filter(Boolean)
    .map(p => p.charAt(0).toUpperCase() + p.slice(1))
    .join(" ");
}

async function getGraphAccessToken() {
  const res = await fetch(
    `https://login.microsoftonline.com/${limpar(AZURE_TENANT_ID)}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: limpar(AZURE_CLIENT_ID),
        client_secret: limpar(AZURE_CLIENT_SECRET),
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials"
      })
    }
  );

  const data = await res.json();

  if (!res.ok || !data.access_token) {
    throw new Error(
      `Erro Azure: ${data.error_description || JSON.stringify(data)}`
    );
  }

  return data.access_token;
}

async function graphGet(url, token) {
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });

  if (!res.ok) {
    throw new Error(`Erro Graph ${res.status}: ${await res.text()}`);
  }

  return res.json();
}

async function obterSiteId(token) {
  const site = await graphGet(
    `https://graph.microsoft.com/v1.0/sites/${SITE_DOMAIN}:${SITE_PATH}`,
    token
  );

  return site.id;
}

async function buscarItens(token) {
  const siteId = await obterSiteId(token);

  let url =
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${NOME_LISTA}/items?expand=fields&$top=1000`;

  const itens = [];

  while (url) {
    const data = await graphGet(url, token);

    itens.push(...(data.value || []));

    url = data["@odata.nextLink"] || null;
  }

  return itens;
}

// Verifica se a data de criação do item (campo padrão "Criado" / createdDateTime) é hoje,
// comparando apenas dd/mm/aaaa no fuso America/Sao_Paulo (hora é ignorada)
function isHojeSaoPaulo(createdDateTime) {
  if (!createdDateTime) return false;

  const d = new Date(createdDateTime);
  if (isNaN(d.getTime())) return false;

  const formatador = new Intl.DateTimeFormat("en-CA", { timeZone: "America/Sao_Paulo" }); // yyyy-mm-dd
  return formatador.format(d) === formatador.format(new Date());
}

(async () => {
  try {
    const token = await getGraphAccessToken();

    const itens = await buscarItens(token);

    // Filtra somente os itens de hoje (não acumulativo), usando o campo padrão "Criado"
    const itensHoje = itens.filter(it => isHojeSaoPaulo(it.createdDateTime));

    let emAndamento = 0;
    let pernoite = 0;
    let concluidos = 0;
    const porClienteOperacao = {};
    const motoristasPorGrupo = {}; // chave -> Set de nomes (Title Case)

    for (const it of itensHoje) {
      const f = it.fields || {};

      const st = limpar(
        f.STATUS ?? f.Status ?? ""
      ).toUpperCase();

      if (st.includes("ANDAMENTO")) {
        emAndamento++;
      } else if (st.includes("PERNOITE")) {
        pernoite++;
      } else if (
        st.includes("CONCLUÍDO") ||
        st.includes("CONCLUIDO")
      ) {
        concluidos++;
      }

      const clienteOperacao = limpar(f[CAMPO_CLIENTE_OPERACAO]);
      if (clienteOperacao) {
        porClienteOperacao[clienteOperacao] = (porClienteOperacao[clienteOperacao] || 0) + 1;

        const nomeMotorista = titleCase(f[CAMPO_MOTORISTA]);
        if (nomeMotorista) {
          if (!motoristasPorGrupo[clienteOperacao]) {
            motoristasPorGrupo[clienteOperacao] = new Set();
          }
          motoristasPorGrupo[clienteOperacao].add(nomeMotorista);
        }
      }
    }

    let blocoClientes = "";
    for (const [chave, qtd] of Object.entries(porClienteOperacao)) {
      const motoristas = Array.from(motoristasPorGrupo[chave] || []).join(", ");

      blocoClientes += `> 📍 ${chave}\n> ▫️ Qtd: ${qtd}\n`;
      if (motoristas) {
        blocoClientes += `> ▫️ Motoristas: ${motoristas}\n`;
      }
      blocoClientes += `\n`;
    }
    if (!blocoClientes) {
      blocoClientes = "> Nenhuma operação registrada hoje.\n\n";
    }

    const dataAtual = new Intl.DateTimeFormat(
      "pt-BR",
      {
        dateStyle: "short",
        timeStyle: "short",
        timeZone: "America/Sao_Paulo"
      }
    ).format(new Date());

    const texto =
      `🚚 *ACOMPANHAMENTO DE OPERAÇÕES* — CARGO VAN EX\n\n` +
      `> 📋 *Clientes em Operação:*\n` +
      blocoClientes +
      `> ───────────────\n` +
      `⠀\n` +
      `> 📊 *Status dos Veículos e Viagens:*\n` +
      `> 🔄 Em Andamento: ${emAndamento}\n` +
      `> 🌙 Pernoite: ${pernoite}\n` +
      `> ✅ Concluído: ${concluidos}\n\n` +
      `🤖 By Tech Solutions Bot\n` +
      `🕒 Atualizado em: ${dataAtual}`;

    const base = limpar(EVOLUTION_API_URL).replace(/\/$/, "");

    const numeros = limpar(WHATSAPP_NUMERO)
      .split(",")
      .map(n => n.trim())
      .filter(Boolean);

    for (const num of numeros) {

      const response = await fetch(
        `${base}/message/sendText/${limpar(EVOLUTION_INSTANCE)}`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            apikey: limpar(EVOLUTION_API_KEY)
          },
          body: JSON.stringify({
            number: num,
            text: texto
          })
        }
      );

      const resultado = await response.text();

      console.log(`Número: ${num}`);
      console.log(`Status: ${response.status}`);
      console.log(`Resposta: ${resultado}`);

      if (!response.ok) {
        throw new Error(
          `Falha ao enviar para ${num}: ${resultado}`
        );
      }

      await sleep(3000);
    }

    console.log("Relatório de Operações enviado com sucesso.");

  } catch (e) {
    console.error("Erro:", e);
    process.exit(1);
  }
})();
