const API_URL = "https://evolution-api-yga4.onrender.com";
const INSTANCE_NAME = "AutomacaoCargoVanv1";
const API_KEY = "SuaChaveSeguraAqui9405";
const NUMERO_DESTINO = "5585994050393";

async function dispararResumo() {
  // Exemplo de dados (substitua depois pelas variáveis dinâmicas do SharePoint)
  const pendentes = 2;
  const recusados = 2;
  const aprovados = 11;

  const mensagem = `*Resumo Diário de Despesas - Cargo Van* 🚛\n\n` +
    `📌 *${pendentes}* itens pendentes para aprovação\n` +
    `❌ *${recusados}* itens recusados\n` +
    `✅ *${aprovados}* itens aprovados\n\n` +
    `Para aprovar ou analisar, aceda ao painel do sistema da Cargo Van e navegue até à aba *Aprovar Despesas*.`;

  try {
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
    console.error("Erro ao enviar mensagem:", error);
    process.exit(1);
  }
}

dispararResumo();
