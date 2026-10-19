// listar_grupos.js — Lista os grupos do WhatsApp conectados na Evolution API

const {
  EVOLUTION_API_URL,
  EVOLUTION_INSTANCE,
  EVOLUTION_API_KEY
} = process.env;

const limpar = (v) => String(v ?? "").trim();

(async () => {
  try {
    console.log("--- Status das Variáveis para Listar Grupos ---");
    console.log("EVOLUTION_API_URL:", EVOLUTION_API_URL ? "Definida" : "⚠️ Vazia");
    console.log("EVOLUTION_INSTANCE:", EVOLUTION_INSTANCE ? "Definida" : "⚠️ Vazia");
    console.log("EVOLUTION_API_KEY:", EVOLUTION_API_KEY ? "Definida" : "⚠️ Vazia");
    console.log("-----------------------------------------------");

    const base = limpar(EVOLUTION_API_URL).replace(/\/$/, "");
    const instancia = limpar(EVOLUTION_INSTANCE);
    const apiKey = limpar(EVOLUTION_API_KEY);

    // Rota corrigida para a versão atual da Evolution API
    const url = `${base}/chat/fetchChats/${instancia}`;
    
    console.log(`Buscando chats em: ${base}/chat/fetchChats/...`);

    const response = await fetch(url, {
      method: "GET",
      headers: {
        "Content-Type": "application/json",
        apikey: apiKey
      }
    });

    const resultado = await response.text();

    if (!response.ok) {
      throw new Error(`Erro ${response.status}: ${resultado}`);
    }

    const chats = JSON.parse(resultado);
    console.log("\n--- Chats e Grupos Encontrados ---");
    
    const grupos = Array.isArray(chats) ? chats.filter(c => c.id && c.id.endsWith("@g.us")) : [];
    
    if (grupos.length === 0) {
      console.log("Nenhum grupo encontrado ou formato de resposta:", resultado);
    } else {
      grupos.forEach(g => {
        console.log(`Nome: ${g.name || "Sem Nome"} | ID: ${g.id}`);
      });
    }

  } catch (e) {
    console.error("Erro fatal ao listar grupos:", e);
    process.exit(1);
  }
})();
