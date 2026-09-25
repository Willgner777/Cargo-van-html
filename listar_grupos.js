// listar_grupos.js — descobre o ID (@g.us) de cada grupo da instância Evolution API

const {
  EVOLUTION_API_URL,
  EVOLUTION_INSTANCE,
  EVOLUTION_API_KEY
} = process.env;

const limpar = (v) => String(v ?? "").trim();

(async () => {
  try {
    const base = limpar(EVOLUTION_API_URL).replace(/\/$/, "");

    const response = await fetch(
      `${base}/group/fetchAllGroups/${limpar(EVOLUTION_INSTANCE)}?getParticipants=false`,
      {
        method: "GET",
        headers: {
          apikey: limpar(EVOLUTION_API_KEY)
        }
      }
    );

    const grupos = await response.json();

    if (!response.ok) {
      throw new Error(`Erro ${response.status}: ${JSON.stringify(grupos)}`);
    }

    console.log(`\nEncontrados ${grupos.length} grupo(s):\n`);

    for (const g of grupos) {
      console.log(`Nome: ${g.subject}`);
      console.log(`ID:   ${g.id}`);
      console.log("---");
    }

  } catch (e) {
    console.error("Erro:", e);
    process.exit(1);
  }
})();
