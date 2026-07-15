// Généré automatiquement par maxit_tracker.py - basé sur app-store-scraper.
const store = require('app-store-scraper');

// app-store-scraper convertit le code pays en identifiant interne de
// vitrine Apple (technologie héritée d'iTunes) pour aller chercher un
// classement. Si le pays n'est pas dans SA table interne, la librairie
// bascule SILENCIEUSEMENT sur le classement des Etats-Unis, sans erreur.
// On récupère cette même table pour détecter ce cas et éviter un résultat
// trompeur (un "None" qui aurait en fait vérifié le mauvais pays).
let knownMarkets = {};
try {
  knownMarkets = require('app-store-scraper/lib/constants').markets;
} catch (e) {
  // Si jamais la structure interne de la librairie change, on continue
  // sans cette vérification plutôt que de faire planter tout le script.
}

async function main() {
  const [appIdArg, country, categoryHint, rankSearchDepthArg] = process.argv.slice(2);
  const rankSearchDepth = parseInt(rankSearchDepthArg, 10) || 200;
  const appIdNum = parseInt(appIdArg, 10);
  const result = { app_id: appIdArg, country: country };

  try {
    const details = await store.app({ id: appIdNum, country });
    result.app_name = details.title ?? null;
    result.developer = details.developer ?? null;
    result.category = details.primaryGenre ?? categoryHint ?? null;
    result.price = details.price ?? 0;
    result.is_free = details.free ?? (details.price === 0);
    result.rating_avg = details.score ?? null;
    result.rating_count = details.reviews ?? null;
    result.installs = null; // Apple ne publie jamais ce chiffre
    result.version = details.version ?? null;
    result.description_short = (details.description ?? '').slice(0, 300);
  } catch (err) {
    result.error = `store.app a échoué: ${err.message}`;
    process.stdout.write(JSON.stringify(result));
    return;
  }

  const countryIsSupportedForRanking = knownMarkets[country.toUpperCase()] !== undefined;

  if (!countryIsSupportedForRanking) {
    // app-store-scraper ne sait pas interroger ce pays (table interne
    // obsolète), mais Apple a bien une vraie vitrine pour ce pays - on
    // scrape donc directement la page web publique des classements.
    // Limite : cette page ne montre que le TOP 25 (contre top 200 pour
    // les pays supportés nativement par la librairie), donc si Max it
    // est classé au-dela de la 25e place, "rank" restera vide ici même
    // si un rang reel existe plus loin.
    try {
      const chartUrl = `https://apps.apple.com/${country.toLowerCase()}/iphone/charts/36?chart=top-free`;
      const res = await fetch(chartUrl, {
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' },
      });
      const html = await res.text();

      // Extrait les ID d'app dans leur ordre d'apparition dans la page
      // (qui correspond a l'ordre du classement).
      const idMatches = [...html.matchAll(/\/id(\d+)"/g)].map((m) => m[1]);
      const idsUniquesOrdonnes = [...new Set(idMatches)];

      const idx = idsUniquesOrdonnes.findIndex((id) => Number(id) === appIdNum);
      if (idx === -1) {
        result.rank = null;
        result.rank_note = `Non trouvé dans le top 25 web (méthode de repli limitée à 25, pays "${country}" non couvert par app-store-scraper).`;
      } else {
        result.rank = idx + 1;
        result.rank_note = 'Rang obtenu via la page web des classements (méthode de repli, top 25 max), pas via app-store-scraper.';
      }
    } catch (err) {
      result.rank = null;
      result.rank_note = `Échec du scraping de repli pour "${country}": ${err.message}`;
    }
    process.stdout.write(JSON.stringify(result));
    return;
  }

  try {
    const topList = await store.list({
      collection: store.collection.TOP_FREE_IOS,
      // Pas de filtre "category" ici : on veut le classement toutes
      // catégories confondues (comme sur Android avec 'APPLICATION'),
      // pas seulement la catégorie Finance.
      country,
      num: rankSearchDepth,
    });
    const idx = topList.findIndex((app) => Number(app.id) === appIdNum);
    result.rank = idx === -1 ? null : idx + 1;
  } catch (err) {
    result.rank = null;
    result.rank_error = `store.list a échoué: ${err.message}`;
  }

  process.stdout.write(JSON.stringify(result));
}

main().catch((err) => {
  process.stdout.write(JSON.stringify({ error: `Erreur inattendue: ${err.message}` }));
});
