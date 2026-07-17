// Généré automatiquement par maxit_tracker.py - basé sur parse-play.
const { fetchAppDetails, fetchTopCharts } = require('parse-play');

async function main() {
  const [appId, country, categoryHint, rankSearchDepthArg] = process.argv.slice(2);
  const rankSearchDepth = parseInt(rankSearchDepthArg, 10) || 200;
  const result = { app_id: appId, country: country };

  try {
    const details = await fetchAppDetails({ appId }, { language: 'EN', country });
    result.app_name = details.name ?? null;
    result.developer = details.developer?.name ?? details.developer ?? null;
    result.category = details.category ?? categoryHint ?? null;
    result.price = details.price ?? 0;
    result.is_free = (details.price ?? 0) === 0;
    result.rating_avg = details.aggregateRating?.ratingValue ?? details.rating ?? null;
    result.rating_count = details.rating_counts?.total ?? null;
    result.installs = details.downloads ?? null;
    result.version = details.version ?? null;
    result.description_short = (details.description ?? '').slice(0, 300);
  } catch (err) {
    // DEBUG: on inclut la stack complète pour diagnostiquer (proxy/réseau vs
    // vraie erreur de parsing). A retirer une fois le problème identifié.
    result.error = `fetchAppDetails a échoué: ${err.message}`;
    result.error_stack = err.stack;
    process.stdout.write(JSON.stringify(result));
    return;
  }

  try {
    const topChart = await fetchTopCharts(
      { category: 'APPLICATION', chart: 'topselling_free', count: rankSearchDepth },
      { country, language: 'EN' }
    );
    const idx = topChart.findIndex((app) => app.app_id === appId);
    result.rank = idx === -1 ? null : idx + 1;
  } catch (err) {
    result.rank = null;
    result.rank_error = `fetchTopCharts a échoué: ${err.message}`;
  }

  process.stdout.write(JSON.stringify(result));
}

main().catch((err) => {
  process.stdout.write(JSON.stringify({ error: `Erreur inattendue: ${err.message}` }));
});
