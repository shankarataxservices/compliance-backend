const { withCors, json } = require('./_common');

// Kept only to avoid confusion / older calls.
// Use jobs_daily5am for actual scheduling.
exports.handler = withCors(async (event) => {
  return json(event, 200, { ok:true, note:'Use /.netlify/functions/jobs_daily5am (this one is a no-op alias).' });
});