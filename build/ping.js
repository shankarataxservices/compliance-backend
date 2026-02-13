// netlify/functions/ping.js
exports.handler = async () => {
  return {
    statusCode: 200,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ok: true, msg: "ping works" })
  };
};
