const llm = require('../llm');

async function runSupplyChainAgent(inputText, opts = {}) {
  const prompt = `You are a supply chain analyst. Identify upstream supplier risks, traceability gaps, concentrations, and possible Scope 3 exposure. Return JSON with keys: findings:[], concentrations:[], recommendations:[]`;
  const messages = [
    { role: 'system', content: 'You are a supply chain risk analyst.' },
    { role: 'user', content: prompt + '\n\nDocument:\n' + inputText }
  ];
  try {
    const resp = await llm.chatCompletion({ model: process.env.OPENAI_MODEL || 'gpt-4o-mini', messages, max_tokens: 700, temperature: 0 });
    const text = resp.choices?.[0]?.message?.content ?? '';
    let parsed = null;
    try { const m = text.match(/\{[\s\S]*\}/); if (m) parsed = JSON.parse(m[0]); } catch (e) { parsed = { raw: text }; }
    return { raw: text, parsed };
  } catch (err) { return { error: String(err) }; }
}
module.exports = { runSupplyChainAgent };
