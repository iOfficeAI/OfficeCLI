const llm = require('../llm');

async function runClimateAgent(inputText, opts = {}) {
  const prompt = `You are a climate analyst. Extract any emissions data, assess whether Scope 1/2/3 are described, estimate emissions intensity where possible, flag missing data, and recommend mitigation actions and targets. Return JSON with keys: emissions_summary, missing_data, recommendations, confidence.`;
  const messages = [
    { role: 'system', content: 'You are a climate scientist and emissions analyst.' },
    { role: 'user', content: prompt + '\n\nDocument:\n' + inputText }
  ];

  try {
    const resp = await llm.chatCompletion({ model: process.env.OPENAI_MODEL || 'gpt-4o-mini', messages, max_tokens: 800, temperature: 0.0 });
    const text = resp.choices?.[0]?.message?.content ?? '';
    let parsed = null;
    try {
      const m = text.match(/\{[\s\S]*\}/);
      if (m) parsed = JSON.parse(m[0]);
    } catch (e) {
      parsed = { raw: text };
    }
    return { raw: text, parsed };
  } catch (err) {
    return { error: String(err) };
  }
}

module.exports = { runClimateAgent };
