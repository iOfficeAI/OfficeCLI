const llm = require('../llm');

async function runRiskAgent(inputText, opts = {}) {
  const prompt = `You are an ESG risk analyst. Given the document text below, identify operational, financial, and reputational risks related to ESG. For each risk provide: id, title, description, likelihood (low/medium/high), impact (low/medium/high), estimated_risk_score (0-100), and remediation_recommendation. Return VALID JSON with keys: risks:[] and summary.`;
  const messages = [
    { role: 'system', content: 'You are a senior ESG risk analyst.' },
    { role: 'user', content: prompt + "\n\nDocument:\n" + inputText }
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

module.exports = { runRiskAgent };
