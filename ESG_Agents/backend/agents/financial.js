const llm = require('../llm');

async function runFinancialAgent(inputText, opts = {}) {
  const prompt = `You are a financial analyst assessing ESG-related financial impacts. Identify potential costs, liabilities, capital expenditure needs, and implications for revenue or margins. Return JSON with keys: impacts:[], estimates:[], recommendations:[]`;
  const messages = [ { role: 'system', content: 'You are a financial analyst.' }, { role: 'user', content: prompt + '\n\nDocument:\n' + inputText } ];
  try {
    const resp = await llm.chatCompletion({ model: process.env.OPENAI_MODEL || 'gpt-4o-mini', messages, max_tokens: 700, temperature: 0 });
    const text = resp.choices?.[0]?.message?.content ?? '';
    let parsed = null;
    try { const m = text.match(/\{[\s\S]*\}/); if (m) parsed = JSON.parse(m[0]); } catch (e) { parsed = { raw: text }; }
    return { raw: text, parsed };
  } catch (err) { return { error: String(err) }; }
}
module.exports = { runFinancialAgent };
