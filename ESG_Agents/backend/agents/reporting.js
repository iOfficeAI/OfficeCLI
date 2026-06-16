const llm = require('../llm');

async function runReportingAgent(inputText, opts = {}) {
  const taxonomies = opts.taxonomies || ['GRI','SASB','EU_Taxonomy'];
  const prompt = `You are an ESG reporting specialist. Map the document's disclosures to common ESG taxonomies (${taxonomies.join(', ')}). Identify disclosure gaps and produce a recommended disclosure checklist. Return JSON with keys: mappings: [{taxonomy, code, description, matched_text}], gaps: [{recommendation, priority}], recommendations: [].`;
  const messages = [
    { role: 'system', content: 'You are a helpful ESG reporting expert.' },
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

module.exports = { runReportingAgent };
