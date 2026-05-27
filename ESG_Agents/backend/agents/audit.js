const llm = require('../llm');

async function runAuditAgent(inputText, opts = {}) {
  const prompt = `You are an assurance/audit specialist. Identify what evidence in the document supports disclosures, note missing evidence, and suggest an evidence-gathering checklist for auditability. Return JSON: {evidence:[], missing:[], checklist:[]} `;
  const messages = [ { role: 'system', content: 'You are an audit and assurance expert.' }, { role: 'user', content: prompt + '\n\nDocument:\n' + inputText } ];
  try {
    const resp = await llm.chatCompletion({ model: process.env.OPENAI_MODEL || 'gpt-4o-mini', messages, max_tokens: 700, temperature: 0 });
    const text = resp.choices?.[0]?.message?.content ?? '';
    let parsed = null;
    try { const m = text.match(/\{[\s\S]*\}/); if (m) parsed = JSON.parse(m[0]); } catch (e) { parsed = { raw: text }; }
    return { raw: text, parsed };
  } catch (err) { return { error: String(err) }; }
}
module.exports = { runAuditAgent };
