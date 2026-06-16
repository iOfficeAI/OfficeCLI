const llm = require('../llm');

async function runSummarizer(inputText, opts = {}) {
  const prompt = `You are an executive summarizer. Produce an executive brief: one-paragraph summary, top 5 findings, and 3 recommended next steps. Return JSON: {summary:'', findings:[], next_steps:[]} `;
  const messages = [ { role: 'system', content: 'You are a concise summarizer.' }, { role: 'user', content: prompt + '\n\nDocument:\n' + inputText } ];
  try {
    const resp = await llm.chatCompletion({ model: process.env.OPENAI_MODEL || 'gpt-4o-mini', messages, max_tokens: 500, temperature: 0 });
    const text = resp.choices?.[0]?.message?.content ?? '';
    let parsed = null;
    try { const m = text.match(/\{[\s\S]*\}/); if (m) parsed = JSON.parse(m[0]); } catch (e) { parsed = { raw: text }; }
    return { raw: text, parsed };
  } catch (err) { return { error: String(err) }; }
}
module.exports = { runSummarizer };
