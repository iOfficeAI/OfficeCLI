const llm = require('../llm');

async function runLegalAgent(inputText, opts = {}) {
  // Simple prompt-template for the legal compliance lens. In production
  // this should be replaced by a more structured chain-of-thought and
  // verification steps, with citation extraction and statutory lookup.
  const jurisdiction = opts.jurisdiction || 'Global';
  const prompt = `You are a legal compliance analyst specialized in ESG regulations.\n\nJurisdictions: ${jurisdiction}\n\nAnalyze the following document text for regulatory and legal implications, list potential compliance issues, likely affected laws/regulations, recommended next steps, and an estimated confidence score (0-100). Provide structured JSON with keys: issues[], citations[], recommendations[], confidence.`;

  const messages = [
    { role: 'system', content: 'You are a helpful legal analyst.' },
    { role: 'user', content: prompt + '\n\nDocument:\n' + inputText }
  ];

  try {
    const resp = await llm.chatCompletion({ model: process.env.OPENAI_MODEL || 'gpt-4o-mini', messages, max_tokens: 800, temperature: 0.0 });
    const text = resp.choices?.[0]?.message?.content ?? '';
    // Try to parse JSON from the model output
    let parsed = null;
    try {
      // Find first { .. } block
      const m = text.match(/\{[\s\S]*\}/);
      if (m) parsed = JSON.parse(m[0]);
    } catch (e) { parsed = { raw: text }; }
    return { raw: text, parsed };
  } catch (err) {
    return { error: String(err) };
  }
}

module.exports = { runLegalAgent };
