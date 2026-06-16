const OpenAI = require('openai');

async function realChatCompletion(opts) {
  const client = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
  // keep compatibility: return object with choices[0].message.content
  const resp = await client.chat.completions.create({
    model: opts.model,
    messages: opts.messages,
    max_tokens: opts.max_tokens,
    temperature: opts.temperature ?? 0.0
  });
  return resp;
}

function cannedResponseFor(roleHint) {
  // roleHint: string indicating which agent
  switch (roleHint) {
    case 'legal':
      return JSON.stringify({ issues: [], citations: [], recommendations: ["Review supplier clauses"], confidence: 75 });
    case 'risk':
      return JSON.stringify({ risks: [{ id: 'R1', title: 'Supply concentration', description: 'High concentration in one supplier', likelihood: 'medium', impact: 'high', estimated_risk_score: 70, remediation_recommendation: 'Diversify suppliers' }], summary: 'Operational and supply risks present.' });
    case 'reporting':
      return JSON.stringify({ mappings: [], gaps: [{ recommendation: 'Add Scope 3 breakdown', priority: 'high' }], recommendations: ['Map disclosures to GRI'] });
    case 'climate':
      return JSON.stringify({ emissions_summary: { scope1: null, scope2: null, scope3: null }, missing_data: ['Scope 3 breakdown'], recommendations: ['Start supplier data collection'], confidence: 60 });
    case 'supply_chain':
      return JSON.stringify({ findings: ['Limited traceability'], concentrations: ['Supplier A: 40% spend'], recommendations: ['Supplier engagement'] });
    case 'financial':
      return JSON.stringify({ impacts: ['CapEx for remediation'], estimates: [], recommendations: ['Quantify CAPEX needs'] });
    case 'audit':
      return JSON.stringify({ evidence: [], missing: ['transaction-level records'], checklist: ['Collect invoices', 'Collect emission factors'] });
    case 'summarizer':
      return JSON.stringify({ summary: 'Executive summary placeholder', findings: ['Key finding 1'], next_steps: ['Engage stakeholders'] });
    default:
      return JSON.stringify({ raw: 'mocked response' });
  }
}

async function chatCompletion(opts) {
  // opts.messages: array of messages; try to detect agent role
  if (!process.env.OPENAI_API_KEY) {
    // detect role from system or user messages
    const msgs = opts.messages || [];
    let roleHint = null;
    for (const m of msgs) {
      const c = (m.content || '').toLowerCase();
      if (c.includes('legal')) { roleHint = 'legal'; break; }
      if (c.includes('risk')) { roleHint = 'risk'; break; }
      if (c.includes('reporting')) { roleHint = 'reporting'; break; }
      if (c.includes('climate')) { roleHint = 'climate'; break; }
      if (c.includes('supply chain') || c.includes('supply')) { roleHint = 'supply_chain'; break; }
      if (c.includes('financial')) { roleHint = 'financial'; break; }
      if (c.includes('audit') || c.includes('assurance')) { roleHint = 'audit'; break; }
      if (c.includes('summariz')) { roleHint = 'summarizer'; break; }
    }
    // fallback to user message keywords
    if (!roleHint && msgs.length > 0) {
      const user = msgs.find(m => m.role === 'user');
      if (user) {
        const cu = (user.content || '').toLowerCase();
        if (cu.includes('regulatory') || cu.includes('jurisdiction')) roleHint = 'legal';
        else if (cu.includes('emissions') || cu.includes('scope')) roleHint = 'climate';
        else if (cu.includes('risk')) roleHint = 'risk';
      }
    }
    const key = roleHint === 'supply_chain' ? 'supply_chain' : (roleHint || 'default');
    const content = cannedResponseFor(key === 'supply_chain' ? 'supply_chain' : (roleHint || 'default'));
    return { choices: [ { message: { content } } ] };
  }

  return await realChatCompletion(opts);
}

module.exports = { chatCompletion };
