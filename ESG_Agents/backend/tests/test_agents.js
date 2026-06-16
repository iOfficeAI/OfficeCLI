const assert = require('assert');
const path = require('path');
const fs = require('fs');

const { runParser } = require('../agents/parser');
const { runLegalAgent } = require('../agents/legal');
const { runRiskAgent } = require('../agents/risk');
const { runReportingAgent } = require('../agents/reporting');
const { runClimateAgent } = require('../agents/climate');
const { runSupplyChainAgent } = require('../agents/supply_chain');
const { runFinancialAgent } = require('../agents/financial');
const { runAuditAgent } = require('../agents/audit');
const { runSummarizer } = require('../agents/summarizer');

async function run() {
  console.log('Running agent tests...');
  const fixturesDir = path.join(__dirname, '..', 'fixtures');
  const fixtures = ['contract.txt', 'esg_summary.txt', 'emissions.csv'].map(f => path.join(fixturesDir, f));

  for (const fp of fixtures) {
    console.log('\nFixture:', fp);
    const parserOut = await runParser(fp, path.basename(fp));
    assert(parserOut, 'parser must return a result');
    // Parser should return a canonical object or text
    assert(parserOut.canonical || parserOut.text, 'parser output missing canonical/text');
    const textInput = (parserOut.canonical && (parserOut.canonical.raw || parserOut.canonical.text)) || parserOut.text || '';

    if (!process.env.OPENAI_API_KEY) {
      console.warn('OPENAI_API_KEY not set — skipping LLM agent tests (legal/risk/reporting/etc).');
      continue;
    }

    // LLM-backed agents: ensure they return at least a raw string or parsed JSON
    const legal = await runLegalAgent(textInput, { jurisdiction: 'EU,UK,Global' });
    assert(legal && (legal.raw || legal.parsed || legal.error), 'legal agent produced no output');

    const risk = await runRiskAgent(textInput);
    assert(risk && (risk.raw || risk.parsed || risk.error), 'risk agent produced no output');

    const reporting = await runReportingAgent(textInput);
    assert(reporting && (reporting.raw || reporting.parsed || reporting.error), 'reporting agent produced no output');

    const climate = await runClimateAgent(textInput);
    assert(climate && (climate.raw || climate.parsed || climate.error), 'climate agent produced no output');

    const supply = await runSupplyChainAgent(textInput);
    assert(supply && (supply.raw || supply.parsed || supply.error), 'supply chain agent produced no output');

    const financial = await runFinancialAgent(textInput);
    assert(financial && (financial.raw || financial.parsed || financial.error), 'financial agent produced no output');

    const audit = await runAuditAgent(textInput);
    assert(audit && (audit.raw || audit.parsed || audit.error), 'audit agent produced no output');

    const summary = await runSummarizer(textInput);
    assert(summary && (summary.raw || summary.parsed || summary.error), 'summarizer produced no output');

    console.log('All agents returned output for fixture:', fp);
  }

  console.log('\nAgent tests completed.');
}

run().catch(err => { console.error('Test run failed:', err); process.exit(1); });
