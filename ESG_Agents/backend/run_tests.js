// Simple test runner to exercise agents against fixtures
const { runParser } = require('./agents/parser');
const { runLegalAgent } = require('./agents/legal');
const { runRiskAgent } = require('./agents/risk');
const { runReportingAgent } = require('./agents/reporting');
const { runClimateAgent } = require('./agents/climate');
const { runSupplyChainAgent } = require('./agents/supply_chain');
const { runFinancialAgent } = require('./agents/financial');
const { runAuditAgent } = require('./agents/audit');
const { runSummarizer } = require('./agents/summarizer');
const path = require('path');

async function main() {
  const fixtures = [
    'fixtures/contract.txt',
    'fixtures/emissions.csv',
    'fixtures/esg_summary.txt'
  ];

  for (const f of fixtures) {
    const fp = path.join(__dirname, f);
    console.log(`\n=== Fixture: ${f} ===`);
    const p = await runParser(fp, path.basename(fp));
    console.log('Parser canonical:', JSON.stringify(p.canonical || p.text || p, null, 2));

    const textInput = (p.canonical && (p.canonical.raw || p.canonical.text)) || p.text || '';
    console.log('Running legal...');
    console.log(await runLegalAgent(textInput, { jurisdiction: 'EU,UK,Global' }));
    console.log('Running risk...');
    console.log(await runRiskAgent(textInput));
    console.log('Running reporting...');
    console.log(await runReportingAgent(textInput));
    console.log('Running climate...');
    console.log(await runClimateAgent(textInput));
    console.log('Running supply_chain...');
    console.log(await runSupplyChainAgent(textInput));
    console.log('Running financial...');
    console.log(await runFinancialAgent(textInput));
    console.log('Running audit...');
    console.log(await runAuditAgent(textInput));
    console.log('Running summarizer...');
    console.log(await runSummarizer(textInput));
  }
}

main().catch(e => { console.error(e); process.exit(1); });
