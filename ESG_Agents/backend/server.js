const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { runParser } = require('./agents/parser');
const { runLegalAgent } = require('./agents/legal');
const { runRiskAgent } = require('./agents/risk');
const { runReportingAgent } = require('./agents/reporting');
const { runClimateAgent } = require('./agents/climate');
const { runSupplyChainAgent } = require('./agents/supply_chain');
const { runFinancialAgent } = require('./agents/financial');
const { runAuditAgent } = require('./agents/audit');
const { runSummarizer } = require('./agents/summarizer');

const UPLOAD_DIR = path.join(__dirname, 'uploads');
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

const upload = multer({ dest: UPLOAD_DIR });
const app = express();
app.use(express.json());

// Health endpoint for readiness checks
app.get('/health', (req, res) => {
  res.json({ status: 'ok', uptime: process.uptime(), timestamp: Date.now() });
});

// Detailed startup/readiness info
const startup = require('./startup');
app.get('/startup', async (req, res) => {
  try {
    const info = await startup.getStartupStatus();
    res.json(info);
  } catch (err) {
    res.status(500).json({ error: String(err) });
  }
});

// In-memory jobs store for prototype
const jobs = {};
let nextJobId = 1;

app.post('/upload', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'no file uploaded' });
  const jobId = String(nextJobId++);
  jobs[jobId] = { status: 'uploaded', file: req.file.path, originalName: req.file.originalname, results: null };
  res.json({ jobId, filename: req.file.originalname });
});

app.post('/run', async (req, res) => {
  // Body: { jobId, agents: ['parser','legal'] }
  const { jobId, agents } = req.body;
  if (!jobId || !jobs[jobId]) return res.status(400).json({ error: 'invalid jobId' });
  const job = jobs[jobId];
  job.status = 'running';

  try {
    let canonical = null;
    const out = {};
    if (agents.includes('parser')) {
      const parseOut = await runParser(job.file, job.originalName);
      out.parser = parseOut;
      canonical = parseOut.canonical;
    }

    if (agents.includes('legal')) {
      // require canonical text; fallback to raw parsed text
      const input = canonical ?? (out.parser && out.parser.text) ?? '';
      const legalOut = await runLegalAgent(input, { jurisdiction: 'EU,UK,Global' });
      out.legal = legalOut;
    }

    if (agents.includes('risk')) {
      const input = canonical ?? (out.parser && out.parser.text) ?? '';
      const riskOut = await runRiskAgent(input, {});
      out.risk = riskOut;
    }

    if (agents.includes('reporting')) {
      const input = canonical ?? (out.parser && out.parser.text) ?? '';
      const reportingOut = await runReportingAgent(input, { taxonomies: ['GRI','SASB','EU_Taxonomy'] });
      out.reporting = reportingOut;
    }

    if (agents.includes('climate')) {
      const input = canonical ?? (out.parser && out.parser.text) ?? '';
      const climateOut = await runClimateAgent(input, {});
      out.climate = climateOut;
    }

    if (agents.includes('supply_chain')) {
      const input = canonical ?? (out.parser && out.parser.text) ?? '';
      out.supply_chain = await runSupplyChainAgent(input, {});
    }

    if (agents.includes('financial')) {
      const input = canonical ?? (out.parser && out.parser.text) ?? '';
      out.financial = await runFinancialAgent(input, {});
    }

    if (agents.includes('audit')) {
      const input = canonical ?? (out.parser && out.parser.text) ?? '';
      out.audit = await runAuditAgent(input, {});
    }

    if (agents.includes('summarizer')) {
      const input = canonical ?? (out.parser && out.parser.text) ?? '';
      out.summary = await runSummarizer(input, {});
    }

    job.status = 'done';
    job.results = out;
    res.json({ jobId, status: job.status, results: out });
  } catch (err) {
    console.error(err);
    job.status = 'error';
    job.results = { error: String(err) };
    res.status(500).json({ error: String(err) });
  }
});

app.get('/results/:jobId', (req, res) => {
  const { jobId } = req.params;
  if (!jobs[jobId]) return res.status(404).json({ error: 'job not found' });
  res.json(jobs[jobId]);
});

const PORT = process.env.PORT || 4000;
app.listen(PORT, () => console.log(`Backend listening on http://localhost:${PORT}`));
