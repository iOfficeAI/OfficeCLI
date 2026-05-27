import React, { useState } from 'react'
import axios from 'axios'

export default function App() {
  const [file, setFile] = useState(null)
  const [jobId, setJobId] = useState(null)
  const [agents, setAgents] = useState({ parser: true, legal: true, risk: false, reporting: false, climate: false, supply_chain: false, financial: false, audit: false, summarizer: false })
  const [results, setResults] = useState(null)

  const upload = async () => {
    if (!file) return alert('Choose a file first')
    const fd = new FormData()
    fd.append('file', file)
    const res = await axios.post('http://localhost:4000/upload', fd)
    setJobId(res.data.jobId)
    alert('Uploaded, jobId=' + res.data.jobId)
  }

  const run = async () => {
    if (!jobId) return alert('Upload first')
    const selected = Object.keys(agents).filter(k => agents[k])
    const res = await axios.post('http://localhost:4000/run', { jobId, agents: selected })
    setResults(res.data.results)
  }

  return (
    <div style={{ padding: 24, fontFamily: 'sans-serif' }}>
      <h1>ESG Agent Pool Prototype</h1>
      <div>
        <input type="file" onChange={e => setFile(e.target.files[0])} />
        <button onClick={upload}>Upload</button>
      </div>
      <div style={{ marginTop: 12 }}>
        <strong>Agents</strong>
        <label style={{ marginLeft: 8 }}><input type="checkbox" checked={agents.parser} onChange={e => setAgents(a => ({...a, parser: e.target.checked}))} /> Parser</label>
        <label style={{ marginLeft: 8 }}><input type="checkbox" checked={agents.legal} onChange={e => setAgents(a => ({...a, legal: e.target.checked}))} /> Legal Compliance</label>
        <label style={{ marginLeft: 8 }}><input type="checkbox" checked={agents.risk} onChange={e => setAgents(a => ({...a, risk: e.target.checked}))} /> Risk Assessor</label>
        <label style={{ marginLeft: 8 }}><input type="checkbox" checked={agents.reporting} onChange={e => setAgents(a => ({...a, reporting: e.target.checked}))} /> Reporting Mapper</label>
        <label style={{ marginLeft: 8 }}><input type="checkbox" checked={agents.climate} onChange={e => setAgents(a => ({...a, climate: e.target.checked}))} /> Climate Analyst</label>
        <label style={{ marginLeft: 8 }}><input type="checkbox" checked={agents.supply_chain} onChange={e => setAgents(a => ({...a, supply_chain: e.target.checked}))} /> Supply Chain</label>
        <label style={{ marginLeft: 8 }}><input type="checkbox" checked={agents.financial} onChange={e => setAgents(a => ({...a, financial: e.target.checked}))} /> Financial Impact</label>
        <label style={{ marginLeft: 8 }}><input type="checkbox" checked={agents.audit} onChange={e => setAgents(a => ({...a, audit: e.target.checked}))} /> Audit / Assurance</label>
        <label style={{ marginLeft: 8 }}><input type="checkbox" checked={agents.summarizer} onChange={e => setAgents(a => ({...a, summarizer: e.target.checked}))} /> Executive Summarizer</label>
        <button style={{ marginLeft: 8 }} onClick={run}>Run Selected Agents</button>
      </div>

      {results && (
        <div style={{ marginTop: 20 }}>
          <h2>Results</h2>
          <pre style={{ whiteSpace: 'pre-wrap', background: '#f6f6f6', padding: 12 }}>{JSON.stringify(results, null, 2)}</pre>
        </div>
      )}
    </div>
  )
}
