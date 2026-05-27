Backend README

Run locally:

```bash
cd backend
npm install
OPENAI_API_KEY=yourkey node server.js
```

Endpoints:
- POST /upload (multipart form-data field `file`) → { jobId }
- POST /run { jobId, agents: ['parser','legal'] } → runs selected agents and returns results
- GET /results/:jobId → get results and status

Notes:
- The parser uses OfficeCLI if available; otherwise falls back to simple extraction for txt/csv.
- Legal agent calls OpenAI; set `OPENAI_API_KEY` and optional `OPENAI_MODEL` env var.

Additional agents:
- Risk, Reporting, and Climate agents are implemented in `backend/agents` and can be invoked via the `/run` endpoint.

Example usage (run multiple agents):

```json
POST /run
{
	"jobId": "1",
	"agents": ["parser","legal","risk","reporting","climate"]
}
```

Notes:
- Agents call OpenAI; set `OPENAI_API_KEY` and optional `OPENAI_MODEL` env var.
- Sample fixtures are available in `backend/fixtures` for quick testing.

Quick test runner:

```bash
cd backend
node run_tests.js
```

This will run the parser and all agents against sample fixtures and print outputs.

Notes on testing without an API key:
- The prototype includes a lightweight LLM wrapper (`llm.js`) that returns canned responses when `OPENAI_API_KEY` is not set, so `node run_tests.js` and `npm test` will run offline for quick verification. Set `OPENAI_API_KEY` to run real OpenAI-backed agents.

Running tests via npm:

```bash
cd backend
npm test
```

CI: A GitHub Actions workflow is provided at `.github/workflows/nodejs-ci.yml` which runs the backend tests. Set `OPENAI_API_KEY` in repository secrets to enable LLM-backed agent tests in CI.
