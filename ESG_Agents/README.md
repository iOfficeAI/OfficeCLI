ESG Agent Pool Prototype

This workspace is a prototype for an ESG analysis agent pool.

Structure:
- backend: Node.js Express backend (file upload, agent runner)
- frontend: React frontend (upload UI, agent selection, results)
- agents: prompt specs and simple agent implementations (in backend/agents)

Quick start (prototype):

1. Backend

```bash
cd backend
npm install
npm start
```

2. Frontend

```bash
cd frontend
npm install
npm run dev
```

Configuration:
- Set `OPENAI_API_KEY` for legal agent calls to OpenAI.
- OfficeCLI usage: backend adapters will attempt to call `officecli` binary if present to extract Office formats.

This is an initial scaffold. See backend/README.md and frontend/README.md for details.
