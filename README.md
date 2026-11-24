# ExamGenius AI
```
cd backend && python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env      # add your OPENAI_API_KEY (stays on the server only)
python app.py             # open http://localhost:5000
```
Flask serves the API and the React/Tailwind/Chart.js frontend (CDN builds, no Node step). SQLite file: `backend/examgenius.db`. Uploads: `backend/uploads/<user>/` (never served publicly).

## Features
Professor modes (Easy/Normal/Strict/Expert), final exam simulation from several lectures, AI weakness detector with personalized quizzes, Study Battle (private rooms, live leaderboard), exam preparation guide, practice mode, analytics, dark/light mode.

## Test
`cd backend && python smoke_test.py` runs an offline end-to-end test of every endpoint (the AI call is mocked, so it needs no API key).

## Publish a shareable link
The app needs a server (Flask, database, your OpenAI key), so it can't be a static page. The simplest route is Render:
1. Push this folder to a GitHub repo.
2. On render.com choose New > Blueprint and select the repo. `render.yaml` sets everything up.
3. Add your `OPENAI_API_KEY` when asked. You get a public `https://...onrender.com` link to share.
This `render.yaml` uses Render's free plan: the site sleeps after 15 idle minutes, and accounts, uploads and results are wiped whenever it restarts (fine for demos). To keep data, switch to a paid plan, add a persistent disk mounted at `/data` and set `DATA_DIR=/data`.
