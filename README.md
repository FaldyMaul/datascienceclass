# Data to Decisions — class labs

Student-facing pages for a guest lecture on data analytics and AI. Static HTML plus four small serverless functions that collect class results. Runs on Vercel's free plan.

| Page | What it is |
|---|---|
| `/` | Landing page with links to the three labs |
| `/02_BI_Dashboard.html` | Lab 1 — BI Detective: filters, charts, four auto-checked missions |
| `/03_Churn_Rescue_Game.html` | Lab 2 — Churn Rescue: four-round decision game |
| `/04_IE_Cafe_Mini_Case.html` | Mini case — menu analysis for a small cafe |
| `/teacher.html` | Lecturer view: live results, CSV download, clear results (not linked from the landing page) |

## Deploy on Vercel

1. Go to vercel.com → **Add New → Project** → import this GitHub repository.
2. Framework preset: **Other**. Leave build command and output directory empty. Deploy.
3. In the project, open **Storage → Create Database → Upstash for Redis** (free plan) and connect it to this project. This adds `KV_REST_API_URL` and `KV_REST_API_TOKEN` (or `UPSTASH_REDIS_REST_URL` / `UPSTASH_REDIS_REST_TOKEN`); the code accepts either pair.
4. In **Settings → Environment Variables** add `ADMIN_KEY` with a password of your choice. It is needed to clear results.
5. **Deployments → Redeploy** so the functions pick up the new variables.

Until step 3 is done the labs still work; scores are simply not saved and the game shows a result code instead.

Every push to the main branch redeploys automatically.

## Results API

| Route | Use |
|---|---|
| `POST /api/submit` | A lab page sends one result. Fields are whitelisted and capped at 40 characters. |
| `GET /api/board` | All results. Read by `teacher.html` and by the lecturer's slide deck. |
| `GET /api/info` | Whether the store is connected, and the number of results. |
| `POST /api/reset` | Clears results. Needs the header `X-Admin-Key`. |

Results are nicknames and scores only. `GET /api/board` is public by design so the lecturer's deck can read it.

## Data

The dashboard uses IBM's public "Telco Customer Churn" sample (7,043 fictional customers). Rupiah amounts, call capacity and cafe sales volumes are teaching numbers.
