# Data to Decisions — class labs

Student-facing pages for a guest lecture on data analytics and AI. Static HTML plus four small serverless functions that collect class results. Runs on Vercel's free plan.

| Page | What it is |
|---|---|
| `/` | Landing page with links to the two student pages |
| `/02_BI_Dashboard.html` | Lab 1 — BI Detective: filters, charts, four auto-checked missions |
| `/04_BI_Challenge.html` | Mini case — Arunika Fiber: brief, six questions, data dictionary, dashboard submission |
| `/data/Arunika_Fiber_BI_Case.xlsx` | The case workbook (14 sheets) students load into Google Sheets and Data Studio |
| `/teacher.html` | Lecturer view: live results, CSV download, clear results (not linked from the landing page) |

## Deploy on Vercel

1. Go to vercel.com → **Add New → Project** → import this GitHub repository.
2. Framework preset: **Other**. Leave build command and output directory empty. Deploy.
3. In the project, open **Storage → Create Database → Upstash for Redis** (free plan) and connect it to this project. This adds `KV_REST_API_URL` and `KV_REST_API_TOKEN` (or `UPSTASH_REDIS_REST_URL` / `UPSTASH_REDIS_REST_TOKEN`); the code accepts either pair.
4. In **Settings → Environment Variables** add `ADMIN_KEY` with a password of your choice. It is needed to clear results.
5. **Deployments → Redeploy** so the functions pick up the new variables.

Until step 3 is done the labs still work; results are simply not saved.

Every push to the main branch redeploys automatically.

## Results API

| Route | Use |
|---|---|
| `POST /api/submit` | A page sends one result: a mission answer, or a BI Challenge dashboard (nickname, question, link, one-sentence recommendation). Fields are whitelisted and length-capped; dashboard links must be Data Studio links. |
| `GET /api/board` | All results. Read by `teacher.html` and by the lecturer's slide deck. |
| `GET /api/info` | Whether the store is connected, and the number of results. |
| `POST /api/reset` | Clears results. Needs the header `X-Admin-Key`. |

Results are nicknames, answers, dashboard links and one-sentence recommendations. `GET /api/board` is public by design so the lecturer's deck can read it.

## Data

The dashboard uses IBM's public "Telco Customer Churn" sample (7,043 fictional customers). 

The BI Challenge uses a generated dataset about a fictional home-internet company, Arunika Fiber. No real company or customer data. To rebuild it: `python tools/build_case_data.py` (needs numpy, pandas, xlsxwriter). The seed is fixed, so the output is the same every time.

After uploading the workbook to Google Sheets, paste the share link into `data/case_config.js` so the case page shows an "Open the Google Sheet" button.
