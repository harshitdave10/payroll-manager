# Payroll Manager — Deployment Guide

## Architecture
```
GitHub (private repo)
  ├── frontend/   → GitHub Pages  (your UI)
  └── backend/    → Render        (your Python API)
```

---

## STEP 1 — Create the GitHub Repository

1. Go to **https://github.com/new**
2. Fill in:
   - **Repository name**: `payroll-manager` (or anything you like)
   - **Visibility**: ✅ **Private**
   - Leave everything else unchecked
3. Click **Create repository**
4. Copy the repo URL shown (e.g. `https://github.com/YOUR_USERNAME/payroll-manager.git`)

---

## STEP 2 — Push Code to GitHub

Open a terminal in the `payroll-app` folder and run:

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/payroll-manager.git
git push -u origin main
```

---

## STEP 3 — Deploy the Backend to Render

1. Go to **https://render.com** → Sign up / Log in (connect your GitHub account)
2. Click **New → Web Service**
3. Select **Connect a repository** → choose `payroll-manager`
4. Configure the service:

   | Field | Value |
   |-------|-------|
   | Name | `payroll-backend` |
   | Root Directory | `backend` |
   | Runtime | `Python 3` |
   | Build Command | `pip install -r requirements.txt` |
   | Start Command | `gunicorn app:app` |
   | Instance Type | `Free` |

5. Scroll down to **Environment Variables** — click **Add Environment Variable** for each:

   | Key | Value |
   |-----|-------|
   | `PAYROLL_USER` | your chosen username |
   | `PAYROLL_PASS` | your chosen password |
   | `FRONTEND_ORIGIN` | (fill in Step 5 after you get the Pages URL) |

6. Click **Create Web Service**
7. Wait ~2 minutes for the first deploy. When done, copy your URL:
   ```
   https://payroll-backend.onrender.com
   ```

---

## STEP 4 — Enable GitHub Pages for the Frontend

1. In your GitHub repo, go to **Settings → Pages**
2. Under **Source**, select **GitHub Actions**
3. The workflow will auto-run on the next push to main
4. After it runs (~1 min), your frontend will be live at:
   ```
   https://YOUR_USERNAME.github.io/payroll-manager/
   ```

> **Note:** GitHub Pages on private repos is publicly accessible by URL — the app is protected by the login screen.

---

## STEP 5 — Connect Frontend to Backend

1. Open `frontend/config.js` and replace the placeholder:
   ```js
   window.PAYROLL_API_URL = 'https://payroll-backend.onrender.com';
   ```
2. Go back to Render → your service → **Environment** → set:
   ```
   FRONTEND_ORIGIN = https://YOUR_USERNAME.github.io
   ```
3. Commit and push:
   ```bash
   git add frontend/config.js
   git commit -m "Set backend URL"
   git push
   ```
4. GitHub Actions will auto-redeploy the frontend in ~1 minute.

---

## STEP 6 — Verify Everything Works

1. Visit your GitHub Pages URL
2. Log in with the credentials you set in Render
3. Test **New Month** with your Truein file — you should get `final_payroll.xlsx` downloaded

---

## Local Development

### Backend
```bash
cd backend
pip install -r requirements.txt
export PAYROLL_USER="admin"
export PAYROLL_PASS="payroll@2026"
python app.py        # Runs on http://localhost:5000
```

### Frontend
```bash
cd frontend
python -m http.server 8080
# Open http://localhost:8080
# config.js falls back to localhost:5000 automatically when PAYROLL_API_URL is not set
```

---

## Making Future Updates

- Any push to `main` touching `frontend/` → GitHub Pages auto-redeploys
- Any push to `main` → Render auto-redeploys backend (enable in Render → Settings → Auto-Deploy)

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| Login fails on live site | Check `PAYROLL_USER` / `PAYROLL_PASS` in Render env vars |
| CORS error in browser console | Make sure `FRONTEND_ORIGIN` in Render matches your exact Pages URL (no trailing slash) |
| Render backend slow first load | Free tier sleeps after inactivity — first request takes ~30s to wake up, normal behaviour |
| GitHub Pages not updating | Go to Actions tab in your repo and check the workflow run status |
