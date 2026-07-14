# LeadSquared FY27 Pipeline Dashboard — Hosted

This is your Claude Artifact dashboard turned into a real website with:

- 🔐 **Login** — username + password (a few fixed accounts you control)
- 💾 **Shared, persistent data** — when anyone uploads the CSV/XLSX files, the data
  is saved to a database and is visible to every logged-in user (it survives
  refreshes and restarts)
- 📎 **Downloadable source files** — the original uploaded files are kept and can
  be downloaded again from the dashboard

It's a [Next.js](https://nextjs.org) app designed to deploy on **Vercel (free tier)**
using a **Neon Postgres** database (also free tier) for storage.

---

## What you need (all free)

1. A **GitHub** account (this code lives in your repo already).
2. A **Vercel** account — https://vercel.com (you said you have one).
3. That's it. The database is added from inside Vercel in one click.

---

## Deploy to Vercel (step by step)

### 1. Import the project
1. Go to https://vercel.com/new
2. Select this GitHub repository (`claudeskillsrepo`).
3. **Important:** set **Root Directory** to `pipeline-dashboard`
   (click *Edit* next to Root Directory and pick the `pipeline-dashboard` folder).
4. Framework preset should auto-detect as **Next.js**. Leave build settings default.
5. Don't click Deploy yet — first add the database and environment variables below.

### 2. Add the database (Neon Postgres)
1. In your new Vercel project, open the **Storage** tab.
2. Click **Create Database → Neon (Serverless Postgres)** and follow the prompts
   (choose the free plan).
3. When it asks, **connect it to this project**. Vercel automatically adds the
   `POSTGRES_URL` environment variable for you. You don't need to touch the
   database again — the app creates its tables automatically on first use.

### 3. Add the login + secret environment variables
In the project's **Settings → Environment Variables**, add:

| Name          | Value (example)                                   | Notes |
|---------------|---------------------------------------------------|-------|
| `APP_USERS`   | `vinayak:MyStrongPass1,teammate:AnotherPass2`     | Your login accounts. Format is `user:password` pairs separated by commas. |
| `AUTH_SECRET` | a long random string                              | Used to sign login sessions. Generate one below. |

To generate a good `AUTH_SECRET`, run this on your computer and paste the output:
```bash
openssl rand -base64 32
```
(or just type a long random string of 40+ characters).

### 4. Deploy
Click **Deploy**. After a minute or two you'll get a URL like
`https://your-project.vercel.app`. Open it, log in with one of the accounts you
put in `APP_USERS`, and upload your CSV/XLSX files. Done.

---

## Managing users later

Edit the `APP_USERS` environment variable in **Vercel → Settings → Environment
Variables**, then redeploy (Vercel → Deployments → ⋯ → Redeploy). Add users by
appending more `user:password` pairs; remove a user by deleting their pair.

> Note: this is simple shared-password auth, appropriate for a small internal
> team tool. Passwords are stored only as Vercel environment variables (never sent
> to the browser). For anything more sensitive, consider a full identity provider.

---

## How data is stored

- Uploaded and parsed data is written to a `kv_store` table in Postgres. All users
  share the same rows, so an upload by one person shows up for everyone.
- The original files are stored in a `files` table and served back for download.
- Both tables are created automatically the first time the app talks to the
  database — there's no manual database setup.

The dashboard keeps four inputs, exactly like the original artifact:
- **Opportunity CSV** → Pipeline Number View tab
- **Target File (XLSX)** → quarterly MRR targets
- **MRR Transaction CSV** → MRR Performance tab
- **Renewals CSV** → Renewal Risks tab

---

## Running locally (optional, for development)

```bash
cd pipeline-dashboard
cp .env.example .env.local     # then edit .env.local
#   - set APP_USERS and AUTH_SECRET
#   - set POSTGRES_URL to a Neon connection string (create a free DB at neon.tech)
npm install
npm run dev                    # open http://localhost:3000
```

You can create a free Neon database at https://neon.tech and paste its connection
string into `POSTGRES_URL` for local testing.

---

## Project structure

```
pipeline-dashboard/
├── app/
│   ├── page.tsx                 # protected home → renders the dashboard
│   ├── login/page.tsx           # login screen
│   └── api/
│       ├── login / logout / session
│       ├── data/[key]/          # key-value store (replaces window.storage)
│       └── files/               # upload / list / download originals (window.fs)
├── components/
│   ├── Dashboard.tsx            # your original dashboard (ported, near-verbatim)
│   ├── DashboardShell.tsx       # header, sign-out, mounts the storage shims
│   └── FileDownloads.tsx        # "Uploaded source files" download panel
├── lib/
│   ├── auth.ts                  # session cookies + credential checks
│   ├── db.ts                    # Postgres + auto-created schema
│   └── browserShims.ts          # window.storage / window.fs → backend APIs
└── middleware.ts                # redirects logged-out users to /login
```

Your original dashboard logic is untouched except for two small additions: a
`'use client'` line at the top, and one line in each upload handler that saves the
original file so it stays downloadable.
