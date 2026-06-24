# Household Ledger

A shared household budget & expense tracker — import bank/card statements, categorize
spending, set monthly budgets, and track everything together in real time.

This is a standalone version of a Claude.ai artifact, rebuilt to run as a normal website
so it can be hosted on GitHub Pages and opened from any browser (including a Fire TV
Stick's browser). Data is stored in **Firebase Firestore** (free tier) so changes made by
one person show up live for everyone else — no manual refresh needed.

Because there's no Claude.ai sandbox here, the app requires **sign-in** so your
household's financial data isn't sitting open to anyone who finds the URL. You add team
members manually in the Firebase console — there's no public sign-up form.

---

## 1. Create a Firebase project (free)

1. Go to <https://console.firebase.google.com> and click **Add project**. Give it any
   name (e.g. "household-ledger") and finish the wizard (you can skip Google Analytics).
2. In the left sidebar, go to **Build → Firestore Database → Create database**.
   - Choose a location close to you.
   - Start in **production mode** (we'll set rules below).
3. Once created, go to the **Rules** tab of Firestore and replace the contents with what's
   in [`firestore.rules`](./firestore.rules) in this repo, then click **Publish**. This
   restricts the data to signed-in users only.
4. In the left sidebar, go to **Build → Authentication → Get started**, then enable the
   **Email/Password** sign-in method (Sign-in method tab → Email/Password → Enable → Save).
5. Still in Authentication, go to the **Users** tab and click **Add user** for yourself and
   anyone else who should have access. This is how you control who's on the "team" — there
   is no self-serve sign-up screen in the app on purpose.
6. Go to **Project settings** (gear icon, top left) → scroll to **Your apps** → click the
   `</>` (web) icon to register a new web app (any nickname is fine, no need for Firebase
   Hosting). You'll be shown a config object with values like `apiKey`, `authDomain`, etc.
   Keep this tab open — you'll need these values in step 3 below.

## 2. Run it locally (optional, but good for testing first)

```bash
npm install
cp .env.example .env
# paste your Firebase config values into .env
npm run dev
```

Open the URL it prints (usually `http://localhost:5173`), sign in with one of the users
you added in step 1.5, and confirm transactions/imports/budgets work and save.

## 3. Push to GitHub

1. Create a new repository on GitHub (public or private — private is recommended since
   this manages financial data, even though it's access-gated).
2. Push this project:
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-repo>.git
   git push -u origin main
   ```
   `.env` is already in `.gitignore`, so your Firebase keys won't be committed in plain
   text — they get supplied to the build separately, via GitHub Actions secrets (next step).

## 4. Add your Firebase config as GitHub Secrets

In your new repo: **Settings → Secrets and variables → Actions → New repository secret**.
Add each of these (same values as your `.env` file):

- `VITE_FIREBASE_API_KEY`
- `VITE_FIREBASE_AUTH_DOMAIN`
- `VITE_FIREBASE_PROJECT_ID`
- `VITE_FIREBASE_STORAGE_BUCKET`
- `VITE_FIREBASE_MESSAGING_SENDER_ID`
- `VITE_FIREBASE_APP_ID`

## 5. Turn on GitHub Pages

In your repo: **Settings → Pages → Build and deployment → Source → GitHub Actions**.

That's it — the included workflow (`.github/workflows/deploy.yml`) builds and deploys the
site automatically every time you push to `main`. After the first push, check the
**Actions** tab for progress; once it's green, your site is live at:

```
https://<your-username>.github.io/<your-repo>/
```

## 6. Open it on the Fire Stick

The Fire TV Stick's built-in **Silk Browser** (or install **Downloader** / **Firefox** from
the Amazon App Store if you want a more modern browser) can open that GitHub Pages URL like
any website. Sign in with one of the accounts you created in step 1.5. Typing with a remote
is slow, so this is best for *viewing* the dashboard — do data entry and imports from a
phone or computer instead.

---

## Notes on security

- A Firebase web API key is **not a secret** in the traditional sense — it's safe to ship
  in a public JS bundle, because access is actually controlled by the Firestore rules and
  Authentication, not by hiding the key. What matters is keeping `firestore.rules` requiring
  `request.auth != null`, and only adding trusted people under Authentication → Users.
- If you'd rather skip sign-in entirely (e.g. just testing), you can set
  `firestore.rules` to `allow read, write: if true;` and remove the auth screen — but then
  *anyone* who finds your site URL can read and edit your transactions. Not recommended for
  real financial data on a public GitHub Pages URL.

## Notes on cost

Firestore's free "Spark" tier includes 50,000 document reads and 20,000 writes per day,
far more than a household budget tool for a small team will ever use. GitHub Pages hosting
is free for public repos (and free for private repos too, with a generous Actions minutes
allowance).

## Project structure

```
├── src/
│   ├── App.jsx        # the whole application
│   ├── firebase.js     # Firebase init (reads config from env vars)
│   ├── storage.js       # Firestore read/write helpers (replaces Claude's window.storage)
│   └── main.jsx          # React entry point
├── firestore.rules        # security rules — copy into Firebase console
├── .env.example             # template for local Firebase config
└── .github/workflows/deploy.yml   # auto-deploy to GitHub Pages on push
```
