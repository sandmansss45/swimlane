# Hosting on GitHub Pages

The app is already git-initialized and committed locally (branch `main`).
Everything below just needs your own GitHub login — I can't push on your
behalf without your credentials.

## 1. Create the GitHub repo

1. Go to https://github.com/new
2. Name it anything (e.g. `swimlane-studio-web`). Public or private both work
   for Pages. **Don't** initialize with a README/gitignore — this project
   already has both, and that would conflict with the push below.
3. Click **Create repository** and copy the URL it gives you
   (`https://github.com/<you>/<repo-name>.git`).

## 2. Push this project to it

Run these from `C:\Users\WilliamSands\Downloads\swimlane-studio-web`
(PowerShell, with a terminal that has git on PATH — this session added a
portable git install to your user PATH, so a fresh terminal window should
already pick it up):

```powershell
git remote add origin https://github.com/<you>/<repo-name>.git
git push -u origin main
```

If it prompts for credentials, use a GitHub Personal Access Token as the
password (GitHub stopped accepting account passwords for git pushes) —
create one at https://github.com/settings/tokens if you don't have one.

## 3. Turn on GitHub Pages (one-time, in the repo's settings)

1. On the repo's GitHub page → **Settings** → **Pages**.
2. Under **Build and deployment** → **Source**, choose **GitHub Actions**
   (not "Deploy from a branch" — this project already includes the right
   workflow at `.github/workflows/deploy.yml`).
3. That's it. The workflow will run automatically on this push (and every
   push to `main` after) and publish to
   `https://<you>.github.io/<repo-name>/`.

## 4. Watch it deploy

Repo page → **Actions** tab → the "Deploy to GitHub Pages" run. Once it's
green, the URL from step 3 is live for anyone signed into a `qleapenergy.com`
account (once the Entra app registration from `AUTH_SETUP.md` is done).

## 5. Send me the final URL

Once it's live, tell me the exact `https://<you>.github.io/<repo-name>/`
URL — I'll add it as the second Entra redirect URI's origin (step 6 in
`AUTH_SETUP.md`) if it hasn't been added yet.
