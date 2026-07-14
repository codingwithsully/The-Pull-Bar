# The Pull Bar — deployment guide

This is your store app, ready to put online so you can install it on your
iPhone and iPad. Follow these steps in order. No coding needed.

---

## What you're doing

1. Put this project on GitHub (a free code-storage site).
2. Connect it to Vercel (a free hosting site) — this gives you a web address.
3. Open that web address in Safari on your iPhone and "Add to Home Screen."

The whole thing is free and takes about 15–20 minutes the first time.

---

## Step 1 — Make a GitHub account and upload this project

1. Go to https://github.com and sign up (free).
2. Click the **+** (top right) → **New repository**.
3. Name it `the-pull-bar`, leave it Public or Private, click **Create repository**.
4. On the next page, click **uploading an existing file**.
5. Drag in ALL the files and folders from this project (the whole `pullbar`
   folder's contents — `package.json`, `index.html`, the `src` folder, the
   `public` folder, etc.).
6. Click **Commit changes**.

## Step 2 — Deploy on Vercel

1. Go to https://vercel.com and click **Sign Up** → **Continue with GitHub**
   (this links the two accounts).
2. Click **Add New… → Project**.
3. Find `the-pull-bar` in the list and click **Import**.
4. Vercel auto-detects it's a Vite app — you don't need to change any
   settings. Just click **Deploy**.
5. Wait ~1 minute. You'll get a live web address like
   `the-pull-bar.vercel.app`.

## Step 3 — Install on your iPhone / iPad

1. Open **Safari** (must be Safari, not Chrome) on your iPhone.
2. Go to your Vercel web address.
3. Tap the **Share** button (the square with an up-arrow).
4. Scroll down and tap **Add to Home Screen**.
5. Tap **Add**. The Pull Bar icon now sits on your home screen and opens
   full-screen like a real app.

Do the same on your iPad. Do it on your laptop too by just bookmarking the
address.

---

## Updating the app later

When you want changes, I give you updated files → you upload them to GitHub
the same way (Step 1, items 4–6) → Vercel automatically rebuilds and your
installed app updates on its own. No reinstalling.

---

## Important note about your data (please read)

Right now the app stores everything **on the device you're using**, inside
that browser. This means:

- Testing on one iPhone: works great.
- Your iPhone, iPad, and laptop will each have **their own separate**
  inventory — they do NOT share data yet.
- Clearing your browser data would erase the app's data on that device.

This is fine for testing and deciding if the app works for you. When you're
ready for real daily use across multiple devices (and safe backups), the next
step is connecting a hosted database. That's a separate build — just ask when
you get there.
