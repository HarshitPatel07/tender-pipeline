# Setting up your Tender Pipeline website (GitHub + Streamlit)

One-time setup, about 5 minutes, all in the browser — no installing anything,
no command line. After this, it's a bookmark you open anytime.

## 1. Create a free GitHub account (skip if you already have one)

1. Go to **github.com** → **Sign up**.
2. Verify your email.

## 2. Create a private repository (a folder for these files)

1. Go to **github.com/new**.
2. **Repository name**: `tender-pipeline` (or anything you like).
3. **Visibility**: choose **Private** — this keeps your tender documents and
   this code restricted to you, since it handles your firm's real bid data.
4. Click **Create repository**.

## 3. Upload the four files

You land on your new, empty repository.

1. Click **uploading an existing file** (or **Add file → Upload files**).
2. Drag in all four files I gave you: `app.py`, `tender_extractor.py`,
   `requirements.txt`, `packages.txt`.
3. Scroll down, click **Commit changes**.

## 4. Create your free account on Streamlit Community Cloud

1. Go to **share.streamlit.io**.
2. Click **Sign up**, then **Continue with GitHub** — sign in with the
   account from step 1 and approve the connection.

## 5. Deploy your app

1. Click **Create app** (or **New app**).
2. **Repository**: pick `tender-pipeline` (the one you just uploaded to).
3. **Branch**: `main`.
4. **Main file path**: `app.py`.
5. Click **Deploy**.

The app builds for a few minutes the first time (installing the tools it
needs), then goes live at a URL like
`https://tender-pipeline-yourname.streamlit.app`. **Bookmark that URL.**

## 6. Use it

The page shows a box for your Drive link, an optional box for a Gemini key,
and an **Extract tenders** button. Paste, press, wait — the dashboard appears
on the same page.

## Keeping it private

Even with a private GitHub repo, by default anyone with your Streamlit link
can open the app. To restrict it to just you:

1. On your app's page on share.streamlit.io, open **Settings → Sharing**.
2. Switch to **Only specific people can view this app**.
3. Add your own email (and anyone else at your firm who should use it).

## If it goes to sleep

Free apps sleep after a period of no use. Opening the bookmarked link wakes
it up again — the first load after sleeping can take 30-60 seconds, then
it's normal.

## Updating it later

If I send you an updated `app.py` or `tender_extractor.py`: open your GitHub
repository, click the file, click the pencil (**Edit**) or use **Add file →
Upload files** to replace it, commit the change. Streamlit rebuilds your live
app automatically within a minute or two — no redeploying by hand.
