# Setting up your Tender Pipeline website

One-time setup, about 5 minutes. After this, it's a bookmark you open anytime.

## 1. Create a free account

1. Go to **huggingface.co** → click **Sign Up** (top right).
2. Sign up with your email. Free, no card needed.
3. Verify your email if it asks you to.

## 2. Create your private Space (this becomes your website)

1. Go to **huggingface.co/new-space**.
2. **Space name**: `tender-pipeline` (or anything you like).
3. **License**: leave as-is.
4. **Select the Space SDK**: choose **Gradio**.
5. **Space hardware**: leave on the free **CPU basic**.
6. **Visibility**: choose **Private**. This keeps it restricted to your own
   Hugging Face account — nobody else can open your dashboard or see your
   tender documents, since this handles your firm's real bid data.
7. Click **Create Space**.

## 3. Upload the four files

You land on an empty Space with a **Files** tab.

1. Click **Files** → **Add file** → **Upload files**.
2. Drag in all four files I gave you: `app.py`, `tender_extractor.py`,
   `requirements.txt`, `packages.txt`.
3. Scroll down, click **Commit changes to main**.

## 4. Wait for it to build

The Space switches to **Building** (installing the tools it needs — a few
minutes the first time), then **Running**. That's it — your website is live.

## 5. Use it

The page shows a box for your Drive link, an optional box for a Gemini key,
and an **Extract tenders** button. Paste, press, wait — the dashboard appears
on the same page.

**Bookmark this Space's URL.** Every future time: open the bookmark, paste
the new folder link, press Extract.

## If it goes to sleep

Free Spaces sleep after a period of no use. Opening the bookmarked link wakes
it up again — the first load after sleeping can take 30-60 seconds, then it's
normal.

## Updating it later

If I send you an updated `app.py` or `tender_extractor.py`, go back to your
Space's **Files** tab and upload the new one over the old one (same
filename) — it rebuilds automatically.
