# EPD Data Extractor — Render deployment

This is a **Flask** website, not Streamlit. The PDF extraction engine, Gemini AI review and 111-column Excel export are preserved. The desktop Tkinter UI is not needed on the server.

## Update your existing GitHub repository

Open `https://github.com/helani2004/EPD-Data-Extractor-Web` → **Add file → Upload files**.
Upload all the files **inside this extracted folder** (not the ZIP). The essential new files are `web_app.py`, `extractor_core.py`, `render.yaml`, and the updated `requirements.txt`. Keep `engine.py`, `ai_review.py`, and `EPDinputFormat.xlsx` in the repository root. Click **Commit changes**. You can leave the old Streamlit files in the repo; Render will ignore them.

## Deploy on Render

1. Go to `https://dashboard.render.com/` and sign in.
2. Click **New → Blueprint** and connect your GitHub repository `helani2004/EPD-Data-Extractor-Web`. Render reads `render.yaml`.
3. When prompted for `GEMINI_API_KEY`, paste your own Gemini API key. **Never put this key in GitHub.**
4. Approve/deploy the Blueprint. Wait for the build and open the `onrender.com` URL.
5. Upload one small EPD PDF, review values, click **Save edits**, then **Download Excel**.

If you prefer **New → Web Service**, choose the same repository, Python runtime, free plan, build command `pip install -r requirements.txt`, and start command `gunicorn web_app:app --bind 0.0.0.0:$PORT --workers 1 --threads 2 --timeout 300`. Add `GEMINI_API_KEY` in **Environment** and set `FLASK_SECRET_KEY` to a long random value.

## Limitations

- Render free services sleep when inactive. First visit after sleep can be slow.
- Uploaded PDFs are processed in temporary folders and deleted after extraction; results are held in server memory, isolated by browser session, and expire after two hours of inactivity. They are **not permanently saved**. Download the Excel before leaving.
- PDF text extraction works without Tesseract. OCR of image-only/scanned PDFs needs the Tesseract executable installed on the host; this package does not install it automatically on Render. The app warns when OCR is missing.
- Gemini PDF processing uses your account's rate limits and may fail during high demand. The app retains Python extraction if AI fails.
- Do not upload confidential PDFs unless your institution allows sending them to Google Gemini.
- This is a public no-login prototype; anyone with the URL can use the website and, if Gemini is enabled, consume your API quota. Add access controls before wide distribution.
