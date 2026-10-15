# Deploy the EPD Extractor website (Streamlit Community Cloud)

1. Create a GitHub repository, and upload the files in this ZIP (not the ZIP itself). Do NOT upload any secrets, SQLite databases or AI caches.
2. Visit https://share.streamlit.io/ and sign in with GitHub.
3. Select **Create app** / **Deploy an app**, then choose your GitHub repository, branch `main`, and entrypoint `streamlit_app.py`.
4. In **Advanced settings → Secrets**, enter:

   GEMINI_API_KEY = "paste-your-own-key-here"

5. Deploy. Wait for dependencies to install and open the resulting URL.
6. Test with one PDF and download the Excel. Share the URL with your researchers.

## Notes
- The existing Python extraction and Gemini logic are reused, but the desktop Tkinter interface is replaced with Streamlit.
- The Gemini API key is stored in Streamlit Cloud Secrets, not GitHub or the website UI.
- PDFs are sent to Google when Gemini is enabled. Confirm your institutional document handling rules before using private PDFs.
- Records persist only within the current browser session; download the Excel before closing or refreshing the session. A permanent shared database is not implemented.
- The free host may sleep or have memory/CPU limits. Gemini quota, PDF size, and OCR still affect speed.
- OCR of scanned pages requires the `tesseract` executable installed on the host. Streamlit Community Cloud can use a `packages.txt` file for Linux packages if OCR is needed.
