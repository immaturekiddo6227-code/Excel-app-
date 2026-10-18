# Excel AI Studio

Streamlit app: upload Excel/CSV, edit, clean, chart, analyse, and give natural-language instructions (Gemini API optional).

## Run locally

```
pip install -r requirements.txt
streamlit run app.py
```

Optional AI key: copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and add your `GEMINI_API_KEY`.

## Deploy on Streamlit Community Cloud

1. Create a GitHub repo and push these files (`app.py`, `requirements.txt`, `.streamlit/config.toml`, `.gitignore`, `README.md`). Do NOT push `secrets.toml`.
2. Go to https://share.streamlit.io and sign in with GitHub.
3. Click **Create app** -> pick the repo, branch `main`, main file `app.py`.
4. Open **Advanced settings** -> **Secrets** and paste: `GEMINI_API_KEY = "your-key"`.
5. Click **Deploy**. You get a public `*.streamlit.app` link.
