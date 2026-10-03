# Legend Engineering — Letter Processor & Response Drafter

A Streamlit application for:

1. Uploading incoming PDF correspondence
2. Extracting selectable PDF text
3. Generating an executive summary
4. Entering the company's response decision
5. Drafting a formal corporate response letter
6. Editing the draft
7. Downloading a professionally formatted Word document

## Important

This application requires an OpenAI API key.

Do NOT put the API key directly into `app.py`.

## Deploy on Streamlit Community Cloud

### 1. Create a GitHub repository

Upload:

- `app.py`
- `requirements.txt`
- `README.md`

### 2. Deploy

Open Streamlit Community Cloud and create an app using:

- Repository: your GitHub repository
- Branch: `main`
- Main file: `app.py`

### 3. Add the API key

In the Streamlit app settings, open **Secrets** and add:

```toml
OPENAI_API_KEY = "YOUR_REAL_OPENAI_API_KEY"
```

Save and restart/redeploy the application.

## Local testing

Install:

```bash
pip install -r requirements.txt
```

Then:

```bash
streamlit run app.py
```

For local secrets, create:

`.streamlit/secrets.toml`

with:

```toml
OPENAI_API_KEY = "YOUR_REAL_OPENAI_API_KEY"
```

Never commit that file to GitHub.

## PDF limitation

The current version extracts selectable text from PDFs. Scanned/image-only PDFs require OCR and are intentionally reported rather than silently producing an empty summary.

## Security

The API key is read from Streamlit Secrets or the environment. It is not hard-coded into the application.
