# Mitsumori Streamlit App (Render Deploy)

This repository contains a Streamlit application for parsing PDF/Excel files and writing results into a template workbook. The existing application code remains unchanged. Additional files have been added to simplify deploying the app to [Render](https://render.com/).

## Deployment on Render
1. Push this repository to GitHub (private or public).
2. In Render, create a **Web Service** using the "+ New" button and select this repo.
3. Render will automatically pick up `render.yaml` as a blueprint. If not, choose:
   - Build command: `apt-get update && apt-get install -y poppler-utils && pip install -r requirements.txt`
   - Start command: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
   - Python version: 3.11 (set via `pythonVersion` in `render.yaml` or `runtime.txt`)
4. Ensure the template file `template_output.xlsx` is present in the project root so the app can write results.

### Files added for Render
- `render.yaml`: Blueprint defining the Python web service, build (including Poppler) and start commands.
- `Procfile`: Alternative process declaration for environments that read Heroku-style process files (Render will also accept it).
- `runtime.txt`: Explicitly pins Python 3.11 so packages like `pandas` install from prebuilt wheels instead of compiling on newer
  interpreters. `pandas` is also pinned to 2.2.3, which ships Python 3.13 wheels in case Render provisions a newer runtime.

The blueprint installs `poppler-utils` so `pdf2image` can convert PDF pages during OCR. An environment variable `SECRET_KEY` is generated automatically through the blueprint for compatibility with Flask-style settings, though the current Streamlit app does not require it.
