# Mitsumori Streamlit App (Render Deploy)

This repository contains a Streamlit application for parsing PDF/Excel files and writing results into a template workbook. The existing application code remains unchanged. Additional files have been added to simplify deploying the app to [Render](https://render.com/).

## Deployment on Render
1. Push this repository to GitHub (private or public).
2. In Render, create a **Web Service** using the "+ New" button and select this repo.
3. Render will automatically pick up `render.yaml` as a blueprint. If not, choose:
   - Build command: `pip install -r requirements.txt`
   - Start command: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
4. Ensure the template file `template_output.xlsx` is present in the project root so the app can write results.

### Files added for Render
- `render.yaml`: Blueprint defining the Python web service, build, and start commands.
- `Procfile`: Alternative process declaration for environments that read Heroku-style process files (Render will also accept it).

An environment variable `SECRET_KEY` is generated automatically through the blueprint for compatibility with Flask-style settings, though the current Streamlit app does not require it.
