# StudyFlow
Flask study assistant for extracting notes, summaries, document questions, and quizzes.

## Upload to GitHub
1. Extract this ZIP and open the studyflow folder.
2. Create a private GitHub repository named studyflow.
3. Use Add file > Upload files. Upload the contents of studyflow, including Dockerfile and .gitignore, at the repository root. Do not upload the ZIP itself.
4. Commit the files. GitHub stores code; GitHub Pages cannot run this Flask application.

## Run on Windows
Install Python, then run in PowerShell from this folder:
```powershell
python -m pip install -r requirements.txt
python app.py
```
Open http://127.0.0.1:5000.
Image OCR also requires the Tesseract Windows program: https://github.com/UB-Mannheim/tesseract/wiki
If needed, before starting the app:
```powershell
$env:TESSERACT_CMD = "C:\Program Files\Tesseract-OCR\tesseract.exe"
```
PDF extraction reads embedded text; scanned PDFs are not OCR-enabled in this version.
The optional .env.example documents variables; the app does not automatically load .env files.

## Render Docker deployment
1. Connect the repository to a new Render Web Service.
2. Select Docker as the runtime. The Dockerfile installs Tesseract automatically.
3. Set SECRET_KEY to a long random value and COOKIE_SECURE to true in the service environment settings.
4. Deploy. Database tables initialize on startup, including when Gunicorn imports the app.
5. Test registration, TXT upload, summary, chat, quiz, and image OCR.

Free Render hosting is suitable for a demo only: its local database and uploads can disappear on restart, spin-down, or redeployment. It also sleeps when idle. This package does not configure external storage.
For persistent operation on a paid Render service, attach a disk at /app/data and set DATA_DIR=/app/data. Back up that directory separately.
Do not point your main domain at this demo if it already serves another website. Connect a separate subdomain only after deployment is working.

## Data and limitations
This archive excludes the supplied database, uploaded image, and editor configuration. New accounts start with a clean database. Files are stored in data/uploads outside the public static directory. Failed extraction is reported as an error rather than saved as successful content.
Use a stable SECRET_KEY for hosting; without one, restarting the app invalidates sessions. Local mode uses extractive heuristics, not an external large language model. This is a starter project; public use with sensitive documents needs further review, including CSRF protection, upload validation, resource limits, and backups.
