# StudyFlow — Document Study Assistant

StudyFlow helps students turn uploaded study materials into summaries, answers, and quizzes. It combines document processing and image OCR with a browser-based study interface.

## Try the Project

- **Live demo:** https://studyflow-q0ib.onrender.com/
- **Source code:** https://github.com/Ali27-rex/studyflow

The free demo may take around a minute to wake up after inactivity. Accounts and uploaded documents may be cleared when the hosting service restarts. Please use sample documents when testing.

## Features

- User registration and login with hashed passwords
- Upload PDF, DOCX, EPUB, TXT, and image files
- Extract text from images using Tesseract OCR
- Generate extractive summaries
- Ask questions about uploaded documents
- Generate multiple-choice quizzes
- View saved documents and chat history
- Delete uploaded documents

## Technology Stack

| Technology | Purpose |
|---|---|
| Python and Flask | Backend and application routes |
| SQLite | Accounts, documents, quizzes, and chat history |
| Werkzeug | Password hashing and filename handling |
| PyPDF2 | PDF text extraction |
| python-docx | Word document processing |
| Pillow and pytesseract | Image processing and OCR |
| ebooklib | EPUB processing |
| BeautifulSoup and lxml | Content parsing |
| requests | External API request support |
| HTML, CSS, and JavaScript | User interface |
| Docker and Gunicorn | Deployment and application serving |
| Render | Live demo hosting |

## Run Locally

Install Python, download or clone the repository, and open a terminal in the project folder.

Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

Start the application:

```powershell
python app.py
```

Open http://127.0.0.1:5000 in your browser.

### Image OCR on Windows

Install Tesseract from:
https://github.com/UB-Mannheim/tesseract/wiki

If Tesseract is not available in your PATH, set its location before starting the app:

```powershell
$env:TESSERACT_CMD = "C:\Program Files\Tesseract-OCR\tesseract.exe"
python app.py
```

## Deployment

The included Dockerfile installs the Python dependencies and Tesseract, then serves the application through Gunicorn.

For HTTPS hosting, set:

- `SECRET_KEY`: a stable, private random value
- `COOKIE_SECURE`: `true`

Persistent hosting requires storage for the database and uploaded files. Configure `DATA_DIR` to point to the persistent storage directory.

## Current Limitations

- Summaries and answers use extractive and keyword-based methods in local mode; they can be incomplete or inaccurate.
- PDF processing extracts embedded text. Scanned PDFs do not currently use OCR.
- Free demo storage is temporary.
- The application needs further security and reliability improvements before use with sensitive documents.
