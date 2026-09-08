# OneNote AI Tutor & Exam Prep

An automated AI-powered study assistant that connects to your Microsoft OneNote notebooks via Microsoft Graph, performs OCR on lecture slides, analyzes past midterms and finals for exam relevance, generates detailed explanations, and appends interactive practice quizzes directly to your course notes.

---

## Features
- **Microsoft Graph Integration**: Secure OAuth2 authentication supporting personal and school/work Microsoft accounts.
- **OCR Engine**: Fast text extraction from slide images using PyTesseract.
- **Syllabus & Exam Context Analysis**: Scans course folders for past midterms and finals to determine high-yield exam themes and weighting.
- **Gemini 2.5 Flash AI Tutor**: Automatically generates detailed explanations and midterm/final relevance breakdowns after every slide.
- **Automated Practice Quizzes**: Appends a custom multiple-choice quiz page at the end of each lecture section.
- **Modern Tailwind UI**: Gorgeous dark-mode dashboard with loading spinners and responsive card layouts.

---

## Project Structure
```text
├── bot.py                # Main Flask web application & auth router
├── ai_pipeline.py        # Core OCR, Gemini AI, and OneNote page patch engine
├── templates/
│   └── index.html        # Tailwind CSS responsive dashboard
├── requirements.txt      # Python dependencies
├── .env.example          # Environment variable template
└── .gitignore            # Git ignore rules
```

---

## Setup & Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/faraaz-bot/onenote-ai-tutor.git
   cd onenote-ai-tutor
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure environment variables:**
   Copy `.env.example` to `.env` and fill in your credentials:
   ```bash
   cp .env.example .env
   ```
   - `AZURE_CLIENT_ID`
   - `AZURE_CLIENT_SECRET`
   - `GEMINI_API_KEY`

4. **Run the application:**
   ```bash
   python3 bot.py
   ```
   Open `http://localhost:8000` in your browser.
