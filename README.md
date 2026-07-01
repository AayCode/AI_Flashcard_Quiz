# AI Flashcard & MCQ Generator

A desktop application that automatically generates multiple-choice quizzes from your documents using NLP techniques — no LLM API required.

Upload a PDF, Word doc, PowerPoint, or plain text file, and the app extracts key terms using POS tagging, generates wrong-answer options using WordNet's lexical relationships, and creates an interactive flashcard quiz.

---

## Features

- **Multi-format document support** — PDF, DOCX, PPTX, TXT, and legacy DOC files
- **NLP-based question generation** — Uses NLTK POS tagging to identify important nouns as answer keywords
- **WordNet distractor generation** — Wrong answer options are generated from synonyms, hyponyms, and sibling terms — not random words
- **Interactive quiz mode** — Flashcard-style UI with color-coded feedback, score tracking, and animated results
- **PDF export** — Export your generated question set (questions + options + answers) as a clean PDF
- **Fully offline** — Runs on your local machine, no API keys needed

---

## Tech Stack

| Component | Library |
|---|---|
| GUI | `Tkinter` (Python built-in) |
| NLP / POS Tagging | `NLTK` |
| Distractor Generation | `WordNet` (via NLTK) |
| PDF Export | `ReportLab` |
| Document Parsing | `PyPDF2`, `python-docx`, `python-pptx` |

---

## How It Works

1. **Text Extraction** — The app reads your uploaded file and extracts raw text (handles PDF, DOCX, PPTX, TXT)
2. **Sentence Tokenization** — Text is split into individual sentences using NLTK
3. **Keyword Selection** — For each sentence, NLTK POS tagging identifies the most important noun (prioritizing longer, non-stopword nouns)
4. **Distractor Generation** — WordNet finds semantically related words (synonyms, hyponyms, sibling terms from shared hypernyms) to use as wrong answer options
5. **MCQ Assembly** — The keyword is blanked out in the sentence; 3 distractors + 1 correct answer are shuffled into a 4-option question
6. **Quiz / Export** — Play the quiz in the app or export all questions to PDF

---

## Setup

### 1. Clone the repository

```bash
git clone https://github.com/Aaycode/ai-flashcard-generator.git
cd ai-flashcard-generator
```

### 2. Install dependencies

```bash
pip install nltk reportlab PyPDF2 python-docx python-pptx
```

### 3. Run the app

```bash
python flashcard_mcq_app.py
```

NLTK data (punkt, wordnet, stopwords, pos tagger) downloads automatically on first run.

---

## Usage

1. Paste text directly into the text box, or click **Open File** to upload a document
2. Set the number of MCQs you want using the spinner
3. Click **Generate MCQs** to create questions from your content
4. Click **Play Quiz** to start the interactive flashcard quiz
5. Click **Export PDF** to save all questions and answers as a PDF file

---

## Known Limitations

- Question quality depends on the text — content-heavy academic or factual text works best; narrative or conversational text produces weaker questions
- WordNet distractors are lexically related but may occasionally be too similar or too unrelated to the correct answer
- Sentences where no suitable noun is found (or where WordNet has fewer than 3 distractors) are skipped
- Does not use an LLM — questions are rule-based, not semantically reasoned

---

## Requirements

```
nltk
reportlab
PyPDF2
python-docx
python-pptx
```

> Note: `textract` is an optional dependency for legacy `.doc` files. It requires additional system libraries and is not installed by default.

---

## License

MIT
