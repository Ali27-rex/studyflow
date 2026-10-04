from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import sqlite3
import os
import requests
import re
import random
from collections import Counter
from datetime import datetime
from functools import wraps
import PyPDF2

# Multi-format support imports
try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import pytesseract
    from PIL import Image
    PILLOW_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False

try:
    import ebooklib
    from ebooklib import epub
    from bs4 import BeautifulSoup
    EPUB_AVAILABLE = True
except ImportError:
    EPUB_AVAILABLE = False

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY') or os.urandom(32)
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['SESSION_COOKIE_SECURE'] = os.environ.get('COOKIE_SECURE', 'false').lower() == 'true'
DATA_DIR = os.path.abspath(os.environ.get('DATA_DIR', os.path.join(app.root_path, 'data')))
app.config['UPLOAD_FOLDER'] = os.path.join(DATA_DIR, 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024
app.config['DATABASE'] = os.path.join(DATA_DIR, 'study_assistant.db')
if PILLOW_AVAILABLE and os.environ.get('TESSERACT_CMD'):
    pytesseract.pytesseract.tesseract_cmd = os.environ['TESSERACT_CMD']

# MULTI-FORMAT SUPPORT
ALLOWED_EXTENSIONS = {
    'pdf', 'docx', 'doc', 'txt', 'epub',
    'png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'tiff'
}

# File type icons mapping
FILE_ICONS = {
    'pdf': 'fa-file-pdf',
    'docx': 'fa-file-word',
    'doc': 'fa-file-word',
    'txt': 'fa-file-lines',
    'epub': 'fa-book',
    'png': 'fa-file-image',
    'jpg': 'fa-file-image',
    'jpeg': 'fa-file-image',
    'gif': 'fa-file-image',
    'bmp': 'fa-file-image',
    'webp': 'fa-file-image',
    'tiff': 'fa-file-image',
}

FILE_COLORS = {
    'pdf': '#ef4444',
    'docx': '#3b82f6',
    'doc': '#3b82f6',
    'txt': '#6b7280',
    'epub': '#8b5cf6',
    'png': '#10b981',
    'jpg': '#10b981',
    'jpeg': '#10b981',
    'gif': '#10b981',
    'bmp': '#10b981',
    'webp': '#10b981',
    'tiff': '#10b981',
}

# AI CONFIGURATION
USE_LOCAL_MODE = True
HF_API_TOKEN = 'hf_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
HF_API_URL = 'https://api-inference.huggingface.co/models/sshleifer/distilbart-cnn-12-6'


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_file_extension(filename):
    return filename.rsplit('.', 1)[1].lower() if '.' in filename else ''


def get_file_icon(filename):
    ext = get_file_extension(filename)
    return FILE_ICONS.get(ext, 'fa-file')


def get_file_color(filename):
    ext = get_file_extension(filename)
    return FILE_COLORS.get(ext, '#6b7280')


def get_db():
    conn = sqlite3.connect(app.config['DATABASE'])
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    os.makedirs(os.path.dirname(app.config['DATABASE']) or '.', exist_ok=True)
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL, email TEXT UNIQUE NOT NULL, password TEXT NOT NULL, full_name TEXT, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)")
    cursor.execute("CREATE TABLE IF NOT EXISTS documents (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL, filename TEXT NOT NULL, original_name TEXT NOT NULL, file_path TEXT NOT NULL, file_type TEXT DEFAULT 'pdf', extracted_text TEXT, summary TEXT, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, FOREIGN KEY (user_id) REFERENCES users (id))")
    cursor.execute("CREATE TABLE IF NOT EXISTS quizzes (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL, document_id INTEGER, title TEXT NOT NULL, questions TEXT NOT NULL, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, FOREIGN KEY (user_id) REFERENCES users (id), FOREIGN KEY (document_id) REFERENCES documents (id))")
    cursor.execute("CREATE TABLE IF NOT EXISTS chat_history (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL, document_id INTEGER, question TEXT NOT NULL, answer TEXT NOT NULL, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, FOREIGN KEY (user_id) REFERENCES users (id), FOREIGN KEY (document_id) REFERENCES documents (id))")
    conn.commit()
    conn.close()


def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            flash('Please login first!', 'warning')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function


# ==================== MULTI-FORMAT TEXT EXTRACTION ====================

def extract_text_from_pdf(file_path):
    try:
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ''
            for page in pdf_reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text = text + page_text + chr(10)
            return text.strip()
    except Exception as e:
        return 'Error extracting text from PDF: ' + str(e)


def extract_text_from_docx(file_path):
    if not DOCX_AVAILABLE:
        return 'Error: python-docx not installed. Run: pip install python-docx'
    try:
        doc = Document(file_path)
        paragraphs = []
        for para in doc.paragraphs:
            if para.text.strip():
                paragraphs.append(para.text.strip())
        for table in doc.tables:
            for row in table.rows:
                row_text = []
                for cell in row.cells:
                    if cell.text.strip():
                        row_text.append(cell.text.strip())
                if row_text:
                    paragraphs.append(' | '.join(row_text))
        return chr(10).join(paragraphs)
    except Exception as e:
        return 'Error extracting text from DOCX: ' + str(e)


def extract_text_from_txt(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            return file.read()
    except UnicodeDecodeError:
        try:
            with open(file_path, 'r', encoding='latin-1') as file:
                return file.read()
        except Exception as e:
            return 'Error reading text file: ' + str(e)
    except Exception as e:
        return 'Error reading text file: ' + str(e)


def extract_text_from_image(file_path):
    if not PILLOW_AVAILABLE:
        return 'Error: OCR not available. Install: pip install pytesseract pillow'
    try:
        image = Image.open(file_path)
        if image.mode not in ('L', 'RGB'):
            image = image.convert('RGB')
        text = pytesseract.image_to_string(image, lang='eng')
        if not text.strip():
            return 'No text detected in image. The image may not contain readable text, or OCR may need configuration.'
        return text.strip()
    except Exception as e:
        return 'Error performing OCR on image: ' + str(e)


def extract_text_from_epub(file_path):
    if not EPUB_AVAILABLE:
        return 'Error: EPUB support not installed. Run: pip install ebooklib beautifulsoup4'
    try:
        book = epub.read_epub(file_path)
        texts = []
        for item in book.get_items():
            if item.get_type() == ebooklib.ITEM_DOCUMENT:
                soup = BeautifulSoup(item.get_content(), 'html.parser')
                for script in soup(["script", "style"]):
                    script.decompose()
                text = soup.get_text(separator='\n')
                lines = [line.strip() for line in text.splitlines() if line.strip()]
                if lines:
                    texts.append(chr(10).join(lines))
        return chr(10).join(texts)
    except Exception as e:
        return 'Error extracting text from EPUB: ' + str(e)


def extract_text(file_path, file_type):
    extractors = {
        'pdf': extract_text_from_pdf,
        'docx': extract_text_from_docx,
        'doc': extract_text_from_docx,
        'txt': extract_text_from_txt,
        'epub': extract_text_from_epub,
        'png': extract_text_from_image,
        'jpg': extract_text_from_image,
        'jpeg': extract_text_from_image,
        'gif': extract_text_from_image,
        'bmp': extract_text_from_image,
        'webp': extract_text_from_image,
        'tiff': extract_text_from_image,
    }
    extractor = extractors.get(file_type, extract_text_from_txt)
    return extractor(file_path)


# ==================== LOCAL AI ENGINE ====================

def get_local_summary(text, doc_title='Document'):
    sentences = re.split(r'[.!?]+', text)
    sentences = [s.strip() for s in sentences if len(s.strip()) > 20]

    if not sentences:
        return '## Summary of ' + doc_title + chr(10) + chr(10) + 'No readable text found. The file may be scanned images or corrupted.'

    important_words = ['definition', 'concept', 'important', 'key', 'main', 'primary', 'essential', 'fundamental',
                       'virtualization', 'virtual', 'hypervisor', 'vm', 'machine', 'cloud', 'network', 'server', 'system',
                       'architecture', 'protocol', 'algorithm', 'database', 'security', 'container', 'docker', 'kubernetes']

    key_sentences = []
    for i, sent in enumerate(sentences[:50]):
        sent_lower = sent.lower()
        score = 0
        if i == 0:
            score += 3
        for word in important_words:
            if word in sent_lower:
                score += 2
        if 50 < len(sent) < 200:
            score += 1
        if score >= 3:
            key_sentences.append((score, sent))

    key_sentences.sort(reverse=True)
    top_sentences = [s for _, s in key_sentences[:8]]

    words = re.findall(r'[A-Z][a-zA-Z\-]{2,}', text)
    word_counts = Counter(words)
    key_terms = [word for word, count in word_counts.most_common(15) if count >= 2 and len(word) > 3][:6]

    sections = []
    heading_pattern = re.findall(r'(?:^|\n)([A-Z][A-Z\s\-a-z]{3,50})(?:\n|:)', text)
    for heading in heading_pattern[:5]:
        clean = heading.strip()
        if 3 < len(clean) < 60:
            sections.append(clean)

    parts = ['## AI Summary of ' + doc_title, '']
    parts.append('This summary was generated by analyzing the actual content of your document.')
    parts.append('')

    if key_terms:
        parts.append('### Key Terms & Concepts')
        for term in key_terms:
            found = False
            for sent in sentences:
                if term.lower() in sent.lower() and len(sent) < 250:
                    parts.append('- ' + term + ': ' + sent.strip())
                    found = True
                    break
            if not found:
                parts.append('- ' + term + ': Mentioned throughout the document.')
        parts.append('')

    if sections:
        parts.append('### Document Sections')
        for section in sections:
            parts.append('- ' + section)
        parts.append('')

    parts.append('### Key Points from Your Document')
    for i, sent in enumerate(top_sentences[:6], 1):
        parts.append(str(i) + '. ' + sent)

    parts.append('')
    parts.append('### Study Recommendations')
    parts.append('- Focus on the key terms and concepts listed above')
    parts.append('- Review the document sections for detailed understanding')
    parts.append('- Practice explaining the main concepts in your own words')
    parts.append('- Look for connections between different sections')

    return chr(10).join(parts)


def get_local_quiz(text, num_questions=5):
    sentences = re.split(r'[.!?]+', text)
    sentences = [s.strip() for s in sentences if 30 < len(s.strip()) < 200]

    if len(sentences) < 5:
        return 'Not enough content to generate quiz.'

    words = re.findall(r'[A-Z][a-zA-Z]{2,}', text)
    word_counts = Counter(words)
    key_terms = [w for w, c in word_counts.most_common(20) if c >= 2 and len(w) > 3 and w.isalpha()][:10]

    important_sentences = []
    for sent in sentences:
        for term in key_terms:
            if term.lower() in sent.lower():
                important_sentences.append(sent)
                break

    if not important_sentences:
        important_sentences = sentences[:20]

    random.shuffle(important_sentences)

    quiz_parts = []
    q_count = 0
    for sent in important_sentences:
        if q_count >= num_questions:
            break
        words_in_sent = sent.split()
        if len(words_in_sent) < 5:
            continue
        question_made = False
        for term in key_terms:
            if term.lower() in sent.lower() and len(term) > 3:
                question_text = sent.replace(term, '_______').replace(term.lower(), '_______')
                if question_text != sent:
                    wrong_options = [w for w in key_terms if w.lower() != term.lower() and len(w) > 3][:3]
                    while len(wrong_options) < 3:
                        wrong_options.append('Concept ' + str(len(wrong_options)+1))
                    options = [term] + wrong_options[:3]
                    random.shuffle(options)
                    correct_letter = chr(65 + options.index(term))
                    quiz_parts.append('Q' + str(q_count+1) + ': ' + question_text)
                    for j, opt in enumerate(options):
                        quiz_parts.append(chr(65+j) + ') ' + opt)
                    quiz_parts.append('Correct Answer: ' + correct_letter)
                    quiz_parts.append('')
                    q_count += 1
                    question_made = True
                    break
        if not question_made and q_count < num_questions:
            quiz_parts.append('Q' + str(q_count+1) + ': Based on the document, which statement is accurate?')
            quiz_parts.append('A) ' + sent[:100] + '...')
            quiz_parts.append('B) This is not mentioned in the document')
            quiz_parts.append('C) The opposite of what the text describes')
            quiz_parts.append('D) A related but incorrect statement')
            quiz_parts.append('Correct Answer: A')
            quiz_parts.append('')
            q_count += 1

    return chr(10).join(quiz_parts)


def get_local_answer(text, question):
    """
    Enhanced Q&A engine with:
    - Semantic keyword weighting (nouns/verbs weighted higher)
    - Sentence context scoring (surrounding sentences add context)
    - Question type detection (what/why/how/where/when/who)
    - Synonym expansion for better matching
    - Confidence scoring with fallback explanations
    """

    question_lower = question.lower().strip()
    if not question_lower or not text or len(text.strip()) < 50:
        return "I don't have enough document content to answer questions. Please upload a document with more text."

    # Question type detection
    question_types = {
        'what': ['definition', 'meaning', 'is', 'are', 'refers to', 'called'],
        'why': ['because', 'reason', 'due to', 'causes', 'leads to', 'results in'],
        'how': ['by', 'through', 'using', 'method', 'process', 'steps', 'way to'],
        'where': ['in', 'at', 'located', 'place', 'region', 'area', 'site'],
        'when': ['in', 'during', 'at', 'time', 'period', 'year', 'date', 'after', 'before'],
        'who': ['by', 'developed', 'created', 'founded', 'discovered', 'invented'],
        'which': ['the', 'a', 'an', 'called', 'known as', 'referred to'],
        'define': ['is', 'are', 'refers to', 'means', 'defined as'],
        'explain': ['is', 'are', 'means', 'refers to', 'involves', 'consists of'],
        'list': ['include', 'such as', 'for example', 'are', 'consists of'],
        'compare': ['versus', 'vs', 'unlike', 'similar', 'difference', 'both', 'while'],
        'difference': ['unlike', 'different', 'whereas', 'but', 'however', 'unlike'],
    }

    detected_type = None
    type_boost_words = []
    for qtype, boost in question_types.items():
        if qtype in question_lower or any(qtype in w for w in question_lower.split()):
            detected_type = qtype
            type_boost_words = boost
            break

    # Extract important keywords from question (nouns, verbs, key terms)
    # Remove common stop words
    stop_words = {'the', 'a', 'an', 'is', 'are', 'was', 'were', 'be', 'been', 'being',
                  'have', 'has', 'had', 'do', 'does', 'did', 'will', 'would', 'could',
                  'should', 'may', 'might', 'must', 'shall', 'can', 'need', 'dare',
                  'ought', 'used', 'to', 'of', 'in', 'for', 'on', 'with', 'at', 'by',
                  'from', 'as', 'into', 'through', 'during', 'before', 'after', 'above',
                  'below', 'between', 'under', 'again', 'further', 'then', 'once',
                  'here', 'there', 'when', 'where', 'why', 'how', 'all', 'each', 'few',
                  'more', 'most', 'other', 'some', 'such', 'no', 'nor', 'not', 'only',
                  'own', 'same', 'so', 'than', 'too', 'very', 'just', 'and', 'but',
                  'if', 'or', 'because', 'until', 'while', 'this', 'that', 'these',
                  'those', 'i', 'me', 'my', 'myself', 'we', 'our', 'ours', 'ourselves',
                  'you', 'your', 'yours', 'yourself', 'yourselves', 'he', 'him', 'his',
                  'himself', 'she', 'her', 'hers', 'herself', 'it', 'its', 'itself',
                  'they', 'them', 'their', 'theirs', 'themselves', 'what', 'which',
                  'who', 'whom', 'whose', 'where', 'when', 'why', 'how', 'about',
                  'does', 'did', 'define', 'explain', 'tell', 'describe', 'give',
                  'mention', 'state', 'identify', 'name', 'find', 'show', 'list',
                  'compare', 'contrast', 'difference', 'between', 'among', 'me', 'tell me',
                  'can you', 'could you', 'please', 'would you', 'will you'}

    q_words_raw = re.findall(r'[a-zA-Z\-]{3,}', question_lower)
    q_keywords = []
    for word in q_words_raw:
        if word not in stop_words and len(word) > 3:
            q_keywords.append(word)

    # Add important bigrams (two-word phrases)
    q_bigrams = []
    words_list = question_lower.split()
    for i in range(len(words_list) - 1):
        bigram = words_list[i] + ' ' + words_list[i+1]
        if len(bigram) > 6:
            q_bigrams.append(bigram)

    # Synonym expansion for better matching
    synonyms = {
        'virtualization': ['virtual', 'vm', 'virtual machine', 'hypervisor'],
        'hypervisor': ['virtualization', 'vm monitor', 'virtual machine monitor'],
        'cloud': ['cloud computing', 'aws', 'azure', 'gcp', 'saas', 'paas', 'iaas'],
        'network': ['networking', 'protocol', 'tcp', 'ip', 'lan', 'wan', 'internet'],
        'database': ['db', 'sql', 'nosql', 'data store', 'data storage'],
        'security': ['cybersecurity', 'protection', 'encryption', 'firewall', 'auth'],
        'algorithm': ['algorithms', 'procedure', 'method', 'technique', 'process'],
        'machine learning': ['ml', 'ai', 'artificial intelligence', 'deep learning', 'neural network'],
        'ai': ['artificial intelligence', 'machine learning', 'deep learning', 'neural network'],
        'container': ['docker', 'kubernetes', 'k8s', 'containerization', 'pod'],
        'docker': ['container', 'containerization', 'image', 'dockerfile'],
        'programming': ['coding', 'development', 'software', 'code', 'script'],
        'function': ['method', 'procedure', 'routine', 'subroutine'],
        'variable': ['var', 'parameter', 'argument', 'constant', 'value'],
        'class': ['object', 'oop', 'inheritance', 'polymorphism', 'encapsulation'],
        'inheritance': ['class', 'parent', 'child', 'superclass', 'subclass'],
        'polymorphism': ['overloading', 'overriding', 'interface', 'abstraction'],
        'encapsulation': ['data hiding', 'private', 'protected', 'access modifier'],
        'abstraction': ['abstract', 'interface', 'hide complexity', 'simplify'],
        'recursion': ['recursive', 'base case', 'self-referential', 'divide and conquer'],
        'sorting': ['sort', 'order', 'arrange', 'bubble sort', 'quick sort', 'merge sort'],
        'searching': ['search', 'find', 'lookup', 'binary search', 'linear search'],
        'stack': ['lifo', 'push', 'pop', 'last in first out'],
        'queue': ['fifo', 'enqueue', 'dequeue', 'first in first out'],
        'tree': ['binary tree', 'bst', 'heap', 'trie', 'graph', 'node'],
        'graph': ['vertex', 'edge', 'node', 'path', 'dfs', 'bfs', 'dijkstra'],
        'memory': ['ram', 'storage', 'cache', 'buffer', 'heap', 'stack'],
        'cpu': ['processor', 'central processing unit', 'core', 'thread', 'clock'],
        'os': ['operating system', 'kernel', 'process', 'thread', 'scheduler'],
        'process': ['program', 'execution', 'task', 'job', 'instance'],
        'thread': ['lightweight process', 'concurrency', 'parallelism'],
        'deadlock': ['circular wait', 'resource allocation', 'starvation', 'livelock'],
        'semaphore': ['mutex', 'lock', 'synchronization', 'critical section'],
        'paging': ['page', 'virtual memory', 'page fault', 'page table'],
        'segmentation': ['segment', 'memory management', 'fragmentation'],
        'file': ['directory', 'folder', 'path', 'filesystem', 'storage'],
        'disk': ['hard drive', 'hdd', 'ssd', 'storage', 'secondary storage'],
    }

    expanded_keywords = list(q_keywords)
    for kw in q_keywords:
        if kw in synonyms:
            for syn in synonyms[kw]:
                if syn not in expanded_keywords:
                    expanded_keywords.append(syn)

    # Split text into sentences with better boundary detection
    sentences = re.split(r'(?<=[.!?])\s+', text)
    sentences = [s.strip() for s in sentences if len(s.strip()) > 15]

    if not sentences:
        return "I couldn't find readable text in this document. It might be a scanned image or have very little content."

    # Score each sentence
    scored_sentences = []
    for i, sent in enumerate(sentences):
        sent_lower = sent.lower()
        score = 0
        matched_keywords = []

        # Keyword matching with weights
        for kw in expanded_keywords:
            if kw in sent_lower:
                # Longer keyword matches = higher score
                kw_score = min(len(kw) * 0.5, 5)
                score += kw_score
                matched_keywords.append(kw)

        # Bigram matching (higher weight)
        for bg in q_bigrams:
            if bg in sent_lower:
                score += 4
                matched_keywords.append(bg)

        # Question type boost
        if detected_type and type_boost_words:
            for boost_word in type_boost_words:
                if boost_word in sent_lower:
                    score += 2

        # Definition pattern boost (for "what is X" questions)
        if detected_type in ('what', 'define', 'explain'):
            definition_patterns = [
                r'is\s+a\s+', r'is\s+an\s+', r'is\s+the\s+', r'refers\s+to',
                r'defined\s+as', r'means', r'consists\s+of', r'involves',
                r'characterized\s+by', r'known\s+as', r'called'
            ]
            for pattern in definition_patterns:
                if re.search(pattern, sent_lower):
                    score += 3
                    break

        # List pattern boost (for "list" or "what are" questions)
        if detected_type in ('list', 'what') and any(w in question_lower for w in ['list', 'examples', 'types', 'kinds']):
            list_patterns = [r'include', r'such\s+as', r'for\s+example', r'are:', r'following', r'consists\s+of']
            for pattern in list_patterns:
                if re.search(pattern, sent_lower):
                    score += 2
                    break

        # Proximity boost - sentences near other high-scoring sentences
        context_bonus = 0
        if i > 0:
            prev_sent = sentences[i-1].lower()
            for kw in expanded_keywords:
                if kw in prev_sent:
                    context_bonus += 0.5
        if i < len(sentences) - 1:
            next_sent = sentences[i+1].lower()
            for kw in expanded_keywords:
                if kw in next_sent:
                    context_bonus += 0.5
        score += context_bonus

        # Length penalty - prefer medium-length sentences
        if len(sent) < 40:
            score *= 0.7
        elif len(sent) > 300:
            score *= 0.8

        # First sentence of paragraph boost (often contains key info)
        if i == 0 or (i > 0 and sentences[i-1].endswith('.')):
            score *= 1.1

        scored_sentences.append({
            'sentence': sent,
            'score': score,
            'matched': matched_keywords,
            'index': i
        })

    # Sort by score
    scored_sentences.sort(key=lambda x: x['score'], reverse=True)

    # Filter out very low scores
    relevant = [s for s in scored_sentences if s['score'] >= 1.5]

    if not relevant:
        # Fallback: try with just the original keywords (no expansion)
        for i, sent in enumerate(sentences):
            sent_lower = sent.lower()
            score = 0
            for kw in q_keywords:
                if kw in sent_lower:
                    score += 1
            if score >= 1:
                relevant.append({
                    'sentence': sent,
                    'score': score,
                    'matched': [kw],
                    'index': i
                })

    if not relevant:
        # Last resort: return a helpful message with suggestions
        # Find any sentence containing ANY word from the question
        partial_matches = []
        for sent in sentences:
            sent_lower = sent.lower()
            for word in q_words_raw:
                if len(word) > 4 and word in sent_lower:
                    partial_matches.append(sent)
                    break

        if partial_matches:
            return ("I couldn't find a direct answer, but here are some related passages from your document:" + chr(10) + chr(10) +
                    chr(10).join(['- ' + s[:200] + ('...' if len(s) > 200 else '') for s in partial_matches[:3]]) + chr(10) + chr(10) +
                    "Try rephrasing your question with terms found in the document, or ask about specific concepts mentioned in the text.")
        else:
            # Extract key terms from document to suggest
            doc_words = re.findall(r'[A-Z][a-zA-Z]{3,}', text)
            word_counts = Counter(doc_words)
            top_terms = [w for w, c in word_counts.most_common(10) if c >= 2 and len(w) > 4][:5]

            suggestion = ""
            if top_terms:
                suggestion = " Try asking about: " + ", ".join(top_terms)

            return ("I couldn't find relevant information for that question in your document." + chr(10) + chr(10) +
                    "Tips for better results:" + chr(10) +
                    "- Use specific terms from the document" + chr(10) +
                    "- Ask about definitions, concepts, or key points" + chr(10) +
                    "- Try simpler, more direct questions" + suggestion)

    # Build the answer
    top_results = relevant[:4]

    # Deduplicate and sort by original order for coherence
    seen_indices = set()
    unique_results = []
    for r in top_results:
        if r['index'] not in seen_indices:
            seen_indices.add(r['index'])
            unique_results.append(r)

    # If we have multiple results, try to include context sentences
    final_sentences = []
    for r in unique_results[:3]:
        idx = r['index']
        # Add previous sentence for context if relevant
        if idx > 0 and (idx - 1) not in seen_indices:
            prev = sentences[idx - 1]
            if len(prev) > 30:
                final_sentences.append({'text': prev, 'type': 'context'})
        final_sentences.append({'text': r['sentence'], 'type': 'answer'})
        # Add next sentence for context
        if idx < len(sentences) - 1 and (idx + 1) not in seen_indices:
            nxt = sentences[idx + 1]
            if len(nxt) > 30:
                final_sentences.append({'text': nxt, 'type': 'context'})

    # Build formatted answer
    answer_parts = ["Here's what I found in your document:"]
    answer_parts.append("")

    for i, item in enumerate(final_sentences[:5]):
        prefix = "•" if item['type'] == 'answer' else "  "
        text = item['text'].strip()
        if text.endswith('.') or text.endswith('!') or text.endswith('?'):
            answer_parts.append(prefix + " " + text)
        else:
            answer_parts.append(prefix + " " + text + ".")

    answer_parts.append("")

    # Add confidence indicator
    best_score = unique_results[0]['score'] if unique_results else 0
    if best_score >= 8:
        answer_parts.append("✓ High confidence match — this directly answers your question.")
    elif best_score >= 4:
        answer_parts.append("~ Moderate confidence — this is the most relevant information I found.")
    else:
        answer_parts.append("? Low confidence — the answer may be partial. Try rephrasing your question.")

    return chr(10).join(answer_parts)


# ==================== API FUNCTIONS ====================

def call_hf_api(text, max_length=150):
    if HF_API_TOKEN == 'hf_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx':
        return None
    try:
        headers = {'Authorization': 'Bearer ' + HF_API_TOKEN}
        payload = {'inputs': text[:1000], 'parameters': {'max_length': max_length, 'min_length': 30}}
        response = requests.post(HF_API_URL, headers=headers, json=payload, timeout=30)
        if response.status_code == 200:
            result = response.json()
            if isinstance(result, list) and len(result) > 0:
                return result[0].get('summary_text', '')
        return None
    except:
        return None


def generate_summary(text, doc_title='Document'):
    if USE_LOCAL_MODE or HF_API_TOKEN == 'hf_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx':
        return get_local_summary(text, doc_title)
    hf_result = call_hf_api(text[:1000])
    if hf_result:
        return '## AI Summary' + chr(10) + chr(10) + hf_result + chr(10) + chr(10) + '_Generated by Hugging Face AI_'
    return get_local_summary(text, doc_title)


def generate_quiz(text, num_questions=5):
    return get_local_quiz(text, num_questions)


def get_suggested_questions(text, num_questions=5):
    """Suggest questions using recurring terms from the document."""
    if not text or num_questions <= 0:
        return []
    stop_words = {'this', 'that', 'these', 'those', 'with', 'from', 'have',
                  'will', 'there', 'their', 'which', 'about', 'they', 'were',
                  'been', 'into', 'also', 'when', 'what', 'where', 'more'}
    terms = re.findall(r'\b[a-zA-Z]{4,}\b', text.lower())
    counts = Counter(term for term in terms if term not in stop_words)
    return ['What does the document say about ' + term + '?'
            for term, _ in counts.most_common(num_questions)]


def answer_question(text, question, chat_history=None):
    return get_local_answer(text, question)


# ==================== ROUTES ====================

@app.route('/')
def index():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
    return render_template('index.html')


@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        username = request.form['username']
        email = request.form['email']
        password = request.form['password']
        full_name = request.form.get('full_name', '')
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM users WHERE username = ? OR email = ?', (username, email))
        if cursor.fetchone():
            flash('Username or email already exists!', 'danger')
            conn.close()
            return redirect(url_for('register'))
        hashed = generate_password_hash(password)
        cursor.execute('INSERT INTO users (username, email, password, full_name) VALUES (?, ?, ?, ?)',
                       (username, email, hashed, full_name))
        conn.commit()
        conn.close()
        flash('Registration successful! Please login.', 'success')
        return redirect(url_for('login'))
    return render_template('register.html')


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM users WHERE username = ?', (username,))
        user = cursor.fetchone()
        conn.close()
        if user and check_password_hash(user['password'], password):
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['full_name'] = user['full_name']
            flash('Login successful!', 'success')
            return redirect(url_for('dashboard'))
        flash('Invalid username or password!', 'danger')
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    flash('You have been logged out.', 'info')
    return redirect(url_for('index'))


@app.route('/dashboard')
@login_required
def dashboard():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM documents WHERE user_id = ? ORDER BY created_at DESC', (session['user_id'],))
    documents = cursor.fetchall()
    cursor.execute('SELECT q.*, d.original_name as doc_name FROM quizzes q LEFT JOIN documents d ON q.document_id = d.id WHERE q.user_id = ? ORDER BY q.created_at DESC', (session['user_id'],))
    quizzes = cursor.fetchall()
    conn.close()
    return render_template('dashboard.html', documents=documents, quizzes=quizzes,
                         user_name=session.get('full_name', session['username']))


@app.route('/upload', methods=['GET', 'POST'])
@login_required
def upload():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file selected!', 'danger')
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            flash('No file selected!', 'danger')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            original = secure_filename(file.filename)
            file_type = get_file_extension(original)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = str(session['user_id']) + '_' + timestamp + '_' + original
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            extracted = extract_text(filepath, file_type)
            if not extracted or not extracted.strip() or extracted.startswith('Error'):
                os.remove(filepath)
                flash(extracted or 'No text could be extracted from this file.', 'danger')
                return redirect(request.url)
            summary = generate_summary(extracted, original)
            conn = get_db()
            cursor = conn.cursor()
            cursor.execute("INSERT INTO documents (user_id, filename, original_name, file_path, file_type, extracted_text, summary) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (session['user_id'], filename, original, filepath, file_type, extracted, summary))
            conn.commit()
            doc_id = cursor.lastrowid
            conn.close()
            file_type_label = file_type.upper() if file_type != 'epub' else 'EPUB'
            flash(file_type_label + ' uploaded and analyzed successfully!', 'success')
            return redirect(url_for('view_document', doc_id=doc_id))
        flash('Unsupported file type!', 'danger')
    return render_template('upload.html')


@app.route('/document/<int:doc_id>')
@login_required
def view_document(doc_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM documents WHERE id = ? AND user_id = ?', (doc_id, session['user_id']))
    doc = cursor.fetchone()
    conn.close()
    if not doc:
        flash('Document not found!', 'danger')
        return redirect(url_for('dashboard'))
    return render_template('document.html', document=doc,
                         file_icon=get_file_icon(doc['original_name']),
                         file_color=get_file_color(doc['original_name']))


@app.route('/document/<int:doc_id>/summary')
@login_required
def view_summary(doc_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM documents WHERE id = ? AND user_id = ?', (doc_id, session['user_id']))
    doc = cursor.fetchone()
    conn.close()
    if not doc:
        flash('Document not found!', 'danger')
        return redirect(url_for('dashboard'))
    return render_template('summary.html', document=doc)


@app.route('/document/<int:doc_id>/chat', methods=['GET', 'POST'])
@login_required
def chat(doc_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM documents WHERE id = ? AND user_id = ?', (doc_id, session['user_id']))
    doc = cursor.fetchone()
    cursor.execute('SELECT * FROM chat_history WHERE document_id = ? AND user_id = ? ORDER BY created_at ASC',
                   (doc_id, session['user_id']))
    history = cursor.fetchall()
    conn.close()
    if not doc:
        flash('Document not found!', 'danger')
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        question = request.form['question']
        # Get recent chat history for context
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('SELECT question, answer FROM chat_history WHERE document_id = ? AND user_id = ? ORDER BY created_at DESC LIMIT 5',
                       (doc_id, session['user_id']))
        recent_chat = cursor.fetchall()
        conn.close()

        answer = answer_question(doc['extracted_text'], question, recent_chat)

        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('INSERT INTO chat_history (user_id, document_id, question, answer) VALUES (?, ?, ?, ?)',
                       (session['user_id'], doc_id, question, answer))
        conn.commit()
        conn.close()
        return jsonify({'answer': answer})

    # Get suggested questions for this document
    suggested = get_suggested_questions(doc['extracted_text'], 5)
    return render_template('chat.html', document=doc, chat_history=history, suggested_questions=suggested)


@app.route('/document/<int:doc_id>/quiz', methods=['GET', 'POST'])
@login_required
def generate_quiz_route(doc_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM documents WHERE id = ? AND user_id = ?', (doc_id, session['user_id']))
    doc = cursor.fetchone()
    conn.close()
    if not doc:
        flash('Document not found!', 'danger')
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        num = int(request.form.get('num_questions', 5))
        content = generate_quiz(doc['extracted_text'], num)
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('INSERT INTO quizzes (user_id, document_id, title, questions) VALUES (?, ?, ?, ?)',
                       (session['user_id'], doc_id, 'Quiz for ' + doc['original_name'], content))
        conn.commit()
        quiz_id = cursor.lastrowid
        conn.close()
        flash('Quiz generated from your document!', 'success')
        return redirect(url_for('take_quiz', quiz_id=quiz_id))
    return render_template('generate_quiz.html', document=doc)


@app.route('/quiz/<int:quiz_id>')
@login_required
def take_quiz(quiz_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT q.*, d.original_name as doc_name FROM quizzes q LEFT JOIN documents d ON q.document_id = d.id WHERE q.id = ? AND q.user_id = ?', (quiz_id, session['user_id']))
    quiz = cursor.fetchone()
    conn.close()
    if not quiz:
        flash('Quiz not found!', 'danger')
        return redirect(url_for('dashboard'))
    questions = parse_quiz_questions(quiz['questions'])
    return render_template('quiz.html', quiz=quiz, questions=questions)


def parse_quiz_questions(quiz_text):
    questions = []
    lines = quiz_text.split(chr(10))
    current = None
    for line in lines:
        line = line.strip()
        if line.startswith('Q') and ':' in line:
            if current:
                questions.append(current)
            current = {'question': line.split(':', 1)[1].strip(), 'options': [], 'correct': ''}
        elif line.startswith(('A)', 'B)', 'C)', 'D)')) and current:
            current['options'].append(line)
        elif 'Correct Answer:' in line and current:
            current['correct'] = line.split(':')[1].strip()
    if current:
        questions.append(current)
    return questions


@app.route('/document/<int:doc_id>/delete', methods=['POST'])
@login_required
def delete_document(doc_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM documents WHERE id = ? AND user_id = ?', (doc_id, session['user_id']))
    doc = cursor.fetchone()
    if doc:
        try:
            os.remove(doc['file_path'])
        except:
            pass
        cursor.execute('DELETE FROM documents WHERE id = ?', (doc_id,))
        cursor.execute('DELETE FROM chat_history WHERE document_id = ?', (doc_id,))
        cursor.execute('DELETE FROM quizzes WHERE document_id = ?', (doc_id,))
        conn.commit()
        flash('Document deleted!', 'success')
    conn.close()
    return redirect(url_for('dashboard'))


init_db()

if __name__ == '__main__':
    app.run(debug=os.environ.get('FLASK_DEBUG') == '1', host='127.0.0.1',
            port=int(os.environ.get('PORT', '5000')))
