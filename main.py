from fastapi import FastAPI
from pydantic import BaseModel
from datetime import datetime
import re
import string

app = FastAPI(
    title="Scriptora API",
    description="Plataforma de analítica textual y escritural: Text + Write.",
    version="0.2.0"
)

class TextAnalysisRequest(BaseModel):
    text: str
    language: str = "es"
    context: str | None = None
    level: str | None = None
    genre: str | None = None
    purpose: str | None = None

class WriteSessionRequest(BaseModel):
    user_id: str
    task_id: str
    language: str = "es"
    task_type: str | None = None
    level: str | None = None
    genre: str | None = None

class WriteSessionSubmit(BaseModel):
    session_id: str
    final_text: str
    events: list[dict]

def split_sentences(text: str):
    sentences = re.split(r'[.!?¿¡]+', text)
    return [s.strip() for s in sentences if s.strip()]

def tokenize_words(text: str):
    text = text.lower()
    text = text.translate(str.maketrans('', '', string.punctuation + "¿¡"))
    return [w for w in text.split() if w.strip()]

def estimate_lexical_density(words):
    if not words:
        return 0

    function_words = {
        "el", "la", "los", "las", "un", "una", "unos", "unas",
        "de", "del", "a", "al", "en", "con", "por", "para", "sin",
        "y", "o", "pero", "que", "se", "me", "te", "lo", "le", "les",
        "es", "son", "fue", "era", "ser", "estar", "está", "están",
        "como", "más", "menos", "muy", "también", "no", "sí"
    }

    content_words = [w for w in words if w not in function_words]
    return round(len(content_words) / len(words), 3)

def interpret_text_metrics(word_count, sentence_count, avg_sentence_length, lexical_density):
    comments = []

    if word_count < 50:
        comments.append("El texto es breve, por lo que las métricas deben interpretarse con cautela.")
    else:
        comments.append("El texto posee extensión suficiente para una caracterización inicial.")

    if avg_sentence_length < 12:
        comments.append("La longitud oracional promedio es baja, compatible con un texto de estructura sintáctica simple.")
    elif avg_sentence_length <= 22:
        comments.append("La longitud oracional promedio se ubica en un rango medio.")
    else:
        comments.append("La longitud oracional promedio es alta, lo que puede incrementar la complejidad de procesamiento.")

    if lexical_density < 0.45:
        comments.append("La densidad léxica estimada es baja.")
    elif lexical_density <= 0.65:
        comments.append("La densidad léxica estimada es media.")
    else:
        comments.append("La densidad léxica estimada es alta, compatible con mayor concentración informativa.")

    return " ".join(comments)

@app.get("/")
def home():
    return {
        "app": "Scriptora",
        "version": "0.2.0",
        "status": "running",
        "modules": ["Text", "Write", "Benchmarks", "Reports"]
    }

@app.get("/api/health")
def health():
    return {
        "status": "ok",
        "timestamp": datetime.utcnow().isoformat()
    }

@app.post("/api/text/analyze")
def analyze_text(request: TextAnalysisRequest):
    text = request.text.strip()
    words = tokenize_words(text)
    sentences = split_sentences(text)

    word_count = len(words)
    char_count = len(text)
    sentence_count = len(sentences)
    unique_words = len(set(words))

    avg_sentence_length = round(word_count / sentence_count, 2) if sentence_count > 0 else 0
    ttr = round(unique_words / word_count, 3) if word_count > 0 else 0
    lexical_density = estimate_lexical_density(words)

    interpretation = interpret_text_metrics(
        word_count,
        sentence_count,
        avg_sentence_length,
        lexical_density
    )

    return {
        "module": "Scriptora Text",
        "version": "0.2.0",
        "language": request.language,
        "context": request.context,
        "level": request.level,
        "genre": request.genre,
        "purpose": request.purpose,
        "raw_metrics": {
            "word_count": word_count,
            "char_count": char_count,
            "sentence_count": sentence_count,
            "unique_words": unique_words,
            "avg_sentence_length": avg_sentence_length,
            "type_token_ratio": ttr,
            "lexical_density_proxy": lexical_density
        },
        "interpretation": interpretation,
        "note": "Análisis textual preliminar. En próximas versiones se integrarán índices TRUNAJOD/MetaSistema y benchmarks contextuales."
    }

@app.post("/api/write/session/start")
def start_write_session(request: WriteSessionRequest):
    session_id = f"SW-{request.user_id}-{int(datetime.utcnow().timestamp())}"

    return {
        "module": "Scriptora Write",
        "session_id": session_id,
        "user_id": request.user_id,
        "task_id": request.task_id,
        "language": request.language,
        "status": "session_started"
    }

@app.post("/api/write/session/submit")
def submit_write_session(request: WriteSessionSubmit):
    final_text = request.final_text.strip()
    event_count = len(request.events)

    return {
        "module": "Scriptora Write",
        "session_id": request.session_id,
        "status": "session_received",
        "final_text_metrics": {
            "word_count": len(tokenize_words(final_text)),
            "char_count": len(final_text)
        },
        "process_metrics": {
            "event_count": event_count
        },
        "message": "Sesión recibida. En próximas versiones se calcularán pausas, bursts, revisiones y timeline."
    }

@app.get("/api/benchmarks")
def list_benchmarks():
    return {
        "benchmarks": [
            {
                "benchmark_id": "ES_TEXT_SCHOOL_1_4_GENERAL",
                "module": "Text",
                "language": "es",
                "status": "planned"
            },
            {
                "benchmark_id": "ES_TEXT_SCHOOL_5_8_GENERAL",
                "module": "Text",
                "language": "es",
                "status": "planned"
            },
            {
                "benchmark_id": "ES_TEXT_SECONDARY_GENERAL",
                "module": "Text",
                "language": "es",
                "status": "planned"
            },
            {
                "benchmark_id": "ES_WRITE_SECONDARY_ARGUMENTATIVE",
                "module": "Write",
                "language": "es",
                "status": "planned"
            }
        ]
    }
