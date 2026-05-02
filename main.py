from fastapi import FastAPI
from pydantic import BaseModel
from datetime import datetime

app = FastAPI(
    title="Scriptora API",
    description="Plataforma de analítica textual y escritural: Text + Write.",
    version="0.1.0"
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

@app.get("/")
def home():
    return {
        "app": "Scriptora",
        "version": "0.1.0",
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
    word_count = len(text.split())
    char_count = len(text)

    return {
        "module": "Scriptora Text",
        "language": request.language,
        "context": request.context,
        "level": request.level,
        "genre": request.genre,
        "purpose": request.purpose,
        "raw_metrics": {
            "word_count": word_count,
            "char_count": char_count
        },
        "message": "Análisis textual básico ejecutado. En próximas versiones se integrarán índices TRUNAJOD/MetaSistema."
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
            "word_count": len(final_text.split()),
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
                "benchmark_id": "ES_WRITE_SECONDARY_ARGUMENTATIVE",
                "module": "Write",
                "language": "es",
                "status": "planned"
            }
        ]
    }
