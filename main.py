from fastapi import FastAPI
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
from datetime import datetime
import re
import string

app = FastAPI(
    title="Scriptora API",
    description="Plataforma de analítica textual y escritural: Text + Write.",
    version="0.3.0"
)

class TextAnalysisRequest(BaseModel):
    text: str
    language: str = "es"
    context: str | None = None
    level: str | None = None
    genre: str | None = None
    purpose: str | None = None

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
        "el","la","los","las","un","una","unos","unas","de","del","a","al",
        "en","con","por","para","sin","y","o","pero","que","se","me","te",
        "lo","le","les","es","son","fue","era","ser","estar","está","están",
        "como","más","menos","muy","también","no","sí"
    }
    content_words = [w for w in words if w not in function_words]
    return round(len(content_words) / len(words), 3)

def interpret_text_metrics(word_count, avg_sentence_length, lexical_density):
    comments = []
    if word_count < 50:
        comments.append("El texto es breve; las métricas deben interpretarse con cautela.")
    else:
        comments.append("El texto posee extensión suficiente para una caracterización inicial.")

    if avg_sentence_length < 12:
        comments.append("La longitud oracional promedio es baja, compatible con estructura sintáctica simple.")
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

@app.get("/", response_class=HTMLResponse)
def home():
    return """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Scriptora</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background: #f7f7fb;
            margin: 0;
            padding: 40px;
            color: #222;
        }
        .container {
            max-width: 900px;
            margin: auto;
            background: white;
            padding: 30px;
            border-radius: 18px;
            box-shadow: 0 8px 30px rgba(0,0,0,0.08);
        }
        h1 { margin-bottom: 5px; }
        .subtitle { color: #666; margin-bottom: 25px; }
        textarea, select, button {
            width: 100%;
            margin-top: 10px;
            margin-bottom: 15px;
            padding: 12px;
            font-size: 15px;
            border-radius: 10px;
            border: 1px solid #ccc;
        }
        textarea { height: 180px; }
        button {
            background: #111827;
            color: white;
            cursor: pointer;
            border: none;
            font-weight: bold;
        }
        button:hover { background: #374151; }
        .result {
            margin-top: 25px;
            padding: 20px;
            background: #f0f4ff;
            border-radius: 14px;
            display: none;
        }
        .metric {
            padding: 8px 0;
            border-bottom: 1px solid #ddd;
        }
        .note {
            margin-top: 20px;
            font-size: 13px;
            color: #666;
        }
    </style>
</head>
<body>
<div class="container">
    <h1>Scriptora</h1>
    <div class="subtitle">Plataforma de analítica textual y escritural · v0.3</div>

    <label>Texto a analizar</label>
    <textarea id="textInput" placeholder="Pega aquí un texto para analizar..."></textarea>

    <label>Nivel / audiencia objetivo</label>
    <select id="level">
        <option value="general">General</option>
        <option value="1_4_basico">1°–4° básico</option>
        <option value="5_8_basico">5°–8° básico</option>
        <option value="media">Enseñanza media</option>
        <option value="universitario">Universitario</option>
        <option value="adulto">Adulto general</option>
    </select>

    <label>Tipo de texto</label>
    <select id="genre">
        <option value="general">General</option>
        <option value="narrativo">Narrativo</option>
        <option value="argumentativo">Argumentativo</option>
        <option value="expositivo">Expositivo</option>
        <option value="tecnico">Técnico</option>
    </select>

    <button onclick="analyzeText()">Analizar texto</button>

    <div id="result" class="result">
        <h2>Resultado preliminar</h2>
        <div id="metrics"></div>
        <h3>Interpretación</h3>
        <p id="interpretation"></p>
        <div class="note">
            Análisis preliminar. Próximas versiones integrarán TRUNAJOD/MetaSistema y benchmarks contextuales.
        </div>
    </div>
</div>

<script>
async function analyzeText() {
    const text = document.getElementById("textInput").value;
    const level = document.getElementById("level").value;
    const genre = document.getElementById("genre").value;

    const response = await fetch("/api/text/analyze", {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify({
            text: text,
            language: "es",
            context: "text_analysis",
            level: level,
            genre: genre,
            purpose: "preliminary_analysis"
        })
    });

    const data = await response.json();
    const metrics = data.raw_metrics;

    document.getElementById("metrics").innerHTML = `
        <div class="metric"><strong>Palabras:</strong> ${metrics.word_count}</div>
        <div class="metric"><strong>Caracteres:</strong> ${metrics.char_count}</div>
        <div class="metric"><strong>Oraciones:</strong> ${metrics.sentence_count}</div>
        <div class="metric"><strong>Palabras únicas:</strong> ${metrics.unique_words}</div>
        <div class="metric"><strong>Longitud oracional promedio:</strong> ${metrics.avg_sentence_length}</div>
        <div class="metric"><strong>TTR:</strong> ${metrics.type_token_ratio}</div>
        <div class="metric"><strong>Densidad léxica estimada:</strong> ${metrics.lexical_density_proxy}</div>
    `;

    document.getElementById("interpretation").innerText = data.interpretation;
    document.getElementById("result").style.display = "block";
}
</script>
</body>
</html>
"""

@app.get("/api/health")
def health():
    return {"status": "ok", "timestamp": datetime.utcnow().isoformat()}

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
        avg_sentence_length,
        lexical_density
    )

    return {
        "module": "Scriptora Text",
        "version": "0.3.0",
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
        "interpretation": interpretation
    }

@app.get("/api/benchmarks")
def list_benchmarks():
    return {
        "benchmarks": [
            {"benchmark_id": "ES_TEXT_SCHOOL_1_4_GENERAL", "module": "Text", "language": "es", "status": "planned"},
            {"benchmark_id": "ES_TEXT_SCHOOL_5_8_GENERAL", "module": "Text", "language": "es", "status": "planned"},
            {"benchmark_id": "ES_TEXT_SECONDARY_GENERAL", "module": "Text", "language": "es", "status": "planned"},
            {"benchmark_id": "ES_WRITE_SECONDARY_ARGUMENTATIVE", "module": "Write", "language": "es", "status": "planned"}
        ]
    }
