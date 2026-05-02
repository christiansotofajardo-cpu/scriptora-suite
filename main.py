from fastapi import FastAPI
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
from datetime import datetime
import re
import string

app = FastAPI(
    title="Scriptora Suite API",
    description="Plataforma integrada de analítica textual y evaluación escritural: Scriptora T + Scriptora W.",
    version="0.4.0"
)

# =========================
# MODELOS DE ENTRADA
# =========================

class TextAnalysisRequest(BaseModel):
    text: str
    language: str = "es"
    context: str | None = None
    level: str | None = None
    genre: str | None = None
    purpose: str | None = None


class WritingEvaluationRequest(BaseModel):
    text: str
    language: str = "es"
    level: str | None = None
    genre: str | None = None
    task: str | None = None
    purpose: str | None = None


# =========================
# FUNCIONES BÁSICAS
# =========================

def split_sentences(text: str):
    sentences = re.split(r'[.!?¿¡]+', text)
    return [s.strip() for s in sentences if s.strip()]


def tokenize_words(text: str):
    text = text.lower()
    text = text.translate(str.maketrans('', '', string.punctuation + "¿¡“”‘’"))
    return [w for w in text.split() if w.strip()]


def estimate_lexical_density(words):
    if not words:
        return 0

    function_words = {
        "el","la","los","las","un","una","unos","unas","de","del","a","al",
        "en","con","por","para","sin","y","o","pero","que","se","me","te",
        "lo","le","les","es","son","fue","era","ser","estar","está","están",
        "como","más","menos","muy","también","no","sí","su","sus","mi","mis",
        "tu","tus","este","esta","estos","estas","ese","esa","eso","hay"
    }

    content_words = [w for w in words if w not in function_words]
    return round(len(content_words) / len(words), 3)


def count_connectors(text: str):
    connectors = [
        "porque", "por lo tanto", "sin embargo", "además", "también",
        "aunque", "en cambio", "por ejemplo", "finalmente", "primero",
        "segundo", "luego", "después", "entonces", "así", "ya que",
        "debido a", "en conclusión", "por otra parte"
    ]

    text_lower = text.lower()
    count = 0
    found = []

    for connector in connectors:
        occurrences = text_lower.count(connector)
        if occurrences > 0:
            count += occurrences
            found.append(connector)

    return count, sorted(list(set(found)))


def estimate_paragraphs(text: str):
    paragraphs = [p.strip() for p in text.split("\n") if p.strip()]
    return len(paragraphs)


# =========================
# INTERPRETACIÓN SCRIPTORA T
# =========================

def interpret_text_metrics(word_count, avg_sentence_length, lexical_density, ttr):
    comments = []

    if word_count < 50:
        comments.append("El texto es breve; las métricas deben interpretarse con cautela.")
    elif word_count <= 250:
        comments.append("El texto posee una extensión suficiente para una caracterización inicial.")
    else:
        comments.append("El texto posee una extensión amplia, adecuada para un análisis más estable.")

    if avg_sentence_length < 12:
        comments.append("La longitud oracional promedio es baja, compatible con una estructura sintáctica simple.")
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

    if ttr < 0.35:
        comments.append("La diversidad léxica superficial es baja.")
    elif ttr <= 0.60:
        comments.append("La diversidad léxica superficial se encuentra en un rango medio.")
    else:
        comments.append("La diversidad léxica superficial es alta, aunque en textos breves este valor puede sobreestimarse.")

    return " ".join(comments)


# =========================
# EVALUACIÓN SCRIPTORA W
# =========================

def score_writing(text, words, sentences, genre):
    word_count = len(words)
    sentence_count = len(sentences)
    paragraph_count = estimate_paragraphs(text)
    connector_count, connectors_found = count_connectors(text)

    avg_sentence_length = round(word_count / sentence_count, 2) if sentence_count > 0 else 0
    unique_words = len(set(words))
    ttr = round(unique_words / word_count, 3) if word_count > 0 else 0

    # Puntuaciones preliminares 0-100
    extension_score = min(100, round((word_count / 250) * 100)) if word_count > 0 else 0

    if paragraph_count >= 3:
        organization_score = 85
    elif paragraph_count == 2:
        organization_score = 70
    elif paragraph_count == 1 and word_count > 80:
        organization_score = 55
    else:
        organization_score = 40

    if connector_count >= 6:
        cohesion_score = 90
    elif connector_count >= 3:
        cohesion_score = 75
    elif connector_count >= 1:
        cohesion_score = 60
    else:
        cohesion_score = 40

    if 10 <= avg_sentence_length <= 24:
        syntax_score = 80
    elif avg_sentence_length < 10:
        syntax_score = 60
    else:
        syntax_score = 65

    if ttr >= 0.55:
        lexical_score = 85
    elif ttr >= 0.40:
        lexical_score = 70
    else:
        lexical_score = 55

    genre_score = 65
    genre_comment = "La adecuación al género se estima de manera preliminar."

    lower_text = text.lower()

    if genre == "argumentativo":
        argument_markers = ["creo", "pienso", "opino", "debería", "por lo tanto", "en conclusión", "argumento", "razón"]
        found_argument_markers = sum(1 for m in argument_markers if m in lower_text)
        genre_score = min(95, 55 + found_argument_markers * 8)
        genre_comment = "En el texto argumentativo se observaron marcas preliminares de postura, justificación o cierre."

    elif genre == "narrativo":
        narrative_markers = ["un día", "luego", "después", "entonces", "finalmente", "personaje", "historia"]
        found_narrative_markers = sum(1 for m in narrative_markers if m in lower_text)
        genre_score = min(95, 55 + found_narrative_markers * 8)
        genre_comment = "En el texto narrativo se observaron marcas preliminares de secuencia o progresión temporal."

    elif genre == "expositivo":
        expository_markers = ["es decir", "por ejemplo", "se define", "consiste", "presenta", "explica"]
        found_expository_markers = sum(1 for m in expository_markers if m in lower_text)
        genre_score = min(95, 55 + found_expository_markers * 8)
        genre_comment = "En el texto expositivo se observaron marcas preliminares de explicación o desarrollo informativo."

    global_score = round(
        extension_score * 0.15 +
        organization_score * 0.20 +
        cohesion_score * 0.20 +
        syntax_score * 0.15 +
        lexical_score * 0.15 +
        genre_score * 0.15,
        1
    )

    if global_score < 50:
        level_label = "Inicial"
    elif global_score < 65:
        level_label = "En desarrollo"
    elif global_score < 80:
        level_label = "Adecuado"
    else:
        level_label = "Avanzado"

    interpretation = (
        f"El desempeño escritural preliminar se ubica en nivel {level_label}. "
        f"La evaluación considera extensión, organización, cohesión, estructura sintáctica, diversidad léxica y adecuación al género. "
        f"{genre_comment} "
        f"Esta estimación es inicial y deberá calibrarse posteriormente con rúbricas humanas, benchmarks y datos TRUNAJOD."
    )

    return {
        "word_count": word_count,
        "sentence_count": sentence_count,
        "paragraph_count": paragraph_count,
        "avg_sentence_length": avg_sentence_length,
        "unique_words": unique_words,
        "type_token_ratio": ttr,
        "connector_count": connector_count,
        "connectors_found": connectors_found,
        "scores": {
            "extension": extension_score,
            "organization": organization_score,
            "cohesion": cohesion_score,
            "syntax_control": syntax_score,
            "lexical_variety": lexical_score,
            "genre_adequacy": genre_score,
            "global_writing_score": global_score
        },
        "level_label": level_label,
        "interpretation": interpretation
    }


# =========================
# INTERFAZ WEB
# =========================

@app.get("/", response_class=HTMLResponse)
def home():
    return """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Scriptora Suite</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background: #f7f7fb;
            margin: 0;
            padding: 40px;
            color: #222;
        }
        .container {
            max-width: 950px;
            margin: auto;
            background: white;
            padding: 30px;
            border-radius: 18px;
            box-shadow: 0 8px 30px rgba(0,0,0,0.08);
        }
        h1 { margin-bottom: 5px; font-size: 34px; }
        .subtitle { color: #666; margin-bottom: 25px; }
        textarea, select, button {
            width: 100%;
            margin-top: 10px;
            margin-bottom: 15px;
            padding: 12px;
            font-size: 15px;
            border-radius: 10px;
            border: 1px solid #ccc;
            box-sizing: border-box;
        }
        textarea { height: 190px; }
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
        .score-box {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 12px;
            margin-top: 10px;
        }
        .score {
            background: white;
            padding: 14px;
            border-radius: 12px;
            border: 1px solid #d8def5;
        }
        .score strong {
            display: block;
            font-size: 14px;
            color: #374151;
        }
        .score span {
            font-size: 22px;
            font-weight: bold;
        }
        .note {
            margin-top: 20px;
            font-size: 13px;
            color: #666;
        }
        .pill {
            display: inline-block;
            padding: 6px 10px;
            border-radius: 999px;
            background: #111827;
            color: white;
            font-size: 13px;
            margin-bottom: 10px;
        }
    </style>
</head>
<body>
<div class="container">
    <h1>Scriptora Suite</h1>
    <div class="subtitle">Plataforma de analítica textual y evaluación escritural · v0.4</div>

    <label>Módulo de análisis</label>
    <select id="module">
        <option value="text">Scriptora T · Analizar texto</option>
        <option value="write">Scriptora W · Evaluar escritura</option>
    </select>

    <label>Texto a analizar</label>
    <textarea id="textInput" placeholder="Pega aquí un texto para analizar o evaluar..."></textarea>

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

    <button onclick="runScriptora()">Analizar con Scriptora</button>

    <div id="result" class="result">
        <div id="moduleLabel" class="pill"></div>
        <h2>Resultado preliminar</h2>
        <div id="metrics"></div>
        <div id="scores"></div>
        <h3>Interpretación</h3>
        <p id="interpretation"></p>
        <div class="note">
            Análisis preliminar. Próximas versiones integrarán TRUNAJOD, MetaSistema, rúbricas humanas, benchmarks contextuales y reportes exportables.
        </div>
    </div>
</div>

<script>
async function runScriptora() {
    const selectedModule = document.getElementById("module").value;
    const text = document.getElementById("textInput").value;
    const level = document.getElementById("level").value;
    const genre = document.getElementById("genre").value;

    if (!text.trim()) {
        alert("Por favor pega un texto antes de analizar.");
        return;
    }

    let endpoint = "/api/text/analyze";
    let payload = {
        text: text,
        language: "es",
        context: "scriptora_suite",
        level: level,
        genre: genre,
        purpose: "preliminary_analysis"
    };

    if (selectedModule === "write") {
        endpoint = "/api/write/evaluate";
        payload = {
            text: text,
            language: "es",
            level: level,
            genre: genre,
            task: "open_writing_task",
            purpose: "preliminary_writing_evaluation"
        };
    }

    const response = await fetch(endpoint, {
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: JSON.stringify(payload)
    });

    const data = await response.json();

    document.getElementById("result").style.display = "block";
    document.getElementById("moduleLabel").innerText = data.module;

    if (selectedModule === "text") {
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

        document.getElementById("scores").innerHTML = "";
        document.getElementById("interpretation").innerText = data.interpretation;
    }

    if (selectedModule === "write") {
        const metrics = data.writing_metrics;
        const scores = metrics.scores;

        document.getElementById("metrics").innerHTML = `
            <div class="metric"><strong>Palabras:</strong> ${metrics.word_count}</div>
            <div class="metric"><strong>Oraciones:</strong> ${metrics.sentence_count}</div>
            <div class="metric"><strong>Párrafos:</strong> ${metrics.paragraph_count}</div>
            <div class="metric"><strong>Longitud oracional promedio:</strong> ${metrics.avg_sentence_length}</div>
            <div class="metric"><strong>TTR:</strong> ${metrics.type_token_ratio}</div>
            <div class="metric"><strong>Conectores detectados:</strong> ${metrics.connector_count}</div>
            <div class="metric"><strong>Conectores:</strong> ${metrics.connectors_found.join(", ") || "No detectados"}</div>
        `;

        document.getElementById("scores").innerHTML = `
            <h3>Puntajes preliminares</h3>
            <div class="score-box">
                <div class="score"><strong>Extensión</strong><span>${scores.extension}</span></div>
                <div class="score"><strong>Organización</strong><span>${scores.organization}</span></div>
                <div class="score"><strong>Cohesión</strong><span>${scores.cohesion}</span></div>
                <div class="score"><strong>Sintaxis</strong><span>${scores.syntax_control}</span></div>
                <div class="score"><strong>Léxico</strong><span>${scores.lexical_variety}</span></div>
                <div class="score"><strong>Adecuación al género</strong><span>${scores.genre_adequacy}</span></div>
                <div class="score"><strong>Puntaje global</strong><span>${scores.global_writing_score}</span></div>
                <div class="score"><strong>Nivel</strong><span>${metrics.level_label}</span></div>
            </div>
        `;

        document.getElementById("interpretation").innerText = metrics.interpretation;
    }
}
</script>
</body>
</html>
"""


# =========================
# ENDPOINTS
# =========================

@app.get("/api/health")
def health():
    return {
        "status": "ok",
        "service": "Scriptora Suite",
        "version": "0.4.0",
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
        avg_sentence_length,
        lexical_density,
        ttr
    )

    return {
        "module": "Scriptora T · Text Analysis",
        "version": "0.4.0",
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


@app.post("/api/write/evaluate")
def evaluate_writing(request: WritingEvaluationRequest):
    text = request.text.strip()
    words = tokenize_words(text)
    sentences = split_sentences(text)

    writing_metrics = score_writing(
        text=text,
        words=words,
        sentences=sentences,
        genre=request.genre or "general"
    )

    return {
        "module": "Scriptora W · Writing Evaluation",
        "version": "0.4.0",
        "language": request.language,
        "level": request.level,
        "genre": request.genre,
        "task": request.task,
        "purpose": request.purpose,
        "writing_metrics": writing_metrics
    }


@app.get("/api/benchmarks")
def list_benchmarks():
    return {
        "benchmarks": [
            {
                "benchmark_id": "ES_TEXT_SCHOOL_1_4_GENERAL",
                "module": "Scriptora T",
                "language": "es",
                "status": "planned"
            },
            {
                "benchmark_id": "ES_TEXT_SCHOOL_5_8_GENERAL",
                "module": "Scriptora T",
                "language": "es",
                "status": "planned"
            },
            {
                "benchmark_id": "ES_TEXT_SECONDARY_GENERAL",
                "module": "Scriptora T",
                "language": "es",
                "status": "planned"
            },
            {
                "benchmark_id": "ES_WRITE_SECONDARY_ARGUMENTATIVE",
                "module": "Scriptora W",
                "language": "es",
                "status": "planned"
            },
            {
                "benchmark_id": "ES_WRITE_SECONDARY_NARRATIVE",
                "module": "Scriptora W",
                "language": "es",
                "status": "planned"
            }
        ]
    }
