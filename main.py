from fastapi import FastAPI
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
from datetime import datetime
import re
import string

app = FastAPI(
    title="Scriptora Suite API",
    description="Plataforma integrada de analítica textual, evaluación escritural y regulación del proceso de escritura.",
    version="0.6.0"
)

# ============================================================
# MODELOS DE ENTRADA
# ============================================================

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


class WritingProcessRequest(BaseModel):
    text: str
    language: str = "es"
    level: str | None = None
    genre: str | None = None
    task: str | None = None
    purpose: str | None = None

    writing_mode: str | None = "live"
    total_time_seconds: float | None = 0
    initial_latency_seconds: float | None = 0
    pause_count: int | None = 0
    long_pause_count: int | None = 0
    edit_count: int | None = 0
    deletion_count: int | None = 0
    insertion_count: int | None = 0
    local_adjustment_count: int | None = 0
    reformulation_count: int | None = 0
    expansion_count: int | None = 0
    reduction_count: int | None = 0
    macro_adjustment_count: int | None = 0
    max_text_length: int | None = 0
    final_text_length: int | None = 0
    input_event_count: int | None = 0


# ============================================================
# FUNCIONES BÁSICAS
# ============================================================

def split_sentences(text: str):
    sentences = re.split(r'[.!?¿¡]+', text)
    return [s.strip() for s in sentences if s.strip()]


def tokenize_words(text: str):
    text = text.lower()
    text = text.translate(str.maketrans('', '', string.punctuation + "¿¡“”‘’"))
    return [w for w in text.split() if w.strip()]


def estimate_paragraphs(text: str):
    paragraphs = [p.strip() for p in text.split("\n") if p.strip()]
    return len(paragraphs)


def estimate_lexical_density(words):
    if not words:
        return 0

    function_words = {
        "el","la","los","las","un","una","unos","unas","de","del","a","al",
        "en","con","por","para","sin","y","o","pero","que","se","me","te",
        "lo","le","les","es","son","fue","era","ser","estar","está","están",
        "como","más","menos","muy","también","no","sí","su","sus","mi","mis",
        "tu","tus","este","esta","estos","estas","ese","esa","eso","hay",
        "ha","han","he","hemos","fui","fueron","tiene","tienen","tengo",
        "cuando","donde","quien","quienes","cual","cuales"
    }

    content_words = [w for w in words if w not in function_words]
    return round(len(content_words) / len(words), 3)


def count_connectors(text: str):
    connectors = [
        "porque", "por lo tanto", "sin embargo", "además", "también",
        "aunque", "en cambio", "por ejemplo", "finalmente", "primero",
        "segundo", "luego", "después", "entonces", "así", "ya que",
        "debido a", "en conclusión", "por otra parte", "por consiguiente",
        "a pesar de", "en primer lugar", "en segundo lugar", "por ende",
        "de este modo", "en síntesis"
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


def count_punctuation_marks(text: str):
    marks = [".", ",", ";", ":", "¿", "?", "¡", "!"]
    return sum(text.count(m) for m in marks)


def detect_closure(text: str):
    closure_markers = [
        "en conclusión", "finalmente", "en síntesis",
        "por último", "para concluir", "en resumen"
    ]
    lower_text = text.lower()
    return any(marker in lower_text for marker in closure_markers)


# ============================================================
# SCRIPTORA T
# ============================================================

def interpret_text_metrics(word_count, avg_sentence_length, lexical_density, ttr):
    comments = []

    if word_count < 50:
        comments.append("El texto es breve; las métricas deben interpretarse con cautela.")
    elif word_count <= 250:
        comments.append("El texto posee una extensión suficiente para una caracterización inicial.")
    else:
        comments.append("El texto posee una extensión amplia, adecuada para un análisis más estable.")

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

    if ttr < 0.35:
        comments.append("La diversidad léxica superficial es baja.")
    elif ttr <= 0.60:
        comments.append("La diversidad léxica superficial se encuentra en un rango medio.")
    else:
        comments.append("La diversidad léxica superficial es alta, aunque en textos breves este valor puede sobreestimarse.")

    return " ".join(comments)


# ============================================================
# SCRIPTORA W: PRODUCTO ESCRITO
# ============================================================

def score_writing_product(text, words, sentences, genre):
    word_count = len(words)
    sentence_count = len(sentences)
    paragraph_count = estimate_paragraphs(text)
    connector_count, connectors_found = count_connectors(text)
    punctuation_count = count_punctuation_marks(text)
    lexical_density = estimate_lexical_density(words)

    avg_sentence_length = round(word_count / sentence_count, 2) if sentence_count > 0 else 0
    unique_words = len(set(words))
    ttr = round(unique_words / word_count, 3) if word_count > 0 else 0

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
        lexical_variety_score = 85
    elif ttr >= 0.40:
        lexical_variety_score = 70
    else:
        lexical_variety_score = 55

    if lexical_density >= 0.65:
        informational_density_score = 85
    elif lexical_density >= 0.45:
        informational_density_score = 70
    else:
        informational_density_score = 55

    if sentence_count >= 3 and word_count >= 80:
        elaboration_score = 80
    elif sentence_count >= 2 and word_count >= 40:
        elaboration_score = 65
    elif word_count >= 20:
        elaboration_score = 50
    else:
        elaboration_score = 35

    if connector_count >= 2 and paragraph_count >= 2:
        global_coherence_score = 80
    elif connector_count >= 1 or paragraph_count >= 2:
        global_coherence_score = 65
    else:
        global_coherence_score = 45

    if punctuation_count >= 6:
        punctuation_control_score = 80
    elif punctuation_count >= 3:
        punctuation_control_score = 65
    elif punctuation_count >= 1:
        punctuation_control_score = 50
    else:
        punctuation_control_score = 35

    closure_present = detect_closure(text)
    closure_score = 80 if closure_present else 45

    lower_text = text.lower()
    genre_score = 65
    genre_comment = "La adecuación al género se estima de manera preliminar."

    if genre == "argumentativo":
        argument_markers = [
            "creo", "pienso", "opino", "debería", "por lo tanto",
            "en conclusión", "argumento", "razón", "postura",
            "estoy de acuerdo", "no estoy de acuerdo", "a favor", "en contra"
        ]
        found_argument_markers = sum(1 for m in argument_markers if m in lower_text)
        genre_score = min(95, 55 + found_argument_markers * 8)
        genre_comment = "En el texto argumentativo se observaron marcas preliminares de postura, justificación o cierre."

    elif genre == "narrativo":
        narrative_markers = [
            "un día", "luego", "después", "entonces", "finalmente",
            "personaje", "historia", "había", "ocurrió", "mientras"
        ]
        found_narrative_markers = sum(1 for m in narrative_markers if m in lower_text)
        genre_score = min(95, 55 + found_narrative_markers * 8)
        genre_comment = "En el texto narrativo se observaron marcas preliminares de secuencia o progresión temporal."

    elif genre == "expositivo":
        expository_markers = [
            "es decir", "por ejemplo", "se define", "consiste",
            "presenta", "explica", "corresponde", "se caracteriza"
        ]
        found_expository_markers = sum(1 for m in expository_markers if m in lower_text)
        genre_score = min(95, 55 + found_expository_markers * 8)
        genre_comment = "En el texto expositivo se observaron marcas preliminares de explicación o desarrollo informativo."

    global_score = round(
        extension_score * 0.10 +
        organization_score * 0.12 +
        cohesion_score * 0.12 +
        syntax_score * 0.10 +
        lexical_variety_score * 0.10 +
        informational_density_score * 0.10 +
        elaboration_score * 0.14 +
        global_coherence_score * 0.12 +
        genre_score * 0.07 +
        closure_score * 0.03,
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
        f"La evaluación considera extensión, organización, cohesión, sintaxis, diversidad léxica, densidad informativa, elaboración de ideas, coherencia global, adecuación al género y cierre textual. "
        f"{genre_comment} "
        f"Esta estimación es exploratoria y deberá calibrarse con rúbricas humanas, benchmarks y datos TRUNAJOD."
    )

    return {
        "word_count": word_count,
        "sentence_count": sentence_count,
        "paragraph_count": paragraph_count,
        "avg_sentence_length": avg_sentence_length,
        "unique_words": unique_words,
        "type_token_ratio": ttr,
        "lexical_density_proxy": lexical_density,
        "connector_count": connector_count,
        "connectors_found": connectors_found,
        "punctuation_count": punctuation_count,
        "closure_present": closure_present,
        "scores": {
            "extension": extension_score,
            "organization": organization_score,
            "cohesion": cohesion_score,
            "syntax_control": syntax_score,
            "lexical_variety": lexical_variety_score,
            "informational_density": informational_density_score,
            "idea_elaboration": elaboration_score,
            "global_coherence": global_coherence_score,
            "punctuation_control": punctuation_control_score,
            "genre_adequacy": genre_score,
            "textual_closure": closure_score,
            "global_writing_score": global_score
        },
        "level_label": level_label,
        "interpretation": interpretation
    }


# ============================================================
# SCRIPTORA W: PROCESO / REGULACIÓN
# ============================================================

def score_writing_process(request: WritingProcessRequest):
    text = request.text.strip()

    writing_mode = request.writing_mode or "live"
    total_time = request.total_time_seconds or 0
    initial_latency = request.initial_latency_seconds or 0
    long_pause_count = request.long_pause_count or 0
    edit_count = request.edit_count or 0
    deletion_count = request.deletion_count or 0
    insertion_count = request.insertion_count or 0
    local_adjustment_count = request.local_adjustment_count or 0
    reformulation_count = request.reformulation_count or 0
    expansion_count = request.expansion_count or 0
    reduction_count = request.reduction_count or 0
    macro_adjustment_count = request.macro_adjustment_count or 0
    max_text_length = request.max_text_length or len(text)
    final_text_length = request.final_text_length or len(text)
    input_event_count = request.input_event_count or 0

    words = tokenize_words(text)
    word_count = len(words)

    words_per_minute = round((word_count / total_time) * 60, 2) if total_time > 0 else 0
    final_stability_ratio = round(final_text_length / max_text_length, 3) if max_text_length > 0 else 1

    if writing_mode == "pasted":
        return {
            "writing_mode": writing_mode,
            "total_time_seconds": total_time,
            "initial_latency_seconds": initial_latency,
            "long_pause_count": long_pause_count,
            "edit_count": edit_count,
            "deletion_count": deletion_count,
            "insertion_count": insertion_count,
            "local_adjustment_count": local_adjustment_count,
            "reformulation_count": reformulation_count,
            "expansion_count": expansion_count,
            "reduction_count": reduction_count,
            "macro_adjustment_count": macro_adjustment_count,
            "max_text_length": max_text_length,
            "final_text_length": final_text_length,
            "input_event_count": input_event_count,
            "words_per_minute": words_per_minute,
            "final_stability_ratio": final_stability_ratio,
            "planning_score": 0,
            "monitoring_score": 0,
            "revision_score": 0,
            "reformulation_score": 0,
            "fluency_score": 0,
            "process_regulation_score": 0,
            "process_regulation_label": "No trazable",
            "interpretation": (
                "El texto parece haber sido pegado o ingresado sin trazabilidad suficiente del proceso. "
                "La evaluación del producto textual puede realizarse, pero la regulación escritural no debe inferirse desde estos datos."
            )
        }

    planning_score = 80 if initial_latency >= 5 else 60 if initial_latency >= 2 else 40

    monitoring_score = 85 if long_pause_count >= 3 else 65 if long_pause_count >= 1 else 35

    revision_events = deletion_count + local_adjustment_count + reduction_count
    revision_score = 85 if revision_events >= 12 else 70 if revision_events >= 5 else 55 if revision_events >= 1 else 35

    reformulation_score = 85 if reformulation_count >= 3 else 70 if reformulation_count >= 1 else 35

    if 8 <= words_per_minute <= 40:
        fluency_score = 80
    elif words_per_minute > 40:
        fluency_score = 60
    elif words_per_minute > 0:
        fluency_score = 55
    else:
        fluency_score = 35

    recursivity_score = 85 if macro_adjustment_count >= 2 else 70 if macro_adjustment_count == 1 else 45

    process_score = round(
        planning_score * 0.15 +
        monitoring_score * 0.20 +
        revision_score * 0.20 +
        reformulation_score * 0.15 +
        fluency_score * 0.15 +
        recursivity_score * 0.15,
        1
    )

    if process_score < 45:
        process_label = "Regulación baja"
    elif process_score < 65:
        process_label = "Regulación emergente"
    elif process_score < 80:
        process_label = "Regulación media"
    else:
        process_label = "Regulación alta"

    parts = []

    if initial_latency < 2:
        parts.append("La latencia inicial fue baja, lo que sugiere inicio rápido de la escritura.")
    else:
        parts.append("La latencia inicial sugiere una fase inicial de planificación.")

    if long_pause_count == 0:
        parts.append("Se observaron pocas pausas largas, con baja evidencia de monitoreo reflexivo.")
    else:
        parts.append("Las pausas largas sugieren momentos de monitoreo, planificación o revisión.")

    if local_adjustment_count > 0:
        parts.append("Se detectaron ajustes locales compatibles con control microtextual.")

    if reformulation_count > 0:
        parts.append("Se detectaron reformulaciones, lo que sugiere revisión de segmentos mayores.")

    if macro_adjustment_count > 0:
        parts.append("Se observaron ajustes macrotextuales compatibles con reorganización del texto.")

    interpretation = (
        f"La regulación escritural preliminar se ubica en: {process_label}. "
        + " ".join(parts)
        + " Esta estimación es exploratoria y deberá calibrarse con datos reales de escritura."
    )

    return {
        "writing_mode": writing_mode,
        "total_time_seconds": total_time,
        "initial_latency_seconds": initial_latency,
        "long_pause_count": long_pause_count,
        "edit_count": edit_count,
        "deletion_count": deletion_count,
        "insertion_count": insertion_count,
        "local_adjustment_count": local_adjustment_count,
        "reformulation_count": reformulation_count,
        "expansion_count": expansion_count,
        "reduction_count": reduction_count,
        "macro_adjustment_count": macro_adjustment_count,
        "max_text_length": max_text_length,
        "final_text_length": final_text_length,
        "input_event_count": input_event_count,
        "words_per_minute": words_per_minute,
        "final_stability_ratio": final_stability_ratio,
        "planning_score": planning_score,
        "monitoring_score": monitoring_score,
        "revision_score": revision_score,
        "reformulation_score": reformulation_score,
        "fluency_score": fluency_score,
        "recursivity_score": recursivity_score,
        "process_regulation_score": process_score,
        "process_regulation_label": process_label,
        "interpretation": interpretation
    }


def integrated_writing_interpretation(product_metrics, process_metrics):
    product_label = product_metrics.get("level_label", "No determinado")
    process_label = process_metrics.get("process_regulation_label", "No determinado")
    product_score = product_metrics.get("scores", {}).get("global_writing_score", 0)
    process_score = process_metrics.get("process_regulation_score", 0)

    if process_label == "No trazable":
        return (
            f"El producto escrito se ubica preliminarmente en nivel {product_label}. "
            "Sin embargo, no existe trazabilidad suficiente para interpretar el proceso de escritura. "
            "La síntesis integrada queda limitada al análisis del producto textual."
        )

    if product_score < 50 and process_score < 50:
        return (
            "La síntesis integrada sugiere bajo desarrollo del producto y baja evidencia de regulación escritural. "
            "Esto puede corresponder a una producción breve, poco elaborada y con escasa revisión."
        )

    if product_score < 50 and process_score >= 65:
        return (
            "La síntesis integrada muestra una tensión relevante: el producto escrito aún es bajo, "
            "pero el proceso evidencia señales de planificación, monitoreo o revisión. "
            "Esto puede indicar esfuerzo regulatorio que todavía no logra consolidarse en el texto final."
        )

    if product_score >= 65 and process_score < 50:
        return (
            "La síntesis integrada sugiere un producto relativamente adecuado, pero con pocas señales de regulación observable. "
            "Esto podría corresponder a escritura fluida, automatizada o a una captura insuficiente del proceso."
        )

    return (
        f"La síntesis integrada muestra un producto en nivel {product_label} junto con un proceso clasificado como {process_label}. "
        "El resultado sugiere una relación positiva entre elaboración textual y control regulatorio durante la escritura."
    )


# ============================================================
# INTERFAZ WEB
# ============================================================

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
            max-width: 980px;
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
        .secondary {
            background: #6b7280;
        }
        .secondary:hover {
            background: #4b5563;
        }
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
        .process-panel {
            background: #eef2ff;
            border: 1px solid #c7d2fe;
            padding: 15px;
            border-radius: 12px;
            margin-bottom: 15px;
            font-size: 14px;
        }
        .small-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
            gap: 8px;
            margin-top: 8px;
        }
        .small-box {
            background: white;
            border-radius: 10px;
            padding: 8px;
            border: 1px solid #d8def5;
        }
    </style>
</head>
<body>
<div class="container">
    <h1>Scriptora Suite</h1>
    <div class="subtitle">Analítica textual, evaluación escritural y regulación del proceso · v0.6</div>

    <label>Módulo de análisis</label>
    <select id="module">
        <option value="text">Scriptora T · Analizar texto</option>
        <option value="write_product">Scriptora W · Evaluar producto escrito</option>
        <option value="write_process">Scriptora W · Producto + proceso escritural</option>
    </select>

    <label>Modo de ingreso</label>
    <select id="writingMode">
        <option value="live">Escribir en vivo / capturar proceso</option>
        <option value="pasted">Pegar texto ya escrito / solo producto</option>
    </select>

    <div class="process-panel">
        <strong>Registro de proceso escritural</strong>
        <div class="small-grid">
            <div class="small-box">Tiempo: <span id="timer">0</span> s</div>
            <div class="small-box">Latencia inicial: <span id="latency">0</span> s</div>
            <div class="small-box">Pausas largas: <span id="longPauses">0</span></div>
            <div class="small-box">Ediciones: <span id="edits">0</span></div>
            <div class="small-box">Inserciones: <span id="insertions">0</span></div>
            <div class="small-box">Borrados: <span id="deletions">0</span></div>
            <div class="small-box">Ajustes locales: <span id="localAdjustments">0</span></div>
            <div class="small-box">Reformulaciones: <span id="reformulations">0</span></div>
            <div class="small-box">Expansiones: <span id="expansions">0</span></div>
            <div class="small-box">Reducciones: <span id="reductions">0</span></div>
            <div class="small-box">Ajustes macro: <span id="macroAdjustments">0</span></div>
            <div class="small-box">Eventos: <span id="events">0</span></div>
        </div>
    </div>

    <label>Texto a analizar</label>
    <textarea id="textInput" placeholder="Escribe aquí o pega un texto para analizar..."></textarea>

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
    <button class="secondary" onclick="resetProcessCapture()">Reiniciar captura</button>

    <div id="result" class="result">
        <div id="moduleLabel" class="pill"></div>
        <h2>Resultado preliminar</h2>

        <h3 id="productTitle">Producto textual</h3>
        <div id="metrics"></div>
        <div id="scores"></div>

        <div id="processSection" style="display:none;">
            <h3>Proceso escritural</h3>
            <div id="processMetrics"></div>
            <div id="processScores"></div>
        </div>

        <h3>Interpretación</h3>
        <p id="interpretation"></p>

        <div id="integratedSection" style="display:none;">
            <h3>Síntesis integrada</h3>
            <p id="integratedInterpretation"></p>
        </div>

        <div class="note">
            Análisis preliminar. Próximas versiones integrarán TRUNAJOD, MetaSistema, rúbricas humanas, benchmarks contextuales, captura avanzada del proceso y reportes exportables.
        </div>
    </div>
</div>

<script>
let sessionStarted = false;
let writingStarted = false;
let processClosed = false;

let startTime = null;
let firstInputTime = null;
let lastInputTime = null;
let timerInterval = null;

let longPauseCount = 0;
let editCount = 0;
let deletionCount = 0;
let insertionCount = 0;
let localAdjustmentCount = 0;
let reformulationCount = 0;
let expansionCount = 0;
let reductionCount = 0;
let macroAdjustmentCount = 0;
let inputEventCount = 0;

let previousText = "";
let maxTextLength = 0;
let previousParagraphCount = 0;

const LONG_PAUSE_THRESHOLD_MS = 3000;
const REFORMULATION_DELTA = 25;

const textarea = document.getElementById("textInput");

function startSessionIfNeeded() {
    if (!sessionStarted && !processClosed) {
        sessionStarted = true;
        startTime = Date.now();
        timerInterval = setInterval(updateTimer, 1000);
    }
}

function updateTimer() {
    if (startTime && !processClosed) {
        const seconds = Math.floor((Date.now() - startTime) / 1000);
        document.getElementById("timer").innerText = seconds;
    }
}

function paragraphCount(text) {
    return text.split("\\n").filter(p => p.trim().length > 0).length;
}

function updatePanel() {
    const now = Date.now();
    const totalTimeSeconds = startTime ? Math.floor((now - startTime) / 1000) : 0;
    const latencySeconds = firstInputTime && startTime ? Math.floor((firstInputTime - startTime) / 1000) : 0;

    document.getElementById("timer").innerText = totalTimeSeconds;
    document.getElementById("latency").innerText = latencySeconds;
    document.getElementById("longPauses").innerText = longPauseCount;
    document.getElementById("edits").innerText = editCount;
    document.getElementById("insertions").innerText = insertionCount;
    document.getElementById("deletions").innerText = deletionCount;
    document.getElementById("localAdjustments").innerText = localAdjustmentCount;
    document.getElementById("reformulations").innerText = reformulationCount;
    document.getElementById("expansions").innerText = expansionCount;
    document.getElementById("reductions").innerText = reductionCount;
    document.getElementById("macroAdjustments").innerText = macroAdjustmentCount;
    document.getElementById("events").innerText = inputEventCount;
}

textarea.addEventListener("focus", function() {
    startSessionIfNeeded();
});

textarea.addEventListener("paste", function() {
    document.getElementById("writingMode").value = "pasted";
});

textarea.addEventListener("input", function() {
    if (processClosed) {
        return;
    }

    startSessionIfNeeded();

    const now = Date.now();
    const currentText = textarea.value;
    const currentParagraphCount = paragraphCount(currentText);

    if (!writingStarted) {
        writingStarted = true;
        firstInputTime = now;
        previousParagraphCount = currentParagraphCount;
    }

    if (lastInputTime) {
        const gap = now - lastInputTime;
        if (gap >= LONG_PAUSE_THRESHOLD_MS) {
            longPauseCount += 1;
        }
    }

    inputEventCount += 1;

    const delta = currentText.length - previousText.length;

    if (delta > 0) {
        insertionCount += delta;
        editCount += 1;
        expansionCount += 1;

        if (delta <= 15) {
            localAdjustmentCount += 1;
        }

        if (delta >= REFORMULATION_DELTA) {
            reformulationCount += 1;
        }
    }

    if (delta < 0) {
        const absDelta = Math.abs(delta);
        deletionCount += absDelta;
        editCount += 1;
        reductionCount += 1;

        if (absDelta <= 15) {
            localAdjustmentCount += 1;
        }

        if (absDelta >= REFORMULATION_DELTA) {
            reformulationCount += 1;
        }
    }

    if (currentParagraphCount !== previousParagraphCount && inputEventCount > 1) {
        macroAdjustmentCount += 1;
    }

    if (currentText.length > maxTextLength) {
        maxTextLength = currentText.length;
    }

    previousText = currentText;
    previousParagraphCount = currentParagraphCount;
    lastInputTime = now;

    updatePanel();
});

function freezeProcess() {
    processClosed = true;

    if (timerInterval) {
        clearInterval(timerInterval);
        timerInterval = null;
    }

    updatePanel();
}

function resetProcessCapture() {
    sessionStarted = false;
    writingStarted = false;
    processClosed = false;

    startTime = null;
    firstInputTime = null;
    lastInputTime = null;

    longPauseCount = 0;
    editCount = 0;
    deletionCount = 0;
    insertionCount = 0;
    localAdjustmentCount = 0;
    reformulationCount = 0;
    expansionCount = 0;
    reductionCount = 0;
    macroAdjustmentCount = 0;
    inputEventCount = 0;

    previousText = "";
    maxTextLength = 0;
    previousParagraphCount = 0;

    if (timerInterval) {
        clearInterval(timerInterval);
        timerInterval = null;
    }

    document.getElementById("timer").innerText = "0";
    document.getElementById("latency").innerText = "0";
    document.getElementById("longPauses").innerText = "0";
    document.getElementById("edits").innerText = "0";
    document.getElementById("insertions").innerText = "0";
    document.getElementById("deletions").innerText = "0";
    document.getElementById("localAdjustments").innerText = "0";
    document.getElementById("reformulations").innerText = "0";
    document.getElementById("expansions").innerText = "0";
    document.getElementById("reductions").innerText = "0";
    document.getElementById("macroAdjustments").innerText = "0";
    document.getElementById("events").innerText = "0";

    document.getElementById("textInput").value = "";
    document.getElementById("result").style.display = "none";
}

async function runScriptora() {
    const selectedModule = document.getElementById("module").value;
    const writingMode = document.getElementById("writingMode").value;
    const text = document.getElementById("textInput").value;
    const level = document.getElementById("level").value;
    const genre = document.getElementById("genre").value;

    if (!text.trim()) {
        alert("Por favor escribe o pega un texto antes de analizar.");
        return;
    }

    const now = Date.now();
    const totalTimeSeconds = startTime ? Math.floor((now - startTime) / 1000) : 0;
    const initialLatencySeconds = firstInputTime && startTime ? Math.floor((firstInputTime - startTime) / 1000) : 0;

    freezeProcess();

    let endpoint = "/api/text/analyze";
    let payload = {
        text: text,
        language: "es",
        context: "scriptora_suite",
        level: level,
        genre: genre,
        purpose: "preliminary_analysis"
    };

    if (selectedModule === "write_product") {
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

    if (selectedModule === "write_process") {
        endpoint = "/api/write/process-evaluate";
        payload = {
            text: text,
            language: "es",
            level: level,
            genre: genre,
            task: "open_writing_task",
            purpose: "preliminary_process_writing_evaluation",
            writing_mode: writingMode,
            total_time_seconds: totalTimeSeconds,
            initial_latency_seconds: initialLatencySeconds,
            pause_count: longPauseCount,
            long_pause_count: longPauseCount,
            edit_count: editCount,
            deletion_count: deletionCount,
            insertion_count: insertionCount,
            local_adjustment_count: localAdjustmentCount,
            reformulation_count: reformulationCount,
            expansion_count: expansionCount,
            reduction_count: reductionCount,
            macro_adjustment_count: macroAdjustmentCount,
            max_text_length: maxTextLength || text.length,
            final_text_length: text.length,
            input_event_count: inputEventCount
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

    document.getElementById("processSection").style.display = "none";
    document.getElementById("integratedSection").style.display = "none";

    if (selectedModule === "text") {
        const metrics = data.raw_metrics;
        document.getElementById("productTitle").innerText = "Producto textual";

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

    if (selectedModule === "write_product") {
        renderWritingProduct(data.writing_metrics);
    }

    if (selectedModule === "write_process") {
        const product = data.product_metrics;
        const process = data.process_metrics;

        renderWritingProduct(product);

        document.getElementById("processSection").style.display = "block";

        document.getElementById("processMetrics").innerHTML = `
            <div class="metric"><strong>Modo de ingreso:</strong> ${process.writing_mode}</div>
            <div class="metric"><strong>Tiempo total:</strong> ${process.total_time_seconds} s</div>
            <div class="metric"><strong>Latencia inicial:</strong> ${process.initial_latency_seconds} s</div>
            <div class="metric"><strong>Pausas largas:</strong> ${process.long_pause_count}</div>
            <div class="metric"><strong>Ediciones:</strong> ${process.edit_count}</div>
            <div class="metric"><strong>Inserciones:</strong> ${process.insertion_count}</div>
            <div class="metric"><strong>Borrados:</strong> ${process.deletion_count}</div>
            <div class="metric"><strong>Ajustes locales:</strong> ${process.local_adjustment_count}</div>
            <div class="metric"><strong>Reformulaciones:</strong> ${process.reformulation_count}</div>
            <div class="metric"><strong>Expansiones:</strong> ${process.expansion_count}</div>
            <div class="metric"><strong>Reducciones:</strong> ${process.reduction_count}</div>
            <div class="metric"><strong>Ajustes macrotextuales:</strong> ${process.macro_adjustment_count}</div>
            <div class="metric"><strong>Eventos de escritura:</strong> ${process.input_event_count}</div>
            <div class="metric"><strong>Palabras por minuto:</strong> ${process.words_per_minute}</div>
            <div class="metric"><strong>Estabilidad final:</strong> ${process.final_stability_ratio}</div>
        `;

        document.getElementById("processScores").innerHTML = `
            <h3>Puntajes preliminares del proceso</h3>
            <div class="score-box">
                <div class="score"><strong>Planificación</strong><span>${process.planning_score}</span></div>
                <div class="score"><strong>Monitoreo</strong><span>${process.monitoring_score}</span></div>
                <div class="score"><strong>Revisión</strong><span>${process.revision_score}</span></div>
                <div class="score"><strong>Reformulación</strong><span>${process.reformulation_score}</span></div>
                <div class="score"><strong>Fluidez</strong><span>${process.fluency_score}</span></div>
                <div class="score"><strong>Recursividad</strong><span>${process.recursivity_score}</span></div>
                <div class="score"><strong>Regulación global</strong><span>${process.process_regulation_score}</span></div>
                <div class="score"><strong>Nivel proceso</strong><span>${process.process_regulation_label}</span></div>
            </div>
        `;

        document.getElementById("interpretation").innerText = process.interpretation;
        document.getElementById("integratedSection").style.display = "block";
        document.getElementById("integratedInterpretation").innerText = data.integrated_interpretation;
    }
}

function renderWritingProduct(metrics) {
    const scores = metrics.scores;

    document.getElementById("productTitle").innerText = "Producto escrito";

    document.getElementById("metrics").innerHTML = `
        <div class="metric"><strong>Palabras:</strong> ${metrics.word_count}</div>
        <div class="metric"><strong>Oraciones:</strong> ${metrics.sentence_count}</div>
        <div class="metric"><strong>Párrafos:</strong> ${metrics.paragraph_count}</div>
        <div class="metric"><strong>Longitud oracional promedio:</strong> ${metrics.avg_sentence_length}</div>
        <div class="metric"><strong>TTR:</strong> ${metrics.type_token_ratio}</div>
        <div class="metric"><strong>Densidad léxica:</strong> ${metrics.lexical_density_proxy}</div>
        <div class="metric"><strong>Conectores detectados:</strong> ${metrics.connector_count}</div>
        <div class="metric"><strong>Conectores:</strong> ${metrics.connectors_found.join(", ") || "No detectados"}</div>
        <div class="metric"><strong>Marcas de puntuación:</strong> ${metrics.punctuation_count}</div>
        <div class="metric"><strong>Cierre textual:</strong> ${metrics.closure_present ? "Detectado" : "No detectado"}</div>
    `;

    document.getElementById("scores").innerHTML = `
        <h3>Puntajes preliminares del producto</h3>
        <div class="score-box">
            <div class="score"><strong>Extensión</strong><span>${scores.extension}</span></div>
            <div class="score"><strong>Organización</strong><span>${scores.organization}</span></div>
            <div class="score"><strong>Cohesión</strong><span>${scores.cohesion}</span></div>
            <div class="score"><strong>Sintaxis</strong><span>${scores.syntax_control}</span></div>
            <div class="score"><strong>Léxico</strong><span>${scores.lexical_variety}</span></div>
            <div class="score"><strong>Densidad informativa</strong><span>${scores.informational_density}</span></div>
            <div class="score"><strong>Elaboración</strong><span>${scores.idea_elaboration}</span></div>
            <div class="score"><strong>Coherencia global</strong><span>${scores.global_coherence}</span></div>
            <div class="score"><strong>Puntuación</strong><span>${scores.punctuation_control}</span></div>
            <div class="score"><strong>Adecuación género</strong><span>${scores.genre_adequacy}</span></div>
            <div class="score"><strong>Cierre</strong><span>${scores.textual_closure}</span></div>
            <div class="score"><strong>Puntaje global</strong><span>${scores.global_writing_score}</span></div>
            <div class="score"><strong>Nivel</strong><span>${metrics.level_label}</span></div>
        </div>
    `;

    document.getElementById("interpretation").innerText = metrics.interpretation;
}
</script>
</body>
</html>
"""


# ============================================================
# ENDPOINTS
# ============================================================

@app.get("/api/health")
def health():
    return {
        "status": "ok",
        "service": "Scriptora Suite",
        "version": "0.6.0",
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
        "version": "0.6.0",
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

    writing_metrics = score_writing_product(
        text=text,
        words=words,
        sentences=sentences,
        genre=request.genre or "general"
    )

    return {
        "module": "Scriptora W · Writing Product Evaluation",
        "version": "0.6.0",
        "language": request.language,
        "level": request.level,
        "genre": request.genre,
        "task": request.task,
        "purpose": request.purpose,
        "writing_metrics": writing_metrics
    }


@app.post("/api/write/process-evaluate")
def evaluate_writing_process(request: WritingProcessRequest):
    text = request.text.strip()
    words = tokenize_words(text)
    sentences = split_sentences(text)

    product_metrics = score_writing_product(
        text=text,
        words=words,
        sentences=sentences,
        genre=request.genre or "general"
    )

    process_metrics = score_writing_process(request)

    integrated = integrated_writing_interpretation(
        product_metrics=product_metrics,
        process_metrics=process_metrics
    )

    return {
        "module": "Scriptora W · Product + Process Evaluation",
        "version": "0.6.0",
        "language": request.language,
        "level": request.level,
        "genre": request.genre,
        "task": request.task,
        "purpose": request.purpose,
        "product_metrics": product_metrics,
        "process_metrics": process_metrics,
        "integrated_interpretation": integrated
    }


@app.get("/api/benchmarks")
def list_benchmarks():
    return {
        "benchmarks": [
            {
                "benchmark_id": "ES_TEXT_GENERAL_V1",
                "module": "Scriptora T",
                "language": "es",
                "status": "prototype"
            },
            {
                "benchmark_id": "ES_WRITE_PRODUCT_V1",
                "module": "Scriptora W · Product",
                "language": "es",
                "status": "prototype"
            },
            {
                "benchmark_id": "ES_WRITE_PROCESS_REGULATION_V1",
                "module": "Scriptora W · Process",
                "language": "es",
                "status": "prototype"
            }
        ]
    }
