"""motor_ia.py - Fallback multi-IA + búsqueda (BM25) sobre la teoría subida."""
import os, re, json, math, unicodedata
from collections import Counter
import streamlit as st

TIMEOUT = 45  # segundos por proveedor (8 s era muy poco para Gemini con prompt largo)


def _key(nombre):
    """Lee la clave desde .streamlit/secrets.toml o variables de entorno."""
    try:
        v = st.secrets.get(nombre)
    except Exception:
        v = None
    return v or os.environ.get(nombre, "")


def _proveedores():
    lista = [
        {"name": "Google (Gemini 3.8 Flash)", "type": "google",
         "key": "GOOGLE_API_KEY", "model": "gemini-3.8-flash"},
        {"name": "Groq (GPT-OSS 120B)", "type": "openai", "key": "GROQ_API_KEY",
         "base_url": "https://api.groq.com/openai/v1", "model": "openai/gpt-oss-120b"},
        {"name": "Mistral", "type": "openai", "key": "MISTRAL_API_KEY",
         "base_url": "https://api.mistral.ai/v1", "model": "mistral-small-latest"},
        {"name": "OpenRouter (GPT-OSS free)", "type": "openai", "key": "OPENROUTER_API_KEY",
         "base_url": "https://openrouter.ai/api/v1", "model": "openai/gpt-oss-120b:free"},
        # Opcionales: solo se usan si pones la clave en secrets (y tienen saldo)
        {"name": "DeepSeek", "type": "openai", "key": "DEEPSEEK_API_KEY",
         "base_url": "https://api.deepseek.com", "model": "deepseek-chat"},
        {"name": "OpenAI", "type": "openai", "key": "OPENAI_API_KEY",
         "base_url": "https://api.openai.com/v1", "model": "gpt-4o-mini"},
    ]
    for p in lista:
        p["api_key"] = _key(p["key"])
    return [p for p in lista if p["api_key"]]  # salta los que no tienen clave


def generar_respuesta_con_fallback(prompt_text):
    errores = []
    provs = _proveedores()
    if not provs:
        raise Exception("No encontré ninguna API key. Revisa que exista el archivo "
                        ".streamlit/secrets.toml junto a app.py y que tenga tus claves.")
    for prov in provs:
        try:
            if prov["type"] == "openai":
                from openai import OpenAI
                client = OpenAI(api_key=prov["api_key"], base_url=prov["base_url"],
                                timeout=TIMEOUT, max_retries=0)
                r = client.chat.completions.create(
                    model=prov["model"], temperature=0.2,
                    messages=[{"role": "user", "content": prompt_text}])
                return r.choices[0].message.content, prov["name"]
            else:
                import google.generativeai as genai
                genai.configure(api_key=prov["api_key"])
                m = genai.GenerativeModel(
                    prov["model"],
                    generation_config={"temperature": 0.2,
                                       "response_mime_type": "application/json"})
                r = m.generate_content(prompt_text, request_options={"timeout": TIMEOUT})
                return r.text, prov["name"]
        except Exception as e:
            errores.append(f"[{prov['name']}]: {str(e)[:300]}")
    raise Exception("Ningún motor respondió.\n\n" + "\n\n".join(errores))


def extraer_json(texto):
    t = texto.strip()
    i, j = t.find("{"), t.rfind("}")
    if i == -1 or j == -1:
        raise ValueError("La IA no devolvió JSON: " + t[:200])
    return json.loads(t[i:j + 1])


# ---------------- Búsqueda en la teoría (BM25) ----------------
_STOP = set("de la el en y a los las un una por con para del al se que es su lo como "
            "mas o pero sus le ya este si porque esta entre cuando muy sin sobre "
            "tambien me hasta hay donde desde todo nos durante todos uno les ni "
            "contra otros ese eso ante esto antes algunos unos otro otra tanto".split())
_PAG = re.compile(r"\[(?:Página interna del PDF|Diapositiva)[:\s]*(\d+)\]")


def _tokens(texto):
    t = unicodedata.normalize("NFD", texto.lower())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    # recorte a 6 letras = "stemming" barato (plurales, apendicitis/apendicectomia...)
    return [w[:6] for w in re.findall(r"[a-z0-9]{3,}", t) if w not in _STOP]


def _trozos(texto, n=1500, solape=200):
    if len(texto) <= n:
        return [texto]
    return [texto[i:i + n] for i in range(0, len(texto), n - solape)]


def _firma(carpetas):
    f = []
    for c in carpetas:
        if os.path.isdir(c):
            for a in sorted(os.listdir(c)):
                if a.endswith(".txt"):
                    p = os.path.join(c, a)
                    f.append((p, os.path.getmtime(p)))
    return tuple(f)


@st.cache_resource(show_spinner=False)
def _indice(firma):
    chunks = []
    for ruta, _ in firma:
        nombre = os.path.basename(ruta)[:-4]
        with open(ruta, "r", encoding="utf-8") as fh:
            partes = _PAG.split(fh.read())  # [pre, num, texto, num, texto...]
        paginas = []
        for k in range(1, len(partes) - 1, 2):
            num, txt = partes[k], partes[k + 1].strip()
            if paginas and paginas[-1][0] == num:
                paginas[-1][1] += " " + txt  # pptx: une shapes de la misma diapositiva
            else:
                paginas.append([num, txt])
        for num, txt in paginas:
            for tr in _trozos(re.sub(r"\s+", " ", txt)):
                tk = _tokens(tr)
                if tk:
                    chunks.append({"archivo": nombre, "pagina": num, "texto": tr,
                                   "tf": Counter(tk), "len": len(tk)})
    df = Counter()
    for c in chunks:
        df.update(c["tf"].keys())
    N = len(chunks)
    idf = {t: math.log(1 + (N - n + 0.5) / (n + 0.5)) for t, n in df.items()}
    avg = (sum(c["len"] for c in chunks) / N) if N else 1
    return {"chunks": chunks, "idf": idf, "avg": avg}


def buscar_fragmentos(carpetas, consulta, k=6, max_chars=9000):
    """Devuelve los k fragmentos (archivo, página, texto) más relevantes."""
    idx = _indice(_firma(carpetas))
    q = set(_tokens(consulta))
    k1, b = 1.5, 0.75
    pts = []
    for c in idx["chunks"]:
        s = 0.0
        for t in q:
            f = c["tf"].get(t, 0)
            if f:
                s += idx["idf"][t] * f * (k1 + 1) / (f + k1 * (1 - b + b * c["len"] / idx["avg"]))
        if s > 0:
            pts.append((s, c))
    pts.sort(key=lambda x: x[0], reverse=True)
    res, total = [], 0
    for _, c in pts[:k]:
        if total + len(c["texto"]) > max_chars:
            break
        res.append(c)
        total += len(c["texto"])
    return res


def contexto_para_prompt(frags):
    return "\n\n".join(f"[FUENTE: {c['archivo']} | Página {c['pagina']}]\n{c['texto']}"
                       for c in frags)


def explicar_con_teoria(p, marcada, clave, carpetas):
    """p = dict de la pregunta. Devuelve (json_ia, motor, fragmentos_usados)."""
    consulta = p["pregunta"] + " " + " ".join(p["opciones"])
    frags = buscar_fragmentos(carpetas, consulta)
    ctx = contexto_para_prompt(frags) or "(no se encontraron fragmentos relevantes)"
    linea_clave = (f"Clave oficial del banco: {clave}" if clave != "N/A"
                   else "La clave no está en el banco: dedúcela tú.")
    prompt = f"""Eres un tutor médico para el ENAM. Basa tu respuesta en los FRAGMENTOS DE TEORÍA.

=== FRAGMENTOS DE TEORÍA ===
{ctx}
=== FIN DE FRAGMENTOS ===

Pregunta: {p['pregunta']}
Opciones: {chr(10).join(p['opciones'])}
El estudiante marcó: {marcada}
{linea_clave}

Tareas:
1. Explica por qué la opción correcta lo es, apoyándote en los fragmentos.
2. Explica por qué la opción marcada por el estudiante es incorrecta.
3. En "fuente_cita" pon archivo y página EXACTOS tomados de las etiquetas [FUENTE: ...] que usaste. No inventes citas.
   Si los fragmentos no cubren el tema, escribe "No cubierto en tu teoría" y explica con conocimiento médico general, indicándolo.

Devuelve SOLO un JSON: {{"correcta": "Letra", "explicacion": "...", "fuente_cita": "Archivo, Página N"}}"""
    texto, motor = generar_respuesta_con_fallback(prompt)
    return extraer_json(texto), motor, frags
