
from flask import Flask, render_template, request, send_file
from docx import Document
from io import BytesIO
from collections import Counter
import re

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024

def terms(text):
    stop = set("the and with for are you your from that this will have has our their job work years experience strong ability skills knowledge required preferred using within across into".split())
    return {w for w in re.findall(r"[a-z][a-z0-9+#.-]{2,}", text.lower()) if w not in stop}

@app.get("/")
def home():
    return render_template("index.html")

@app.post("/analyze")
def analyze():
    jd = request.form.get("jd","").strip()
    file = request.files.get("resume")
    if not jd or not file or not file.filename.lower().endswith(".docx"):
        return render_template("index.html", error="Please provide a job description and a .docx resume.")
    raw = file.read()
    try:
        doc = Document(BytesIO(raw))
    except Exception:
        return render_template("index.html", error="Could not read Word resume.")
    content = "\n".join(p.text for p in doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            content += "\n" + " ".join(c.text for c in row.cells)
    keywords = terms(jd)
    found = sorted(keywords & terms(content))
    missing = sorted(keywords - terms(content))
    score = round(100 * len(found) / max(len(keywords),1))
    return render_template("index.html", score=score, found=found[:35], missing=missing[:35],
                           jd=jd, message="Keyword overlap is an estimate, not an official ATS score. Never add skills you cannot demonstrate.")

if __name__ == "__main__":
    app.run(debug=True)
