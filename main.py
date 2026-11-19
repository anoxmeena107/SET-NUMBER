from flask import Flask, request, render_template_string, Response, send_file
from openpyxl import load_workbook
from io import BytesIO
import csv
import re
import socket
import os

app = Flask(__name__)

# ============================================================
#  BASE LAYOUT – shared shell, upar 2 bots ka selector
# ============================================================
BASE_TEMPLATE = """<!DOCTYPE html>
<html lang="hi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{{ title }}</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; background: #0f172a; color: #e2e8f0; min-height: 100vh; }

  /* Top bar with Bot selector */
  .topbar { background: #1e293b; border-bottom: 1px solid #334155; padding: 14px 20px; display: flex; align-items: center; gap: 10px; flex-wrap: wrap; position: sticky; top: 0; z-index: 10; }
  .brand { font-weight: 800; color: #fff; font-size: 17px; margin-right: 14px; letter-spacing: .3px; }
  .nav-btn { text-decoration: none; color: #cbd5e1; padding: 8px 18px; border-radius: 999px; font-size: 14px; border: 1px solid transparent; transition: all .15s ease; }
  .nav-btn:hover { background: #334155; color: #fff; }
  .nav-btn.active { background: #2563eb; color: #fff; box-shadow: 0 2px 10px rgba(37,99,235,.35); }

  .container { max-width: 920px; margin: 28px auto; padding: 0 20px; }
  .card { background: #1e293b; border: 1px solid #334155; border-radius: 14px; padding: 22px; margin-bottom: 20px; }

  h2 { color: #fff; font-size: 22px; margin-bottom: 14px; }
  h3 { color: #f1f5f9; margin: 4px 0 10px; font-size: 16px; }

  .hint { color: #94a3b8; font-size: 13px; margin-top: 12px; }
  input[type=file] { color: #cbd5e1; background: #0f172a; border: 1px dashed #475569; border-radius: 8px; padding: 10px; width: 100%; }
  textarea { width: 100%; height: 200px; background: #0f172a; color: #e2e8f0; border: 1px solid #475569; border-radius: 8px; padding: 12px; font-family: Consolas, monospace; font-size: 13px; }
  textarea:focus, input[type=file]:focus { outline: 2px solid #2563eb; }

  .btn { padding: 10px 20px; border: 0; border-radius: 8px; cursor: pointer; font-size: 14px; font-weight: 600; color: #fff; transition: filter .15s ease; text-decoration: none; display: inline-block; }
  .btn:hover { filter: brightness(1.1); }
  .btn-blue { background: #2563eb; }
  .btn-green { background: #16a34a; }

  pre { background: #0f172a; border: 1px solid #334155; border-radius: 10px; padding: 14px; overflow-x: auto; white-space: pre-wrap; word-break: break-word; font-size: 13px; color: #dbeafe; }
  .empty { color: #64748b; font-style: italic; }

  .stats { display: flex; gap: 12px; flex-wrap: wrap; margin: 12px 0 4px; }
  .stat { background: #0f172a; border: 1px solid #334155; border-radius: 10px; padding: 10px 16px; font-size: 13px; }
  .stat b { color: #fff; font-size: 18px; display: block; }

  .tag { display: inline-block; padding: 3px 12px; border-radius: 999px; font-size: 13px; font-weight: 600; }
  .tag-blue { background: #1d4ed8; color: #dbeafe; }
  .tag-green { background: #15803d; color: #dcfce7; }
  .tag-orange { background: #c2410c; color: #ffedd5; }

  /* Home bot cards */
  .bot-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 20px; }
  .bot-card { background: linear-gradient(160deg, #1e293b, #0f172a); border: 1px solid #334155; border-radius: 16px; padding: 26px; text-align: center; transition: transform .15s ease, border-color .15s ease; }
  .bot-card:hover { transform: translateY(-4px); border-color: #3b82f6; }
  .bot-icon { font-size: 44px; margin-bottom: 12px; }
  .bot-card h3 { color: #fff; margin: 10px 0 8px; }
  .bot-card p { color: #94a3b8; font-size: 14px; line-height: 1.55; margin-bottom: 18px; }

  .row { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; margin-top: 12px; }
</style>
</head>
<body>

  <!-- Bot selector: yahaan se doono bots ke beech switch karo -->
  <nav class="topbar">
    <span class="brand">🤖 Dual Tool Hub</span>
    <a class="nav-btn {{ 'active' if active == 'home' else '' }}" href="/">🏠 Home</a>
    <a class="nav-btn {{ 'active' if active == 'sort' else '' }}" href="/sort">📄 Bot 1 · Pipe Sorting</a>
    <a class="nav-btn {{ 'active' if active == 'extract' else '' }}" href="/extract">📱 Bot 2 · Code Extractor</a>
  </nav>

  {{ body | safe }}

</body>
</html>
"""

# ============================================================
#  HOME PAGE – 2 bots ke cards
# ============================================================
HOME_BODY = """
<div class="container">
  <h2>एक script, दो bots — जिसका उपयोग करना हो उसे चुनें 🚀</h2>
  <div class="bot-grid">
    <div class="bot-card">
      <div class="bot-icon">📄</div>
      <h3>Bot 1 · Pipe Lines Sorting</h3>
      <p>TXT / CSV फ़ाइल अपलोड करें। 3 pipes (|) वाली lines ऊपर, 2 pipes वाली नीचे छाँटकर download करें।</p>
      <a class="btn btn-blue" href="/sort">इस bot का उपयोग करें →</a>
    </div>
    <div class="bot-card">
      <div class="bot-icon">📱</div>
      <h3>Bot 2 · Phone Code Extractor</h3>
      <p>TXT / CSV / XLSX / XLSM फ़ाइलों या pasted text से phone numbers और उनके आगे मौजूद codes (5–8 अंक) निकालें।</p>
      <a class="btn btn-green" href="/extract">इस bot का उपयोग करें →</a>
    </div>
  </div>
</div>
"""

# ============================================================
#  BOT 1 – Pipe Lines Sorting
# ============================================================
SORT_BODY = """
<div class="container">
  <div class="card">
    <h2>📄 Bot 1 – Pipe Lines Sorting</h2>
    <form method="POST" enctype="multipart/form-data">
      <input type="file" name="file" accept=".txt,.csv" required>
      <div class="row">
        <button class="btn btn-blue" type="submit">⬆️ Upload & Sort</button>
      </div>
    </form>
    <p class="hint">नियम: <b>3 pipes (|)</b> वाली lines ऊपर, <b>2 pipes</b> वाली नीचे, बाकी सबसे नीचे।</p>
  </div>

  {% if stats %}
  <div class="card">
    <h3>📊 Stats</h3>
    <div class="stats">
      <div class="stat">Total lines <b>{{ stats.total }}</b></div>
      <div class="stat">3 pipes <b>{{ stats.pipe3 }}</b></div>
      <div class="stat">2 pipes <b>{{ stats.pipe2 }}</b></div>
      <div class="stat">Others <b>{{ stats.other }}</b></div>
    </div>
    <a class="btn btn-green" href="/sort/download" style="margin-top:14px;">⬇️ Download sorted file</a>
  </div>

  <div class="card">
    <h3><span class="tag tag-blue">3 pipes (top)</span></h3>
    {% if pipe3 %}
      <pre>{% for line in pipe3 %}{{ line }}
{% endfor %}</pre>
    {% else %}
      <p class="empty">No lines with 3 pipes found.</p>
    {% endif %}
  </div>

  <div class="card">
    <h3><span class="tag tag-green">2 pipes (below)</span></h3>
    {% if pipe2 %}
      <pre>{% for line in pipe2 %}{{ line }}
{% endfor %}</pre>
    {% else %}
      <p class="empty">No lines with 2 pipes found.</p>
    {% endif %}
  </div>

  {% if other %}
  <div class="card">
    <h3><span class="tag tag-orange">Other lines (1 pipe or no pipe)</span></h3>
    <pre>{% for line in other %}{{ line }}
{% endfor %}</pre>
  </div>
  {% endif %}
  {% endif %}
</div>
"""

last_sorted = {"pipe3": [], "pipe2": [], "other": [], "total": 0}

@app.route("/sort", methods=["GET", "POST"])
def sort_bot():
    if request.method == "POST":
        f = request.files.get("file")
        if f and f.filename:
            content = f.read().decode("utf-8", errors="ignore")
            lines = content.splitlines()
            pipe3, pipe2, other = [], [], []
            for line in lines:
                c = line.count("|")
                if c == 3:
                    pipe3.append(line)
                elif c == 2:
                    pipe2.append(line)
                else:
                    other.append(line)
            last_sorted["pipe3"] = pipe3
            last_sorted["pipe2"] = pipe2
            last_sorted["other"] = other
            last_sorted["total"] = len(lines)

            body = render_template_string(
                SORT_BODY,
                stats={"total": len(lines), "pipe3": len(pipe3),
                       "pipe2": len(pipe2), "other": len(other)},
                pipe3=pipe3, pipe2=pipe2, other=other,
            )
            return page("Bot 1 – Pipe Sorting", "sort", body)

    body = render_template_string(SORT_BODY, stats=None, pipe3=[], pipe2=[], other=[])
    return page("Bot 1 – Pipe Sorting", "sort", body)


@app.route("/sort/download")
def sort_download():
    """Download sorted file: 3-pipe lines first, then 2-pipe, then others."""
    sorted_lines = (
        last_sorted["pipe3"]
        + last_sorted["pipe2"]
        + last_sorted["other"]
    )
    return Response(
        "\n".join(sorted_lines),
        mimetype="text/plain",
        headers={"Content-Disposition": "attachment; filename=sorted_pipe_lines.txt"},
    )

# ============================================================
#  BOT 2 – Phone + Code Extractor
# ============================================================
EXTRACT_BODY = """
<div class="container">
  <div class="card">
    <h2>📱 Bot 2 – Phone + Code Extractor</h2>
    <form method="POST" enctype="multipart/form-data">
      <p style="color:#94a3b8;font-size:14px;margin-bottom:8px;">TXT, CSV, XLSX या XLSM फ़ाइलें चुनें:</p>
      <input type="file" name="files" multiple accept=".txt,.csv,.xlsx,.xlsm">
      <p style="color:#94a3b8;font-size:14px;margin:14px 0 8px;">या यहाँ text paste करें:</p>
      <textarea name="pasted" placeholder="Paste data here...">{{ pasted or '' }}</textarea>
      <div class="row">
        <button class="btn btn-green" type="submit">🔍 Extract Now</button>
      </div>
    </form>
  </div>

  {% if result is not none %}
  <div class="card">
    <h3>✅ Extracted: {{ count }} pairs</h3>
    {% if result %}
      <pre>{{ result }}</pre>
      <form method="POST" action="/extract/download" class="row">
        <input type="hidden" name="result" value="{{ result }}">
        <button class="btn btn-blue" type="submit">⬇️ Download TXT</button>
      </form>
    {% else %}
      <p class="empty">कोई phone-code pair नहीं मिला।</p>
    {% endif %}
  </div>
  {% endif %}
</div>
"""


def normalize_digits(text):
    """Arabic/Persian digits ko English digits me convert karo."""
    table = str.maketrans(
        "٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹",
        "01234567890123456789"
    )
    return text.translate(table)


def cell_to_text(value):
    """Excel cell value ko text banata hai (float 12.0 -> '12')."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def read_file(file):
    """TXT / CSV / XLSX / XLSM file ko plain text me convert karo."""
    filename = file.filename.lower()
    data = file.read()

    if filename.endswith((".xlsx", ".xlsm")):
        workbook = load_workbook(BytesIO(data), read_only=True, data_only=True)
        lines = []
        for sheet in workbook.worksheets:
            for row in sheet.iter_rows(values_only=True):
                line = " ".join(cell_to_text(v) for v in row if v is not None)
                if line.strip():
                    lines.append(line)
        workbook.close()
        return "\n".join(lines)

    if filename.endswith(".csv"):
        decoded = data.decode("utf-8-sig", errors="ignore")
        rows = csv.reader(decoded.splitlines())
        return "\n".join(" ".join(row) for row in rows)

    return data.decode("utf-8-sig", errors="ignore")


def extract_pairs(text):
    """Phone numbers aur unke baad wale 5–8 digit codes pair banata hai."""
    text = normalize_digits(text)

    # Date aur time formats hatao (galti se code na ban jayein)
    text = re.sub(
        r"\b(?:\d{4}[-/]\d{1,2}[-/]\d{1,2}|"
        r"\d{1,2}[-/]\d{1,2}[-/]\d{2,4})\b",
        " ", text
    )
    text = re.sub(r"\b\d{1,2}:\d{2}(?::\d{2})?\b", " ", text)

    # Phone: 10–15 digits, optional + aage
    phone_pattern = re.compile(r"(?<!\d)\+?\d{10,15}(?!\d)")
    # Code: exactly 5, 6, 7 ya 8 digits
    code_pattern = re.compile(r"(?<!\d)\d{5,8}(?!\d)")

    phones = list(phone_pattern.finditer(text))
    results = []
    seen = set()

    for index, phone_match in enumerate(phones):
        number = phone_match.group().lstrip("+")
        search_start = phone_match.end()
        search_end = phones[index + 1].start() if index + 1 < len(phones) else len(text)
        code_match = code_pattern.search(text[search_start:search_end])
        if not code_match:
            continue
        pair = f"{number}|{code_match.group()}"
        if pair not in seen:
            seen.add(pair)
            results.append(pair)

    return results


@app.route("/extract", methods=["GET", "POST"])
def extract_bot():
    result = None
    count = 0
    pasted = ""

    if request.method == "POST":
        pasted = request.form.get("pasted", "")
        texts = [pasted] if pasted.strip() else []
        for f in request.files.getlist("files"):
            if f and f.filename:
                try:
                    texts.append(read_file(f))
                except Exception:
                    continue
        pairs = extract_pairs("\n".join(texts))
        result = "\n".join(pairs)
        count = len(pairs)

    body = render_template_string(EXTRACT_BODY, result=result, count=count, pasted=pasted)
    return page("Bot 2 – Code Extractor", "extract", body)


@app.route("/extract/download", methods=["POST"])
def extract_download():
    result = request.form.get("result", "")
    buf = BytesIO(result.encode("utf-8"))
    buf.seek(0)
    return send_file(
        buf,
        mimetype="text/plain",
        as_attachment=True,
        download_name="number_code.txt",
    )

# ============================================================
#  HELPERS + MAIN
# ============================================================
def page(title, active, body):
    """Ek page render karo: base layout + body."""
    return render_template_string(BASE_TEMPLATE, title=title, active=active, body=body)


def get_local_ip():
    """Apna LAN IP nikalo (same WiFi wale access kar saken)."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))  # koi data send nahi hota, sirf route check
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


@app.route("/")
def home():
    return page("Home – Dual Tool Hub", "home", HOME_BODY)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    local_ip = get_local_ip()

    print()
    print("=" * 60)
    print("  ✅ Combined app ready: 2 Bots (Pipe Sorting + Code Extractor)")
    print()
    print(f"  🏠 Local (khud ka PC):  http://127.0.0.1:{port}/")
    print(f"  📱 Network (WiFi wale):  http://{local_ip}:{port}/")
    print()
    print("  💡 Public URL chahiye? Ngrok chalayein:")
    print(f"      ngrok http {port}")
    print("=" * 60)
    print()

    app.run(host="0.0.0.0", port=port, debug=False)
