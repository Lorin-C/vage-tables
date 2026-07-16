from flask import Flask, render_template, request, send_file
from docx import Document
from datetime import datetime
import os

app = Flask(__name__)

TEMPLATE = "uploads/form2026.docx"

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    customer = request.form.get("customer", "")
    date = request.form.get("date", "")
    items = []

    names = request.form.getlist("dish[]")
    for i in range(47):
        items.append({
            "dish": names[i] if i < len(names) else ""
        })

    doc = Document(TEMPLATE)

    # form2026 is a table based document. Fill cells while keeping layout.
    table = doc.tables[0]
    # 仅填写菜品列，保持数量、单价、金额及模板格式不变
    row_index = 1
    for item in items:
        if row_index < len(table.rows):
            cells = table.rows[row_index].cells
            if len(cells) >= 1:
                cells[0].text = item["dish"]
        row_index += 1

    out = "output/蔬菜单_" + datetime.now().strftime("%Y%m%d%H%M%S") + ".docx"
    os.makedirs("output", exist_ok=True)
    doc.save(out)
    return send_file(out, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
