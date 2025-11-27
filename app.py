from flask import Flask, render_template, request, jsonify
import pandas as pd
import os

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

excel_data = {}

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    file = request.files["file"]
    path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(path)

    xls = pd.ExcelFile(path)
    sheets = {}

    for sheet in xls.sheet_names:
        df = xls.parse(sheet).fillna("")
        sheets[sheet] = df.to_dict(orient="records")

    global excel_data
    excel_data = sheets

    return jsonify({"sheets": list(sheets.keys())})

@app.route("/sheet")
def sheet():
    name = request.args.get("name")
    return jsonify(excel_data.get(name, []))

if __name__ == "__main__":
    app.run(debug=True)
