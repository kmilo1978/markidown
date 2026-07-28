import os
import tempfile
from flask import Flask, render_template, request, jsonify
from markitdown import MarkItDown

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB max

ALLOWED_EXTENSIONS = {
    "pdf", "docx", "doc", "pptx", "ppt", "xlsx", "xls",
    "html", "htm", "txt", "csv", "json", "xml",
    "jpg", "jpeg", "png", "gif", "bmp", "tiff",
    "mp3", "wav", "m4a",
    "zip", "epub", "ipynb",
}


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        return jsonify({"error": "No se envió ningún archivo"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No se seleccionó ningún archivo"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": f"Tipo de archivo no soportado. Extensiones permitidas: {', '.join(sorted(ALLOWED_EXTENSIONS))}"}), 400

    # Guardar temporalmente el archivo para que markitdown lo procese
    suffix = os.path.splitext(file.filename)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        file.save(tmp.name)
        tmp_path = tmp.name

    try:
        md = MarkItDown()
        result = md.convert(tmp_path)
        markdown_content = result.text_content
    except Exception as e:
        return jsonify({"error": f"Error al convertir: {str(e)}"}), 500
    finally:
        os.unlink(tmp_path)

    return jsonify({
        "filename": file.filename,
        "markdown": markdown_content,
    })


if __name__ == "__main__":
    app.run(debug=True, port=5000)
