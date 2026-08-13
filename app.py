import os
import tempfile
import xml.etree.ElementTree as ET
import requests as http_requests
from flask import Flask, render_template, request, jsonify
from markitdown import MarkItDown
from urllib.parse import urlparse

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


def fetch_url(url):
    """Descarga contenido de una URL y retorna (contenido_bytes, extension, netloc)."""
    parsed = urlparse(url)
    if parsed.scheme not in ("http", "https"):
        raise ValueError("La URL debe comenzar con http:// o https://")

    response = http_requests.get(url, timeout=30, headers={
        "User-Agent": "Mozilla/5.0 (compatible; MarkItDown/1.0)"
    })
    response.raise_for_status()

    path = parsed.path
    ext = os.path.splitext(path)[1].lower() if "." in path else ""
    if not ext:
        content_type = response.headers.get("Content-Type", "")
        if "pdf" in content_type:
            ext = ".pdf"
        elif "word" in content_type or "docx" in content_type:
            ext = ".docx"
        else:
            ext = ".html"

    return response.content, ext, parsed.netloc


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        return jsonify({"error": "No se envio ningun archivo"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No se selecciono ningun archivo"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": f"Tipo de archivo no soportado. Extensiones permitidas: {', '.join(sorted(ALLOWED_EXTENSIONS))}"}), 400

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


@app.route("/convert-url", methods=["POST"])
def convert_url():
    data = request.get_json()
    if not data or "url" not in data:
        return jsonify({"error": "No se proporciono una URL"}), 400

    url = data["url"].strip()
    if not url:
        return jsonify({"error": "La URL esta vacia"}), 400

    try:
        content, ext, netloc = fetch_url(url)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except http_requests.exceptions.Timeout:
        return jsonify({"error": "Timeout: la URL tardo demasiado en responder"}), 400
    except http_requests.exceptions.RequestException as e:
        return jsonify({"error": f"Error al descargar la URL: {str(e)}"}), 400

    filename = netloc.replace(".", "_") + ext

    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
        tmp.write(content)
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
        "filename": filename,
        "url": url,
        "markdown": markdown_content,
    })


@app.route("/convert-sitemap", methods=["POST"])
def convert_sitemap():
    data = request.get_json()
    if not data or "sitemap_url" not in data:
        return jsonify({"error": "No se proporciono la URL del sitemap"}), 400

    sitemap_url = data["sitemap_url"].strip()
    max_pages = data.get("max_pages", 20)

    if not sitemap_url:
        return jsonify({"error": "La URL del sitemap esta vacia"}), 400

    # Descargar el sitemap XML
    try:
        response = http_requests.get(sitemap_url, timeout=30, headers={
            "User-Agent": "Mozilla/5.0 (compatible; MarkItDown/1.0)"
        })
        response.raise_for_status()
    except http_requests.exceptions.RequestException as e:
        return jsonify({"error": f"Error al descargar el sitemap: {str(e)}"}), 400

    # Parsear el XML del sitemap
    try:
        root = ET.fromstring(response.content)
        ns = {"sm": "http://www.sitemaps.org/schemas/sitemap/0.9"}
        urls = []

        for url_elem in root.findall(".//sm:url/sm:loc", ns):
            urls.append(url_elem.text.strip())

        # Intentar sin namespace si no encuentra nada
        if not urls:
            for url_elem in root.iter():
                if url_elem.tag.endswith("loc"):
                    if url_elem.text:
                        urls.append(url_elem.text.strip())

    except ET.ParseError:
        return jsonify({"error": "El contenido no es un sitemap XML valido"}), 400

    if not urls:
        return jsonify({"error": "No se encontraron URLs en el sitemap"}), 400

    urls = urls[:max_pages]

    # Convertir cada URL
    results = []
    errors = []
    md = MarkItDown()

    for page_url in urls:
        try:
            content, ext, netloc = fetch_url(page_url)

            with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                tmp.write(content)
                tmp_path = tmp.name

            try:
                convert_result = md.convert(tmp_path)
                parsed = urlparse(page_url)
                path_name = parsed.path.strip("/").replace("/", "_") or "index"
                filename = f"{netloc}_{path_name}.md"

                results.append({
                    "url": page_url,
                    "filename": filename,
                    "markdown": convert_result.text_content,
                })
            finally:
                os.unlink(tmp_path)

        except Exception as e:
            errors.append({
                "url": page_url,
                "error": str(e),
            })

    return jsonify({
        "total_urls": len(urls),
        "converted": len(results),
        "failed": len(errors),
        "results": results,
        "errors": errors,
    })


if __name__ == "__main__":
    app.run(debug=True, port=5000)
