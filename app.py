from pathlib import Path
from io import BytesIO
from tempfile import TemporaryDirectory

from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename

from campo_generator.generator import GenerationError, generate_field_book


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "data" / "templates" / "libro_campo_modelo.xlsm"

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/api/libro-campo")
def create_field_book():
    uploaded = request.files.get("source_file")
    if not uploaded or uploaded.filename == "":
        return {"error": "Subí un archivo Excel fuente."}, 400

    suffix = Path(uploaded.filename).suffix.lower()
    if suffix not in {".xlsx", ".xlsm"}:
        return {"error": "El archivo debe ser .xlsx o .xlsm."}, 400
    try:
        repetitions = int(request.form.get("repetitions", "4"))
    except ValueError:
        return {"error": "La cantidad de repeticiones debe ser un número."}, 400
    if repetitions < 1:
        return {"error": "La cantidad de repeticiones debe ser mayor a 0."}, 400

    with TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        source_path = tmp_dir / secure_filename(uploaded.filename)
        output_path = tmp_dir / "Libro de campo.xlsx"
        uploaded.save(source_path)

        try:
            metadata = generate_field_book(
                source_path=source_path,
                template_path=TEMPLATE_PATH,
                output_path=output_path,
                repetitions=repetitions,
            )
        except GenerationError as exc:
            return {"error": str(exc)}, 400

        download_name = f"Libro de campo - {metadata.study_name}.xlsx"
        file_bytes = BytesIO(output_path.read_bytes())

    return send_file(
        file_bytes,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
