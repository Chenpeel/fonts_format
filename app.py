import json
import uuid
from pathlib import Path

from flask import Flask, Response, jsonify, request, send_file, send_from_directory

from core import DEFAULT_FONTS, HANDLERS, ensure_fonts_installed, process_file

app = Flask(__name__, static_folder="static")

WORK_DIR = Path("/tmp/fonts-format")
UPLOAD_DIR = WORK_DIR / "uploads"
OUTPUT_DIR = WORK_DIR / "outputs"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/api/upload", methods=["POST"])
def upload():
    result = []
    for f in request.files.getlist("files"):
        if not f.filename:
            continue
        p = Path(f.filename)
        if p.suffix.lower() not in HANDLERS:
            continue
        fid = str(uuid.uuid4())
        f.save(UPLOAD_DIR / f"{fid}{p.suffix.lower()}")
        result.append({"id": fid, "name": f.filename, "ext": p.suffix.lower()})
    return jsonify(result)


@app.route("/api/process", methods=["POST"])
def process():
    data = request.get_json()
    files = data.get("files", [])
    fonts = {
        "chinese":  data.get("chinese",  DEFAULT_FONTS["chinese"]),
        "latin":    data.get("latin",     DEFAULT_FONTS["latin"]),
        "japanese": data.get("japanese",  DEFAULT_FONTS["japanese"]),
    }

    def stream():
        total = len(files)
        for i, item in enumerate(files):
            fid, name = item["id"], item["name"]
            src = UPLOAD_DIR / f'{fid}{Path(name).suffix.lower()}'

            yield f"data: {json.dumps({'type': 'start', 'name': name, 'index': i, 'total': total})}\n\n"

            if not src.exists():
                yield f"data: {json.dumps({'type': 'error', 'name': name, 'msg': 'upload not found'})}\n\n"
                continue
            try:
                out = process_file(src, OUTPUT_DIR, fonts, lambda _: None)
                # 重命名：保留原始文件名方便下载
                final = OUTPUT_DIR / f"{fid}_{Path(name).stem}_font_fixed{Path(name).suffix}"
                out.rename(final)
                yield f"data: {json.dumps({'type': 'done', 'name': name, 'did': final.name})}\n\n"
            except Exception as e:
                yield f"data: {json.dumps({'type': 'error', 'name': name, 'msg': str(e)})}\n\n"

        yield f"data: {json.dumps({'type': 'complete'})}\n\n"

    return Response(
        stream(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/api/download/<path:did>")
def download(did: str):
    p = OUTPUT_DIR / did
    if not p.exists():
        return "not found", 404
    # 还原用户可见文件名（去掉 uuid 前缀）
    visible = "_".join(did.split("_")[1:])
    return send_file(p, as_attachment=True, download_name=visible)


if __name__ == "__main__":
    ensure_fonts_installed()
    app.run(host="0.0.0.0", port=5000, threaded=True)
