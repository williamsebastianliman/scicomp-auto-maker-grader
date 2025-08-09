from flask import Flask, request, jsonify, send_file, abort
import os
import shutil
from document_generator import make_cases, generate_docs

app = Flask(__name__)

@app.route('/generate', methods=['GET', 'POST'])
def generate():
    """
    HTTP endpoint to generate one or more document cases.
    Accepts optional query parameters or JSON body:
      - docs_count: number of cases to generate (default 1)
      - seed: integer seed to deterministically reproduce output (default -1)
    Returns a JSON payload with the list of generated case directories.
    """
    if request.method == 'POST':
        data = request.get_json(force=True) or {}
    else:
        data = request.args or {}
    docs_count = int(data.get('docs_count', 1))
    seed_val = data.get('seed')
    seed = int(seed_val) if seed_val is not None else -1
    out_dir = "output_api"
    os.makedirs(out_dir, exist_ok=True)
    cases = make_cases(docs_count=docs_count, out_root=out_dir, seed=seed)
    return jsonify({"cases": cases})

@app.route('/download/<case_id>', methods=['GET'])
def download_case(case_id):
    """
    Download a generated case directory as a zip file.
    The <case_id> should be the ID used during generation (zero‑padded).
    """
    out_dir = "output_api"
    dir_path = os.path.join(out_dir, f"Case_{case_id}")
    if not os.path.isdir(dir_path):
        abort(404, description="Case not found")
    archive_path = shutil.make_archive(dir_path, 'zip', dir_path)
    return send_file(archive_path, as_attachment=True, download_name=f"{case_id}.zip")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)