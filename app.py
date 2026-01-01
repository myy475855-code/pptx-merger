from flask import Flask, request, send_file, render_template, jsonify 
from io import BytesIO
import os
import tempfile
import shutil
import cloudmersive_convert_api_client
from cloudmersive_convert_api_client.rest import ApiException

app = Flask(__name__)
app.secret_key = "supersecretkey"

# -------------------------------
configuration = cloudmersive_convert_api_client.Configuration()
configuration.api_key['Apikey'] = ''  
api_instance = cloudmersive_convert_api_client.MergeDocumentApi(
    cloudmersive_convert_api_client.ApiClient(configuration)
)
# -------------------------------

def allowed_file(filename):
    return filename.lower().endswith(".pptx")

@app.route("/merge", methods=["POST"])
def merge_pptx_api():
    if "files[]" not in request.files:
        return jsonify({"error": "No files uploaded"}), 400

    files = request.files.getlist("files[]")
    name = request.form.get("outputName") or "merged"
    if len(files) < 2:
        return jsonify({"error": "Please upload at least 2 PPTX files"}), 400

    # Save uploaded files to temp folder
    temp_dir = tempfile.mkdtemp()
    file_paths = []

    try:
        for i, f in enumerate(files[:10]):  # Cloudmersive supports up to 10 files
            if allowed_file(f.filename):
                path = os.path.join(temp_dir, f"file{i+1}.pptx")
                f.save(path)
                file_paths.append(path)

        while len(file_paths) < 10:
            file_paths.append(None)

        try:
            api_response = api_instance.merge_document_pptx_multi(
                input_file1=file_paths[0],
                input_file2=file_paths[1],
                input_file3=file_paths[2],
                input_file4=file_paths[3],
                input_file5=file_paths[4],
                input_file6=file_paths[5],
                input_file7=file_paths[6],
                input_file8=file_paths[7],
                input_file9=file_paths[8],
                input_file10=file_paths[9]
            )
        except ApiException as e:
            return jsonify({"error": f"Cloudmersive API error: {e}"}), 500

        merged_file = BytesIO(api_response)
        merged_file.seek(0)
        return send_file(
            merged_file,
            as_attachment=True,
            download_name=f"{name}.pptx",
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    finally:
        # Delete the temporary upload folder after processing
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

@app.route("/")
def index():
    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
