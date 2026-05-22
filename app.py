from flask import Flask, render_template, request, redirect, url_for, session, jsonify
import json, os, config
from data_loader import load_dashboard_data

try:
    from werkzeug.utils import secure_filename
except ImportError:
    def secure_filename(f): return os.path.basename(f)

app = Flask(__name__, template_folder=os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates"))
app.secret_key = config.SECRET_KEY
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"]   = True
app.config["MAX_CONTENT_LENGTH"]      = 50 * 1024 * 1024   # 50 MB max upload

_cache = {}

def get_data(force=False):
    if force or "data" not in _cache:
        _cache["data"] = load_dashboard_data(
            config.DATA_FOLDER,
            config.TRACKER_FILES,
            onedrive_cfg=config.ONEDRIVE_CFG,
        )
    return _cache["data"]

@app.route("/")
def index():
    return redirect(url_for("login"))

@app.route("/login", methods=["GET","POST"])
def login():
    error = ""
    if request.method == "POST":
        u = request.form.get("username","")
        p = request.form.get("password","")
        if config.USERS.get(u) == p:
            session["user"] = u
            return redirect(url_for("dashboard"))
        error = "Invalid username or password."
    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))
    data = get_data()
    return render_template("dashboard.html", data=json.dumps(data), user=session["user"])

@app.route("/api/refresh", methods=["POST"])
def api_refresh():
    if "user" not in session:
        return jsonify({"error": "Not logged in"}), 401
    data = get_data(force=True)
    return jsonify({"last_updated": data["last_updated"], "errors": data["errors"]})

@app.route("/api/data")
def api_data():
    if "user" not in session:
        return jsonify({"error": "Not logged in"}), 401
    return jsonify(get_data())

@app.route("/api/upload", methods=["POST"])
def api_upload():
    if "user" not in session:
        return jsonify({"error": "Not logged in"}), 401

    # Build a reverse map: filename → AM label
    name_to_label = {v: k for k, v in config.TRACKER_FILES.items()}
    expected_names = list(config.TRACKER_FILES.values())

    uploaded = []
    skipped  = []
    errors   = []

    files = request.files.getlist("files")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "No files received"}), 400

    os.makedirs(config.DATA_FOLDER, exist_ok=True)

    for f in files:
        if not f.filename:
            continue
        fname = secure_filename(f.filename)
        if fname in name_to_label:
            dest = os.path.join(config.DATA_FOLDER, fname)
            try:
                f.save(dest)
                uploaded.append({"file": fname, "am": name_to_label[fname]})
            except Exception as e:
                errors.append("Could not save {}: {}".format(fname, str(e)))
        else:
            skipped.append(fname)

    if uploaded:
        try:
            get_data(force=True)
        except Exception as e:
            errors.append("Data reload error: " + str(e))

    return jsonify({
        "uploaded":       uploaded,
        "skipped":        skipped,
        "errors":         errors,
        "expected_files": expected_names,
    })

if __name__ == "__main__":
    get_data()
    app.run(host=config.HOST, port=config.PORT, debug=config.DEBUG)
