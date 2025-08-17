from flask import Flask, request, render_template, send_file, flash
import os
from services.grading_service import grade_student

app = Flask(__name__)
app.secret_key = 'your-secret-key'  # Needed for flash messages

UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:
            if 'student_file' not in request.files or 'memo_file' not in request.files:
                flash('Please upload both files')
                return render_template("index.html")
            
            student_file = request.files["student_file"]
            memo_file = request.files["memo_file"]

            # Save uploaded files
            student_path = os.path.join(UPLOAD_FOLDER, "students.xlsx")
            memo_path = os.path.join(UPLOAD_FOLDER, "memo.xlsx")
            student_file.save(student_path)
            memo_file.save(memo_path)

            # Run grading
            output_path = os.path.join(UPLOAD_FOLDER, "Marking.xlsx")
            output = grade_student(student_path, memo_path, output_path)

            return send_file(output, as_attachment=True)
            
        except Exception as e:
            flash(f'An error occurred: {str(e)}')
            return render_template("index.html")

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)