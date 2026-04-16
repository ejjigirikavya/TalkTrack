from flask import Flask, render_template, request, redirect
from pptx import Presentation
import difflib
import os

app = Flask(__name__)

# ---------- PPT TEXT ----------
def extract_ppt_text(file):
    prs = Presentation(file)
    text = ""
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text += shape.text + " "
    return text.lower()


# ---------- ACCURACY ----------
def calculate_accuracy(ppt_text, spoken_text):
    return round(difflib.SequenceMatcher(None, ppt_text, spoken_text).ratio() * 100, 2)


# ---------- FILLERS ----------
def count_fillers(text):
    fillers = ["um", "uh", "like", "you know", "basically", "and"]
    count = 0
    for word in fillers:
        count += text.lower().count(word)
    return count


# ---------- PAUSES ----------
def count_pauses(text):
    return text.count("...")


# ---------- SPEED (FIXED HUMAN RANGE) ----------
def calculate_wpm(text):
    words = len(text.split())

    wpm = words * 3  # approximation

    if wpm < 60:
        return 60
    elif wpm > 150:
        return 150
    else:
        return wpm


# ---------- AI FEEDBACK ----------
def generate_ai_feedback(acc, fillers, wpm, pauses):
    feedback = []

    if acc < 40:
        feedback.append("You need to follow PPT more closely.")
    else:
        feedback.append("Good clarity.")

    if fillers > 3:
        feedback.append("Avoid filler words.")

    if wpm > 110:
        feedback.append("Speak slower.")
    elif wpm < 70:
        feedback.append("Speak faster.")

    if pauses > 5:
        feedback.append("Too many pauses.")
    else:
        feedback.append("Good flow.")

    return feedback


# ---------- ROUTES ----------
@app.route('/')
def home():
    return render_template('index.html')


# ✅ LOGIN ROUTE FIXED
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        return redirect('/dashboard')
    return render_template('index.html')


@app.route('/dashboard')
def dashboard():
    return render_template('dashboard.html', done=False)


@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['ppt']
    file.save("uploaded.pptx")
    return redirect('/dashboard')


@app.route('/analyze', methods=['POST'])
def analyze():
    spoken_text = request.form['text'].lower()

    if len(spoken_text.strip()) == 0:
        return render_template('dashboard.html', error="No speech detected", done=False)

    ppt_text = extract_ppt_text("uploaded.pptx")

    accuracy = calculate_accuracy(ppt_text, spoken_text)
    fillers = count_fillers(spoken_text)
    pauses = count_pauses(spoken_text)
    wpm = calculate_wpm(spoken_text)

    feedback = generate_ai_feedback(accuracy, fillers, wpm, pauses)

    return render_template('dashboard.html',
                           spoken=spoken_text,
                           accuracy=accuracy,
                           fillers=fillers,
                           speed=wpm,
                           pauses=pauses,
                           feedback=feedback,
                           done=True)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))