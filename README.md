# LifeMate AI

> **Tagline:** *"Your AI companion for smarter living and stronger communities."*

## 🌟 Overview
People use different apps for planning, health, budgeting, eco-friendly habits, and emergency guidance. This makes everyday decision-making fragmented and inefficient.

**LifeMate AI** brings these essential services together into one AI-powered assistant, helping individuals make better daily decisions while encouraging safer, healthier, and more connected communities.

---

## 🛠️ MVP Features
*   **🏠 Home Dashboard:** A modern, glassmorphic landing page introducing the LifeMate AI toolkit.
*   **💬 AI Smart Assistant:** Ask Gemini general life-management questions (e.g., study habits, exam preparations, general queries).
*   **📅 Daily Planner:** Input tasks and available hours to generate a structured timeline and customized productivity hacks.
*   **💰 Budget Assistant:** Input income and expenses to analyze savings and receive customized financial recommendations.
*   **❤️ Wellness Guide:** Input age, lifestyle, and fitness objectives to receive customized exercise plans, meal ideas, and sleep advice.
*   **🌍 Community & Sustainability:** Get eco-friendly habits ideas and templates to organize neighborhood clean-ups or tool-sharing systems.
*   **🚨 Emergency Guide:** Access immediate safety checklists, emergency Go-Bag packers, and post-crisis procedures for Floods, Fires, Earthquakes, and Heatwaves.

---

## 💻 Tech Stack
*   **Frontend:** HTML5, CSS3 (Vanilla glassmorphic design), JavaScript (ES6+ for dynamically loading views and parsing Markdown)
*   **Backend:** Python 3.14+, Flask (for server routing and coordinating inputs)
*   **AI Engine:** Google Gemini API (`gemini-2.5-flash`) via the modern `google-genai` client libraries
*   **Presentation Compilation:** `python-pptx` (generates the hackathon submission slide deck programmatically)

---

## 📂 Folder Structure
```
LifeMate-AI/
│
├── app.py                      # Flask server routes & Gemini API integration
├── requirements.txt            # Python dependencies (Flask, google-genai, python-pptx, dotenv)
├── README.md                   # Project documentation
├── generate_presentation.py    # Programmatic presentation compiler script
├── LifeMate_AI_Presentation.pptx # Compiled hackathon slide deck
│
├── templates/                  # Flask HTML views
│   ├── index.html              # Landing page
│   ├── dashboard.html          # Unified tool dashboard console
│   └── result.html             # Standalone report export / print viewer
│
└── static/                     # Assets & client logic
    ├── style.css               # Glassmorphic dark styling system
    └── script.js               # Client AJAX, SPA routes, and markdown renderer
```

---

## 🚀 Quick Start & Installation

### 1. Clone & Navigate
Place the codebase in your directory and navigate to the project root:
```bash
cd LifeMate-AI
```

### 2. Install Dependencies
Install Python requirements (Flask, google-genai, python-pptx, and python-dotenv) using:
```bash
pip install -r requirements.txt
```

### 3. Setup Gemini API Key
To connect the application to Gemini, you can configure your API Key in two ways:
*   **Local environment variable:** Add a `.env` file in the root directory and paste your key:
    ```env
    GEMINI_API_KEY=your_gemini_api_key_here
    FLASK_SECRET_KEY=custom_secret_key_for_sessions
    ```
*   **Interactive settings modal:** When running the dashboard, click the **"Gemini AI: Setup Key"** indicator badge in the top right. Enter your Gemini API Key in the field and save it. It will be stored securely in your local browser session for all requests.

### 4. Run the Application
Launch the Flask development server:
```bash
python app.py
```
Open your browser and navigate to: `http://127.0.0.1:5000`

---

## 📊 Presentation Slides
The hackathon submission presentation deck **`LifeMate_AI_Presentation.pptx`** is already compiled and available in the project root folder. 

If you make modifications to the system features and want to regenerate the slide deck content, run:
```bash
python generate_presentation.py
```
This will compile a professional, dark-themed 11-slide presentation matching the **Gen AI Academy - APAC Edition** hackathon submission guidelines.
