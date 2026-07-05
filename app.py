import os
import logging
from flask import Flask, render_template, request, jsonify, session
from dotenv import load_dotenv

# Load local .env file if it exists
load_dotenv()

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "lifemate_ai_secret_encryption_key_12389!")

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Try importing the Google GenAI SDKs (supporting both new and old for compatibility)
GENAI_MODERN_AVAILABLE = False
GENAI_LEGACY_AVAILABLE = False

try:
    from google import genai
    from google.genai import types
    GENAI_MODERN_AVAILABLE = True
    logger.info("google-genai (modern SDK) is available.")
except ImportError:
    try:
        import google.generativeai as genai_legacy
        GENAI_LEGACY_AVAILABLE = True
        logger.info("google-generativeai (legacy SDK) is available.")
    except ImportError:
        logger.warning("No Gemini SDK could be imported. App will run in mock mode unless SDKs are installed.")

def get_api_key():
    """Retrieves the API key from session or environment variables."""
    # First check Flask session
    if "gemini_api_key" in session and session["gemini_api_key"]:
        return session["gemini_api_key"]
    # Then check environment variables
    key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    return key

def call_gemini(prompt):
    """Calls Gemini API using either modern, legacy SDK, or fallback mock mode."""
    api_key = get_api_key()
    if not api_key:
        raise ValueError("Gemini API key is not configured. Please set it in the settings panel or environment.")

    # 1. Try modern google-genai SDK
    if GENAI_MODERN_AVAILABLE:
        try:
            client = genai.Client(api_key=api_key)
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=prompt,
            )
            return response.text
        except Exception as e:
            logger.error(f"Error calling modern Gemini SDK: {e}")
            # Fall back to legacy if available, or raise error
            if not GENAI_LEGACY_AVAILABLE:
                raise e

    # 2. Try legacy google-generativeai SDK
    if GENAI_LEGACY_AVAILABLE:
        try:
            genai_legacy.configure(api_key=api_key)
            # Use gemini-1.5-flash for compatibility with older SDK packages
            model = genai_legacy.GenerativeModel('gemini-1.5-flash')
            response = model.generate_content(prompt)
            return response.text
        except Exception as e:
            logger.error(f"Error calling legacy Gemini SDK: {e}")
            raise e

    raise RuntimeError("No Google Gemini SDK is available on this system. Please check requirements.txt installation.")

# Routes
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/dashboard')
def dashboard():
    return render_template('dashboard.html')

@app.route('/result')
def result():
    # Render detailed results page for viewing or printing
    content = request.args.get('content', '')
    title = request.args.get('title', 'LifeMate AI Guide')
    return render_template('result.html', content=content, title=title)

# API Keys Configuration endpoints
@app.route('/api/check_key', methods=['GET'])
def check_key():
    has_key = get_api_key() is not None
    return jsonify({"configured": has_key})

@app.route('/api/set_key', methods=['POST'])
def set_key():
    data = request.json or {}
    key = data.get("api_key", "").strip()
    if not key:
        return jsonify({"success": False, "error": "API Key cannot be empty"}), 400
    
    session["gemini_api_key"] = key
    return jsonify({"success": True, "message": "API Key saved in session successfully."})

@app.route('/api/clear_key', methods=['POST'])
def clear_key():
    session.pop("gemini_api_key", None)
    return jsonify({"success": True, "message": "API Key cleared from session."})

# Feature processing endpoint
@app.route('/api/generate', methods=['POST'])
def generate():
    data = request.json or {}
    feature = data.get("feature")
    payload = data.get("data", {})

    if not feature:
        return jsonify({"success": False, "error": "Missing 'feature' parameter"}), 400

    prompt = ""

    # Build prompts for each feature
    if feature == "chat":
        user_msg = payload.get("message", "")
        prompt = f"""
        You are LifeMate AI, a personalized digital companion for smarter, healthier, and more connected living.
        Respond to the user's query with practical, actionable, and encouraging advice.
        If appropriate, link the response to sustainability, productivity, or wellness.
        Use clean Markdown styling with bullet points where necessary.
        
        User Query: {user_msg}
        """
        
    elif feature == "planner":
        tasks = payload.get("tasks", "")
        available_time = payload.get("available_time", "")
        wellness_goals = payload.get("wellness_goals", "None")
        
        prompt = f"""
        You are LifeMate AI's Daily Planner Assistant.
        Optimize the user's day based on their tasks, available time, and wellness objectives.
        
        User Inputs:
        - Tasks to accomplish: {tasks}
        - Total available productive hours: {available_time} hours
        - Wellness goals: {wellness_goals}
        
        Please provide your response in the following Markdown structure:
        
        # 📅 Your Optimized Schedule
        Create an hourly/time-block schedule allocating time for the tasks, including breaks and wellness.
        
        # 🔑 Priority Checklist
        Rank the tasks in order of importance (High, Medium, Low) with a brief justification.
        
        # ⚡ Productivity & Wellness Hacks
        Provide 3-4 specific tips to stay focused, manage stress, and integrate their wellness goals into this specific schedule.
        """

    elif feature == "budget":
        income = payload.get("income", "")
        expenses = payload.get("expenses", "")
        savings_goal = payload.get("savings_goal", "")
        
        prompt = f"""
        You are LifeMate AI's Budget Assistant.
        Analyze the user's finances and provide a smart budget breakdown and savings recommendations.
        
        User Inputs:
        - Monthly Income: {income}
        - Current Monthly Expenses: {expenses}
        - Monthly Savings Goal: {savings_goal}
        
        Please provide your response in the following Markdown structure:
        
        # 💰 Financial Diagnostics & Breakdown
        - Analyze the savings potential (Income - Expenses).
        - Break down the finances using the 50/30/20 rule (Needs, Wants, Savings) based on their inputs.
        
        # 🚀 Savings Recommendations (Target: {savings_goal})
        - Provide 3-5 specific, practical tips to reduce expenses.
        - Give a concrete plan to bridge the gap if they are falling short of their savings goal.
        
        # 📈 Financial Health Advice
        - Give 2-3 expert-level tips on emergency funds, smart investing, or debt management customized for their range.
        """

    elif feature == "wellness":
        age = payload.get("age", "")
        lifestyle = payload.get("lifestyle", "")
        goals = payload.get("goals", "")
        diet_pref = payload.get("diet_pref", "No restrictions")
        
        prompt = f"""
        You are LifeMate AI's Wellness & Lifestyle Guide.
        Create a comprehensive, healthy lifestyle plan.
        
        User Inputs:
        - Age: {age}
        - Current Lifestyle: {lifestyle}
        - Health/Fitness Goals: {goals}
        - Dietary Preferences/Restrictions: {diet_pref}
        
        Please provide your response in the following Markdown structure:
        
        # 🏃‍♂️ Personalized Exercise Plan
        - Recommend types of workouts, frequency, and intensity suitable for their age ({age}) and lifestyle ({lifestyle}).
        
        # 🥗 Diet & Nutrition Ideas
        - Suggest meal structures and dietary choices tailored for their goal ({goals}) and dietary preference ({diet_pref}).
        
        # 💤 Sleep & Recovery Guidance
        - Provide recommendations for optimizing sleep quality, duration, and recovery routines.
        """

    elif feature == "community":
        question = payload.get("question", "")
        
        prompt = f"""
        You are LifeMate AI's Community & Sustainability Guide.
        Help the user build eco-friendly habits and foster local cooperation.
        
        User Query/Interest: {question}
        
        Please provide your response in the following Markdown structure:
        
        # 🌍 Practical Sustainability Ideas
        - Provide 3-5 high-impact, easy-to-implement actions the user can take (e.g. energy/water saving, waste reduction).
        
        # 👥 Community Clean-up & Engagement Plan
        - Outline a step-by-step blueprint to organize neighbors or community members (e.g. neighborhood cleanup, composting hub, tool-sharing network).
        
        # 🌱 Long-term Green Impact
        - Explain the environmental and social benefits of these actions to inspire others.
        """

    elif feature == "emergency":
        emergency_type = payload.get("emergency_type", "")
        location = payload.get("location", "General")
        
        prompt = f"""
        You are LifeMate AI's Emergency Preparedness Guide.
        Provide immediate, clear, and lifesaving guidelines for a {emergency_type} scenario.
        
        Context:
        - Emergency Type: {emergency_type}
        - General Location/Context: {location}
        
        Please provide your response in the following Markdown structure:
        
        # 🚨 IMMEDIATE SAFETY ACTIONS
        - Give 3-5 bold, clear steps to take RIGHT NOW for safety. Keep them brief and high-priority.
        
        # 🎒 Emergency Checklist & Prep
        - List essential items to pack in a Go-Bag or prepare in the home (water, meds, docs, etc.).
        
        # 📞 Communication & Post-Crisis Guide
        - How to stay in touch with loved ones and what to do once the immediate danger has passed.
        """
    else:
        return jsonify({"success": False, "error": f"Unknown feature: {feature}"}), 400

    try:
        response_text = call_gemini(prompt)
        return jsonify({"success": True, "result": response_text})
    except ValueError as ve:
        return jsonify({"success": False, "error": str(ve), "needs_key": True}), 400
    except Exception as e:
        logger.error(f"Error processing generation request: {e}")
        return jsonify({"success": False, "error": f"Failed to generate response: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
