# AI-Powered PPT Generator 🎨📊

This project is a **Flask + Bootstrap web app** that uses **Google Gemini API** to automatically create professional PowerPoint presentations.

---

## ✨ Features
- 📝 Input **bulk text or markdown** and optional **guidance** (tone, audience, style).
- 📂 Upload your own **PPTX template** or let the app create a clean default theme.
- 🎨 Auto-generates **5 professional slides** with titles, highlights, bullet points, and speaker notes.
- 🌈 Preserves template **design and color scheme**, only replacing text.
- ⚡ Built with **Flask, Bootstrap, python-pptx**, and **Google Gemini API**.

---

## 🚀 Getting Started

### 1️⃣ Clone the Repo
```bash
git clone https://github.com/your-username/ai-ppt-generator.git
cd ai-ppt-generator

### 2️⃣ Install Dependencies 
python -m venv venv
source venv/bin/activate   # (Linux/macOS)
venv\Scripts\activate      # (Windows)

pip install -r requirements.txt

### 3️⃣ Add Google Gemini API Key

Create a .env file in the project root:

GEMINI_API_KEY=your_google_gemini_api_key
FLASK_SECRET=supersecretkey





### 4️⃣ Run the App
python app.py


App runs at: http://127.0.0.1:5000

### 🖥️ Usage

Open the web app in your browser.

Enter your bulk text / markdown.

(Optional) Add guidance (e.g., "Make it formal, business style").

(Optional) Upload a PPT template (.pptx).

Click Generate PPT → wait for AI to create your deck.

Download your AI-powered presentation.

### 📦 Tech Stack

Backend: Flask (Python)

Frontend: Bootstrap 5

AI: Google Gemini API (gemini-pro)

PPT Generation: python-pptx

### 📸 Demo


<img width="1919" height="974" alt="image" src="https://github.com/user-attachments/assets/34df42a2-100f-4105-a2e6-8051f3c6f227" />
