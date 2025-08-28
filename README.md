# AI-Powered PPT Generator 🎨📊
Flask + Bootstrap web app using Google Gemini API to automatically generate professional PowerPoint presentations.

✨ Features
- Input bulk text or markdown with optional guidance (tone, audience, style)
- Upload your own PPTX template or use the default theme
- Auto-generates 5 slides with titles, highlights, bullet points, and speaker notes
- Preserves template design and color scheme, replacing only text
- Built with Flask, Bootstrap, python-pptx, and Google Gemini API

🚀 Getting Started
1. Clone the repo: git clone https://github.com/your-username/ai-ppt-generator.git
2. Navigate to the folder: cd ai-ppt-generator
3. Install dependencies: python -m venv venv, source venv/bin/activate (Linux/macOS) or venv\Scripts\activate (Windows), pip install -r requirements.txt
4. Add Google Gemini API key: Create a .env file in project root with GEMINI_API_KEY=your_google_gemini_api_key and FLASK_SECRET=supersecretkey
5. Run the app: python app.py and open http://127.0.0.1:5000 in your browser

🖥️ Usage
- Enter bulk text or markdown
- (Optional) Add guidance: tone, style, audience
- (Optional) Upload PPT template (.pptx)
- Click Generate PPT and wait for AI to create your presentation
- Download your AI-powered PPT

📦 Tech Stack
- Backend: Flask (Python)
- Frontend: Bootstrap 5
- AI Engine: Google Gemini API (gemini-pro)
- PPT Generation: python-pptx

⚡ Notes
- Ensure Google Gemini API key is valid
- PPT templates must be .pptx
- Generated PPT preserves design & color scheme



💡 Contribution
- Open issues or submit pull requests for improvements, bug fixes, or new features
