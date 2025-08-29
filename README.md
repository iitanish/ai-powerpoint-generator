
````markdown
# 🎨 Presentify – Auto-Generate Beautiful Presentations

Turn your **bulk text / Markdown** into a polished **PowerPoint** deck that perfectly matches your uploaded **template’s style** — colors, fonts, layouts, and images included.  
Bring your own LLM API key (OpenAI, Anthropic, or Gemini).  
⚡ **No AI image generation** — the app **reuses images** from your uploaded template/presentation.

---

## ✨ Features
- 📝 **Smart slide creation** – paste long text/markdown + optional one-line guidance (e.g., *“Investor Pitch Deck”*).  
- 🎨 **Template-aware design** – upload a `.pptx` or `.potx` template, all fonts, colors, and layouts are preserved.  
- 🤖 **Bring your own LLM API key** – supports:
  - ![OpenAI](https://img.shields.io/badge/OpenAI-API-412991?logo=openai&logoColor=white)  
  - ![Anthropic](https://img.shields.io/badge/Anthropic-Claude-000000?logo=anthropic&logoColor=white)  
  - ![Gemini](https://img.shields.io/badge/Google-Gemini-4285F4?logo=google&logoColor=white)  
- 📑 **Intelligent slide splitting & bulleting** (LLM-generated JSON structure).  
- 🖼 **Image reuse** – pulls images from your uploaded template and places them tastefully.  
- 💾 **One-click download** – export your new `.pptx` instantly.  

---

## 🛠️ Tech Stack
- **Backend**: FastAPI + [`python-pptx`](https://python-pptx.readthedocs.io/)  
- **Frontend**: HTML / CSS / JavaScript  
- **Providers**: OpenAI (Responses API), Anthropic (Messages API), Gemini (`google-genai`)  

---

## 🚀 Getting Started

### 1️⃣ Run locally
```bash
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
uvicorn app:app --reload
# open http://localhost:8000
````

### 2️⃣ Deploy easily

* Render / Railway / Fly / Heroku / Hugging Face Spaces (with included `Procfile`).
* **Docker**:

  ```bash
  docker build -t presentify .
  docker run -p 8000:8000 presentify
  ```

---

## 🔒 Privacy First

* API keys are entered in the form and used **only for that request** (kept in memory).
* Keys are **never stored** or logged.
* ✅ For maximum privacy, **self-host**.

---

## 📌 Notes & Limitations

* ⏳ Very large inputs are truncated to safe token limits before LLM calls.
* 🎯 Layout inference is heuristic — defaults to the closest *“Title + Content”* layout.
* 🖼 Background images might not always be detected — only visible picture shapes are reused.

---

## 📷 Demo (coming soon)

*(Add GIF/screenshots here to showcase workflow)*

---

## 🤝 Contributing

Pull requests are welcome! For major changes, please open an issue first to discuss.

---

## 📜 License

MIT © 2025 – Presentify Team

```

This adds:  
- Emojis 🎨🚀🔒 for quick scanning  
- Shields.io badges (logos for OpenAI, Anthropic, Google Gemini)  
- Clean **section dividers**  
- A **demo placeholder** (you can later add a GIF of the app in action)  
- More concise + professional formatting  

---

👉 Do you want me to also **design a matching sample `demo.gif` workflow outline** (step-by-step) so your README feels like a real product page?
```
