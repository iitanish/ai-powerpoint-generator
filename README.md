# AI PowerPoint Generator

<div align="center">

(see the generated image above)

[
[
[![FastAPI](https://img.shields.io/badge/FastAPI-%230db7ed.svg?style=flat&logo=docker&logoColor=whitenthropic-191919.svg?style=flat&logo=anthropic&logoColor=-powered automation*

[🚀 **Demo**](https://ai-ppt-generator.demo.com) | [📖 **Documentation**](https://docs.ai-ppt-generator.com) | [🐛 **Report Bug**](https://github.com/yourusername/ai-powerpoint-generator/issues) | [✨ **Request Feature**](https://github.com/yourusername/ai-powerpoint-generator/issues)

</div>

***

## 🌟 Overview

An intelligent presentation generator that transforms raw text and Markdown into professionally formatted PowerPoint presentations while preserving your custom template's design and branding. Powered by cutting-edge AI models from OpenAI, Anthropic, and Google Gemini.

### ✨ Key Features

- 🎨 **Template Preservation**: Automatically inherits styles, colors, fonts, and layouts from uploaded `.pptx` or `.potx` templates
- 🤖 **Multi-Provider AI Support**: Compatible with OpenAI, Anthropic, and Google Gemini APIs
- 🧠 **Intelligent Content Structuring**: Uses LLM-powered slide splitting and bullet point generation with structured JSON output
- 🖼️ **Smart Image Reuse**: Intelligently places existing images from templates throughout the presentation
- 📝 **Flexible Input**: Supports both plain text and Markdown with optional presentation style guidance
- 🔒 **Privacy-First**: API keys are processed in memory only and never stored or logged

## 📸 Screenshots

(see the generated image above)

*Clean and intuitive web interface for generating presentations*

## 🛠️ Technology Stack

(see the generated image above)

<div align="center">

| Frontend | Backend | AI/ML | DevOps | Database |
|----------|---------|-------|--------|----------|
|  |  |  |  |  |
|  |  |  | ![GitHub Actions](https://img.shields.io/badge/github%20actions-%232671E5.svg?style=for-the-badge&logo=githubactions&logoColor=Backend**: FastAPI with `python-pptx` for PowerPoint manipulation
- **Frontend**: Vanilla HTML, CSS, and JavaScript
- **AI Providers**: OpenAI (GPT), Anthropic (Claude), Google Gemini
- **Containerization**: Docker for consistent deployment
- **Cloud Platforms**: Support for Render, Railway, Fly.io, Heroku, Hugging Face Spaces

## 🚀 Quick Start

### Prerequisites

- Python 3.8 or higher
- pip (Python package manager)
- Git
- An API key from OpenAI, Anthropic, or Google

### 💻 Local Development

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/ai-powerpoint-generator.git
   cd ai-powerpoint-generator
   ```

2. **Create and activate virtual environment**
   ```bash
   # Windows
   python -m venv .venv
   .venv\Scripts\activate
   
   # macOS/Linux
   python -m venv .venv
   source .venv/bin/activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the application**
   ```bash
   uvicorn app:app --reload
   ```

5. **Open your browser**
   Navigate to `http://localhost:8000`

### 🐳 Docker Setup

1. **Build the Docker image**
   ```bash
   docker build -t ai-powerpoint-generator .
   ```

2. **Run the container**
   ```bash
   docker run -p 8000:8000 ai-powerpoint-generator
   ```

3. **Access the application**
   Open `http://localhost:8000` in your browser

### ☁️ One-Click Deploy

[
[
[![Deploy to Heroku](https://www.herokucdn.com/deploy/button. Content**: Write or paste your text/Markdown content into the input field
2. **🎯 Add Guidance** (Optional): Provide instructions like "investor pitch deck" or "training presentation"  
3. **📎 Upload Template**: Select a `.pptx` or `.potx` file with your desired styling
4. **🔑 Configure AI**: Enter your API key for OpenAI, Anthropic, or Gemini
5. **⚡ Generate**: Click the generate button to create your presentation
6. **📥 Download**: Receive your professionally formatted `.pptx` file

### Input Examples

```markdown
# Company Overview
Our startup revolutionizes the way people create presentations...

## Market Analysis
- Market size: $2.1 billion
- Growth rate: 15% annually
- Target audience: Business professionals

## Financial Projections
We project $1M in revenue by year 2...
```

## 🏗️ Project Structure

```
ai-powerpoint-generator/
├── 📁 app/
│   ├── 📄 __init__.py
│   ├── 📄 main.py           # FastAPI application
│   ├── 📄 models.py         # Data models
│   ├── 📄 services.py       # Business logic
│   └── 📄 utils.py          # Utility functions
├── 📁 static/
│   ├── 📁 css/
│   ├── 📁 js/
│   └── 📁 images/
├── 📁 templates/
│   └── 📄 index.html        # Main UI template
├── 📁 tests/
│   ├── 📄 test_main.py
│   └── 📄 test_services.py
├── 📄 requirements.txt      # Python dependencies
├── 📄 Dockerfile          # Container configuration
├── 📄 Procfile            # Deployment configuration
├── 📄 docker-compose.yml  # Multi-container setup
└── 📄 README.md           # Project documentation
```

## 🧪 API Reference

### Generate Presentation

```http
POST /api/generate
```

**Request Body:**
```json
{
  "content": "Your presentation content...",
  "guidance": "investor pitch deck",
  "template": "base64_encoded_pptx",
  "api_key": "your_api_key",
  "provider": "openai"
}
```

**Response:**
```json
{
  "status": "success",
  "presentation_url": "/download/abc123.pptx",
  "slide_count": 12,
  "processing_time": 15.2
}
```

## 🤝 Contributing

We love contributions! Here's how you can help:

1. **🍴 Fork** the repository
2. **🌿 Create** a feature branch (`git checkout -b feature/AmazingFeature`)
3. **💾 Commit** your changes (`git commit -m 'Add some AmazingFeature'`)
4. **📤 Push** to the branch (`git push origin feature/AmazingFeature`)
5. **🔄 Open** a Pull Request

### Development Guidelines

- Follow PEP 8 style guide for Python code
- Add tests for new features
- Update documentation as needed
- Ensure all tests pass before submitting PR

## 📊 Performance & Limits

| Metric | Value |
|--------|-------|
| Max input size | 100,000 characters |
| Processing time | ~15-30 seconds |
| Max slides generated | 50 slides |
| Supported file formats | `.pptx`, `.potx` |
| Concurrent users | 100+ |

## 🔒 Privacy & Security

- **🔐 Zero Data Retention**: API keys are used in-memory only for active requests
- **📝 No Logging**: API keys and user content are never logged or stored
- **🏠 Self-Hosting Recommended**: Deploy on your own infrastructure for maximum privacy
- **🛡️ Secure Processing**: All communication uses HTTPS encryption

## 🐛 Troubleshooting

### Common Issues

**Q: "API key invalid" error**
```bash
# Ensure your API key is valid and has sufficient credits
export OPENAI_API_KEY="your-key-here"
```

**Q: Large files failing to process**
```bash
# Check file size limits and reduce content if necessary
# Maximum recommended input: 50,000 characters
```

**Q: Template not preserving formatting**
```bash
# Ensure template uses standard PowerPoint layouts
# Avoid heavily customized or corrupted template files
```

## 📈 Roadmap

- [ ] **v2.0**: Advanced template analysis and layout detection
- [ ] **v2.1**: Custom AI model fine-tuning support
- [ ] **v2.2**: Batch processing for multiple presentations
- [ ] **v2.3**: Real-time collaboration features
- [ ] **v3.0**: Integration with popular presentation platforms

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- FastAPI team for the excellent web framework
- python-pptx developers for PowerPoint manipulation capabilities
- OpenAI, Anthropic, and Google for providing powerful AI APIs
- The open-source community for continuous inspiration and support

## 📞 Support & Community

<div align="center">

[
[
[

**[📧 Email](mailto:support@ai-ppt-generator.com)** | **[💬 Discord](https://discord.gg/ai-ppt-generator)** | **[🐦 Twitter](https://twitter.com/aipptgen)**

</div>

***

<div align="center">

**⭐ Star this repo if you find it helpful!**

Made with ❤️ by the AI PowerPoint Generator team

</div>