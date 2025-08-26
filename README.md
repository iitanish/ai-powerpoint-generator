# 🚀 AI PowerPoint Generator Pro

<div align="center">

![AI PowerPoint Generator](https://img.shields.io/badge/AI-PowerPoint%20Generator-blue?style=for-the-badge&logo=microsoft-powerpoint)
![FastAPI](https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi)
![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

**Transform your text into professional presentations with AI-powered content generation and complete template style preservation**

[Demo](#demo) • [Features](#features) • [Quick Start](#quick-start) • [Documentation](#documentation) • [Contributing](#contributing)

</div>

## ✨ Features

### 🧠 **Intelligent Content Processing**

- **Multi-LLM Support**: Works with OpenAI GPT, Anthropic Claude, and Google Gemini
- **Smart Content Structuring**: Automatically organizes text into logical slides with balanced content
- **Topic Detection**: Automatically detects content category and applies appropriate themes
- **Dynamic Slide Count**: Creates 4-8 slides based on content length and complexity

### 🎨 **Complete Template Style Preservation**

- **Advanced Style Extraction**: Preserves colors, fonts, layouts, and visual elements from uploaded templates
- **Brand Consistency**: Maintains your organization's visual identity across all generated slides
- **Image Reuse**: Strategically incorporates existing template images and graphics
- **Professional Typography**: Applies original template fonts and formatting hierarchies

### 🛡️ **Enterprise-Grade Reliability**

- **Comprehensive Error Handling**: Graceful handling of all failure scenarios with detailed logging
- **Input Validation**: Advanced validation for all user inputs and file uploads
- **Rate Limiting**: Built-in protection against abuse with configurable limits
- **Security Features**: File integrity verification and input sanitization

### 🎯 **Production-Ready Architecture**

- **Retry Mechanisms**: Automatic retry with exponential backoff for API failures
- **Enhanced Logging**: Structured logging with color-coded output and error tracking
- **Health Monitoring**: Built-in health checks and metrics endpoints
- **Scalable Design**: Clean architecture supporting future enhancements

## 🚀 Quick Start

### Prerequisites

- Python 3.8 or higher
- PowerPoint template file (.pptx or .potx)
- API key from one of the supported providers:
  - [OpenAI API Key](https://platform.openai.com/api-keys)
  - [Anthropic API Key](https://console.anthropic.com/)
  - [Google Gemini API Key](https://makersuite.google.com/app/apikey)

### Installation

1. **Clone the repository**

   ```
   git clone https://github.com/yourusername/ai-powerpoint-generator.git
   cd ai-powerpoint-generator
   ```

2. **Install dependencies**

   ```
   # Using pip
   pip install -r requirements.txt

   # Or using uv (recommended)
   pip install uv
   uv pip install -r requirements.txt
   ```

3. **Start the application**

   ```
   # Development mode
   uvicorn main:app --reload --host 127.0.0.1 --port 8000

   # Or using uv
   uv run uvicorn main:app --reload --host 127.0.0.1 --port 8000
   ```

4. **Open your browser**
   Navigate to `http://localhost:8000` to access the web interface.

### 🎯 Basic Usage

1. **Prepare your content**: Write or paste your text content (minimum 50 characters)
2. **Upload template**: Select a PowerPoint template file (.pptx or .potx)
3. **Add guidance** (optional): Provide specific instructions for tone or structure
4. **Enter API key**: Your OpenAI, Anthropic, or Gemini API key
5. **Generate**: Click "Generate Presentation" and wait for processing
6. **Download**: Download your professionally formatted presentation

## 📋 Requirements

### System Requirements

```
fastapi==0.104.1
uvicorn[standard]==0.24.0
jinja2==3.1.2
python-multipart==0.0.6
python-pptx==0.6.23
requests==2.31.0
aiofiles==23.2.1
pillow==10.1.0
```

### File Requirements

- **Template files**: .pptx or .potx format, maximum 15MB
- **Text content**: 50-50,000 characters
- **Supported layouts**: Any PowerPoint template with standard layouts

## 🏗️ Architecture

```
ai-powerpoint-generator/
├── main.py                 # FastAPI application with all endpoints
├── requirements.txt        # Python dependencies
├── templates/
│   ├── index.html         # Web interface
│   ├── 404.html          # Error pages
│   └── 500.html
├── static/
│   └── style.css         # Styling
├── uploads/              # Temporary template storage
├── generated/            # Generated presentations
└── logs/                 # Application logs
```

### Core Components

- **`EnhancedLLMProcessor`**: Handles multi-provider LLM API calls with retry logic
- **`TemplateStyleExtractor`**: Extracts complete styling information from templates
- **`AdvancedPowerPointGenerator`**: Creates presentations with full style preservation
- **`InputValidator`**: Comprehensive input validation and sanitization
- **Custom Exception Classes**: Granular error handling and user feedback

## 📊 API Endpoints

### Main Endpoints

- `GET /` - Web interface
- `POST /generate` - Generate presentation
- `GET /download/{session_id}` - Download generated presentation
- `GET /health` - Health check and system status

### Monitoring

- `GET /metrics` - Application metrics
- `GET /template-info/{session_id}` - Template analysis details

## 🔧 Configuration

### Environment Variables

```
# Optional configuration
PRODUCTION=false              # Enable production mode
MAX_FILE_SIZE=15728640       # Maximum file size (15MB)
MAX_TEXT_LENGTH=50000        # Maximum text length
CLEANUP_INTERVAL=3600        # File cleanup interval (1 hour)
```

### Rate Limiting

- **Generation**: 5 requests per 5 minutes per IP
- **Downloads**: 20 requests per 5 minutes per IP

## 🧪 Testing

### Manual Testing

```
# Start the server
uv run uvicorn main:app --reload

# Test health endpoint
curl http://localhost:8000/health

# Test metrics endpoint
curl http://localhost:8000/metrics
```

### Sample Test Data

```
Content: "Artificial Intelligence is transforming industries by automating tasks and providing insights. Key applications include healthcare diagnosis, financial analysis, and customer service automation. Implementation requires careful planning, data quality, and stakeholder buy-in."

Guidance: "Create a professional business presentation for executives"
```

## 🚀 Deployment

### Docker Deployment

```
# Build image
docker build -t ai-powerpoint-generator .

# Run container
docker run -d -p 8000:8000 \
  -v $(pwd)/uploads:/app/uploads \
  -v $(pwd)/generated:/app/generated \
  ai-powerpoint-generator
```

### Production Deployment

```
# Set production mode
export PRODUCTION=true

# Run with Gunicorn
gunicorn main:app -w 4 -k uvicorn.workers.UvicornWorker -b 0.0.0.0:8000
```

### Cloud Deployment

The application is ready for deployment on:

- **Heroku**: Use included `Procfile`
- **Railway**: Use included `railway.toml`
- **DigitalOcean App Platform**: Use `.do/app.yaml`
- **AWS/GCP/Azure**: Docker-compatible

## 🔍 Monitoring & Logging

### Logging Levels

- **INFO**: General application flow
- **WARNING**: Recoverable issues
- **ERROR**: Serious problems requiring attention

### Log Files

- `app_detailed.log` - All application logs
- `errors.log` - Error-specific logs with full tracebacks

### Health Checks

The `/health` endpoint provides comprehensive system status:

```
{
  "status": "healthy",
  "version": "2.2.0",
  "system_checks": {
    "directories": {...},
    "disk_space": {...}
  },
  "features": [...]
}
```

## 🛡️ Security Features

- **Input Validation**: All inputs validated and sanitized
- **File Verification**: File signatures and content verification
- **Rate Limiting**: Protection against abuse
- **Error Handling**: No sensitive information exposure
- **Session Security**: Secure session ID generation and validation

## 🎨 Supported Template Features

### Fully Preserved

- ✅ Color schemes and themes
- ✅ Font families and sizes
- ✅ Layout structures
- ✅ Background styles
- ✅ Existing images and graphics
- ✅ Brand consistency

### Enhanced Features

- ✅ Section divider slides for long content
- ✅ Summary slides for presentations with 5+ slides
- ✅ Professional speaker notes with presentation tips
- ✅ Intelligent content organization

## 🤝 Contributing

We welcome contributions! Please see our [Contributing Guidelines](CONTRIBUTING.md) for details.

### Development Setup

```
# Clone repository
git clone https://github.com/yourusername/ai-powerpoint-generator.git

# Install dependencies
uv pip install -r requirements.txt

# Install development dependencies
uv pip install pytest black flake8 mypy

# Run tests
pytest

# Format code
black .
```

### Contribution Areas

- 🐛 Bug fixes and improvements
- ✨ New features and integrations
- 📚 Documentation improvements
- 🧪 Test coverage expansion
- 🎨 UI/UX enhancements

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- **FastAPI** for the excellent web framework
- **python-pptx** for PowerPoint manipulation capabilities
- **OpenAI, Anthropic, Google** for LLM API services
- **Contributors** who help improve this project

## 📞 Support

- 📧 **Email**: support@yourproject.com
- 💬 **Issues**: [GitHub Issues](https://github.com/yourusername/ai-powerpoint-generator/issues)
- 📖 **Documentation**: [Full Documentation](https://docs.yourproject.com)
- 💡 **Feature Requests**: [GitHub Discussions](https://github.com/yourusername/ai-powerpoint-generator/discussions)

## 🗺️ Roadmap

### Version 2.3.0 (Planned)

- [ ] Multi-language support
- [ ] Batch processing capabilities
- [ ] Custom theme creation
- [ ] Advanced analytics dashboard

### Version 2.4.0 (Future)

- [ ] Real-time collaboration
- [ ] Template marketplace
- [ ] Advanced formatting options
- [ ] Export to additional formats

---

<div align="center">

**Made with ❤️ for creating better presentations**

[⭐️ Star this repo](https://github.com/yourusername/ai-powerpoint-generator) if you find it helpful!

</div>
```

This README.md provides:

1. **Professional presentation** with badges and clear sections
2. **Comprehensive feature list** highlighting key capabilities
3. **Step-by-step setup instructions** for different environments
4. **Architecture overview** and technical details
5. **API documentation** and configuration options
6. **Deployment guides** for various platforms
7. **Security and monitoring** information
8. **Contributing guidelines** and development setup
9. **Support information** and roadmap

The format follows GitHub README best practices with clear navigation, professional styling, and comprehensive coverage of all aspects of your project.

[1](https://github.com/fastapi/full-stack-fastapi-template)
[2](https://github.com/fastapi/full-stack-fastapi-template/releases)
[3](https://fastapi.tiangolo.com/project-generation/)
[4](https://github.com/TimoReusch/FastAPI-project-template)
[5](https://docs.techstartucalgary.com/projects/readme/index.html)
[6](https://github.com/alhytham-tech/fastapi-project-template)
[7](https://github.com/render-examples/fastapi)
[8](https://github.com/Salfiii/fastapi-template)
[9](https://github.com/alecordev/fastapi-template/blob/master/README.md)
[10](https://github.com/teamhide/fastapi-boilerplate)
