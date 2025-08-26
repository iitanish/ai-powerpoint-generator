from fastapi import FastAPI, Request, File, UploadFile, Form, HTTPException
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse
import os
import aiofiles
import tempfile
import shutil
from datetime import datetime
import uuid
import requests
import json

# Add these imports for PowerPoint generation
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# Create directories if they don't exist
os.makedirs("uploads", exist_ok=True)
os.makedirs("generated", exist_ok=True)
os.makedirs("static", exist_ok=True)
os.makedirs("templates", exist_ok=True)

app = FastAPI(title="PowerPoint Auto-Generator", version="1.0.0")

# Mount static files and templates
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

class LLMProcessor:
    """Handle LLM API calls for different providers"""
    
    def __init__(self):
        self.providers = {
            'openai': self._call_openai,
            'anthropic': self._call_anthropic,
            'gemini': self._call_gemini
        }
    
    def detect_provider(self, api_key: str) -> str:
        """Detect which LLM provider based on API key format"""
        if api_key.startswith('sk-proj-') or api_key.startswith('sk-'):
            return 'openai'
        elif api_key.startswith('sk-ant-'):
            return 'anthropic'
        elif api_key.startswith('AIza'):
            return 'gemini'
        else:
            # Try to guess based on length and format
            if len(api_key) > 35 and 'AIza' in api_key:
                return 'gemini'
            return 'openai'  # Default fallback
    
    async def process_text(self, text_content: str, guidance: str, api_key: str) -> dict:
        """Process text through LLM to extract slide structure"""
        
        provider = self.detect_provider(api_key)
        
        prompt = self._build_prompt(text_content, guidance)
        
        try:
            result = await self.providers[provider](prompt, api_key)
            return self._parse_llm_response(result)
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"LLM API error: {str(e)}")
    
    def _build_prompt(self, text_content: str, guidance: str) -> str:
        """Build the prompt for LLM"""
        base_prompt = f"""You are a presentation expert. Convert the following text into a structured PowerPoint presentation format.

GUIDANCE: {guidance if guidance else "Create a professional, well-structured presentation"}

TEXT CONTENT:
{text_content}

Please return ONLY a JSON response with this exact structure:
{{
    "title": "Main presentation title",
    "slides": [
        {{
            "title": "Slide title",
            "content": [
                "Bullet point 1",
                "Bullet point 2",
                "Bullet point 3"
            ],
            "speaker_notes": "Optional speaker notes for this slide"
        }}
    ]
}}

Rules:
- Create 3-10 slides maximum
- Each slide should have 2-5 bullet points
- Keep titles concise and descriptive
- Make content engaging and well-structured
- Include speaker notes if helpful
- Return ONLY valid JSON, no other text"""
        return base_prompt
    
    async def _call_openai(self, prompt: str, api_key: str) -> str:
        """Call OpenAI API"""
        url = "https://api.openai.com/v1/chat/completions"
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        
        data = {
            "model": "gpt-3.5-turbo",
            "messages": [
                {"role": "user", "content": prompt}
            ],
            "max_tokens": 2000,
            "temperature": 0.3
        }
        
        response = requests.post(url, headers=headers, json=data, timeout=30)
        response.raise_for_status()
        
        result = response.json()
        return result['choices'][0]['message']['content']
    
    async def _call_anthropic(self, prompt: str, api_key: str) -> str:
        """Call Anthropic API"""
        url = "https://api.anthropic.com/v1/messages"
        headers = {
            "x-api-key": api_key,
            "Content-Type": "application/json",
            "anthropic-version": "2023-06-01"
        }
        
        data = {
            "model": "claude-3-sonnet-20240229",
            "max_tokens": 2000,
            "messages": [
                {"role": "user", "content": prompt}
            ]
        }
        
        response = requests.post(url, headers=headers, json=data, timeout=30)
        response.raise_for_status()
        
        result = response.json()
        return result['content'][0]['text']
    
    async def _call_gemini(self, prompt: str, api_key: str) -> str:
        """Call Google Gemini API with updated endpoint"""
        # Updated endpoint for Gemini
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key={api_key}"
        headers = {
            "Content-Type": "application/json"
        }
        
        data = {
            "contents": [
                {
                    "parts": [
                        {"text": prompt}
                    ]
                }
            ],
            "generationConfig": {
                "temperature": 0.3,
                "maxOutputTokens": 2000,
                "topP": 0.8,
                "topK": 10
            }
        }
        
        response = requests.post(url, headers=headers, json=data, timeout=30)
        response.raise_for_status()
        
        result = response.json()
        
        # Handle different response structures
        if 'candidates' in result and len(result['candidates']) > 0:
            candidate = result['candidates'][0]
            if 'content' in candidate and 'parts' in candidate['content']:
                return candidate['content']['parts'][0]['text']
            else:
                raise Exception("Unexpected Gemini API response structure")
        else:
            raise Exception("No candidates returned from Gemini API")
    
    def _parse_llm_response(self, response: str) -> dict:
        """Parse and validate LLM response"""
        try:
            # Clean response (remove code blocks if present)
            clean_response = response.strip()
            
            # Handle markdown code blocks
            if clean_response.startswith('```json'):
                clean_response = clean_response[7:]
                if clean_response.endswith('```'):
                    clean_response = clean_response[:-3]
                clean_response = clean_response.strip()
            elif clean_response.startswith('```'):
                clean_response = clean_response[3:]
                if clean_response.endswith('```'):
                    clean_response = clean_response[:-3]
                clean_response = clean_response.strip()
            
            # Parse JSON
            parsed = json.loads(clean_response)
            
            # Validate structure
            if 'title' not in parsed or 'slides' not in parsed:
                raise ValueError("Invalid response structure: missing title or slides")
            
            if not isinstance(parsed['slides'], list) or len(parsed['slides']) == 0:
                raise ValueError("No slides found in response")
            
            # Validate each slide
            for i, slide in enumerate(parsed['slides']):
                if 'title' not in slide or 'content' not in slide:
                    raise ValueError(f"Slide {i+1} missing title or content")
                
                if not isinstance(slide['content'], list):
                    raise ValueError(f"Slide {i+1} content must be a list")
            
            return parsed
            
        except json.JSONDecodeError as e:
            raise HTTPException(status_code=400, detail=f"Failed to parse JSON response: {str(e)}")
        except ValueError as e:
            raise HTTPException(status_code=400, detail=f"Invalid response structure: {str(e)}")
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Failed to parse LLM response: {str(e)}")

class PowerPointGenerator:
    """Generate PowerPoint presentations using template and AI-structured content"""
    
    def __init__(self):
        pass
    
    def generate_presentation(self, slide_structure: dict, template_path: str, session_id: str) -> str:
        """Generate PowerPoint presentation from structured data and template"""
        
        try:
            # Load the template presentation
            prs = Presentation(template_path)
            print(f"Loaded template with {len(prs.slide_layouts)} layouts")
            
            # Remove all existing slides except keep one as reference for styling
            slides_to_remove = list(prs.slides._sldIdLst)
            for slide_id in slides_to_remove:
                prs.slides._sldIdLst.remove(slide_id)
            
            # Get the best layouts for different slide types
            content_layout = self._get_best_layout(prs, 'content')
            title_layout = self._get_best_layout(prs, 'title') 
            
            # Create title slide
            title_slide = prs.slides.add_slide(title_layout)
            self._populate_title_slide(title_slide, slide_structure['title'])
            
            # Create content slides
            for slide_data in slide_structure['slides']:
                content_slide = prs.slides.add_slide(content_layout)
                self._populate_content_slide(content_slide, slide_data)
            
            # Save the generated presentation
            output_path = f"generated/presentation_{session_id}.pptx"
            prs.save(output_path)
            print(f"Generated presentation saved to: {output_path}")
            
            return output_path
            
        except Exception as e:
            raise Exception(f"PowerPoint generation failed: {str(e)}")
    
    def _get_best_layout(self, prs: Presentation, slide_type: str):
        """Get the best layout for the slide type"""
        layouts = prs.slide_layouts
        
        if slide_type == 'title':
            # Try to find title slide layout (usually index 0)
            for i, layout in enumerate(layouts):
                if 'title' in layout.name.lower() and len(layout.placeholders) >= 1:
                    return layout
            # Fallback to first layout
            return layouts[0]
        
        elif slide_type == 'content':
            # Try to find content layout with title and body
            for i, layout in enumerate(layouts):
                if len(layout.placeholders) >= 2:  # Has title and content
                    return layout
            # Fallback to second layout or first if only one exists
            return layouts[1] if len(layouts) > 1 else layouts[0]
        
        return layouts[0]  # Ultimate fallback
    
    def _populate_title_slide(self, slide, title_text: str):
        """Populate the title slide with the main title"""
        try:
            # Find title placeholder
            title_placeholder = None
            subtitle_placeholder = None
            
            for placeholder in slide.placeholders:
                if placeholder.placeholder_format.type == 1:  # Title
                    title_placeholder = placeholder
                elif placeholder.placeholder_format.type == 2:  # Subtitle
                    subtitle_placeholder = placeholder
            
            # Set title
            if title_placeholder:
                title_placeholder.text = title_text
                
            # Set subtitle
            if subtitle_placeholder:
                subtitle_placeholder.text = "Generated by AI PowerPoint Generator"
                
        except Exception as e:
            print(f"Warning: Could not populate title slide: {e}")
            # Try alternative method
            try:
                if slide.shapes.title:
                    slide.shapes.title.text = title_text
            except:
                pass
    
    def _populate_content_slide(self, slide, slide_data: dict):
        """Populate a content slide with title and bullet points"""
        try:
            # Find placeholders
            title_placeholder = None
            content_placeholder = None
            
            for placeholder in slide.placeholders:
                if placeholder.placeholder_format.type == 1:  # Title
                    title_placeholder = placeholder
                elif placeholder.placeholder_format.type in [2, 7, 8]:  # Content/Body
                    content_placeholder = placeholder
            
            # Set slide title
            if title_placeholder:
                title_placeholder.text = slide_data['title']
            elif slide.shapes.title:
                slide.shapes.title.text = slide_data['title']
            
            # Set slide content
            if content_placeholder:
                # Clear existing text
                content_placeholder.text = ""
                
                # Add bullet points
                text_frame = content_placeholder.text_frame
                text_frame.clear()  # Remove any existing paragraphs
                
                for i, bullet_point in enumerate(slide_data['content']):
                    if i == 0:
                        # Use the first paragraph
                        p = text_frame.paragraphs[0]
                    else:
                        # Add new paragraphs for subsequent points
                        p = text_frame.add_paragraph()
                    
                    p.text = bullet_point
                    p.level = 0  # Top-level bullet
                    
                    # Optional: Set bullet point formatting
                    try:
                        p.font.size = Pt(18)
                    except:
                        pass
            
            # Add speaker notes if available
            if 'speaker_notes' in slide_data and slide_data['speaker_notes']:
                try:
                    notes_slide = slide.notes_slide
                    notes_text_frame = notes_slide.notes_text_frame
                    notes_text_frame.text = slide_data['speaker_notes']
                except Exception as e:
                    print(f"Could not add speaker notes: {e}")
                    
        except Exception as e:
            print(f"Warning: Could not populate content slide: {e}")
    
    def get_template_info(self, template_path: str) -> dict:
        """Get information about the template"""
        try:
            prs = Presentation(template_path)
            
            info = {
                "layouts_count": len(prs.slide_layouts),
                "layouts": [],
                "slide_count": len(prs.slides)
            }
            
            for i, layout in enumerate(prs.slide_layouts):
                layout_info = {
                    "index": i,
                    "name": layout.name,
                    "placeholders": len(layout.placeholders)
                }
                info["layouts"].append(layout_info)
            
            return info
            
        except Exception as e:
            return {"error": f"Could not analyze template: {str(e)}"}

# Initialize processors
llm_processor = LLMProcessor()
ppt_generator = PowerPointGenerator()

@app.get("/", response_class=HTMLResponse)
async def get_form(request: Request):
    """Serve the main form page"""
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate")
async def generate_presentation(
    request: Request,
    text_content: str = Form(..., description="Bulk text or markdown content"),
    guidance: str = Form("", description="Optional guidance for tone/structure"),
    api_key: str = Form(..., description="LLM API key"),
    template_file: UploadFile = File(..., description="PowerPoint template file")
):
    """Handle presentation generation"""
    
    # Basic validation
    if not text_content.strip():
        raise HTTPException(status_code=400, detail="Text content cannot be empty")
    
    if not api_key.strip():
        raise HTTPException(status_code=400, detail="API key is required")
    
    if len(text_content) > 50000:  # 50k character limit
        raise HTTPException(status_code=400, detail="Text content too long (max 50,000 characters)")
    
    # File validation
    if not template_file.filename:
        raise HTTPException(status_code=400, detail="No file uploaded")
        
    if not template_file.filename.endswith(('.pptx', '.potx')):
        raise HTTPException(status_code=400, detail="Please upload a valid PowerPoint template (.pptx or .potx)")
    
    # Read and validate file size
    try:
        file_content = await template_file.read()
        file_size = len(file_content)
        
        if file_size > 10 * 1024 * 1024:  # 10MB
            raise HTTPException(status_code=400, detail="File size exceeds 10MB limit")
        
        if file_size < 1000:  # Too small to be a real PowerPoint file
            raise HTTPException(status_code=400, detail="Invalid template file - file too small")
            
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Error reading uploaded file: {str(e)}")
    
    # Generate unique session ID
    session_id = str(uuid.uuid4())[:8]
    template_path = None
    generated_ppt_path = None
    
    try:
        # Save template file temporarily
        template_path = f"uploads/template_{session_id}.pptx"
        async with aiofiles.open(template_path, 'wb') as f:
            await f.write(file_content)
        
        # Process text through LLM
        provider = llm_processor.detect_provider(api_key)
        print(f"Processing text with {provider} API...")
        
        slide_structure = await llm_processor.process_text(text_content, guidance, api_key)
        
        # Generate PowerPoint presentation
        print("Generating PowerPoint presentation...")
        generated_ppt_path = ppt_generator.generate_presentation(
            slide_structure, template_path, session_id
        )
        
        # Get template analysis
        template_info = ppt_generator.get_template_info(template_path)
        
        # Success response with download capability
        return {
            "status": "success",
            "session_id": session_id,
            "message": "Presentation generated successfully!",
            "slide_structure": slide_structure,
            "download_url": f"/download/{session_id}",
            "stats": {
                "original_text_length": len(text_content),
                "slides_generated": len(slide_structure['slides']) + 1,  # +1 for title slide
                "template_size_mb": round(file_size / (1024 * 1024), 2),
                "provider_used": provider.upper(),
                "template_info": template_info
            }
        }
        
    except HTTPException:
        # Re-raise HTTP exceptions
        raise
    except Exception as e:
        # Handle unexpected errors
        error_msg = f"Processing error: {str(e)}"
        print(f"Error in generate_presentation: {error_msg}")
        raise HTTPException(status_code=500, detail=error_msg)
        
    finally:
        # Clean up template file (but keep generated presentation for download)
        if template_path and os.path.exists(template_path):
            try:
                os.remove(template_path)
                print(f"Cleaned up temporary template: {template_path}")
            except Exception as cleanup_error:
                print(f"Error cleaning up template: {cleanup_error}")

@app.get("/download/{session_id}")
async def download_presentation(session_id: str):
    """Download generated presentation"""
    
    file_path = f"generated/presentation_{session_id}.pptx"
    
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Presentation not found or expired")
    
    return FileResponse(
        path=file_path,
        filename=f"AI_Generated_Presentation_{session_id}.pptx",
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

@app.get("/template-info/{session_id}")
async def get_template_info(session_id: str):
    """Get template analysis information"""
    
    template_path = f"uploads/template_{session_id}.pptx"
    
    if not os.path.exists(template_path):
        raise HTTPException(status_code=404, detail="Template not found")
    
    info = ppt_generator.get_template_info(template_path)
    return info

# Updated startup event handler with proper FastAPI 0.100+ syntax
@app.on_event("startup")
async def startup_event():
    """Clean up old files on startup"""
    print("Starting PowerPoint Auto-Generator...")
    cleanup_old_files()

# Updated shutdown event handler  
@app.on_event("shutdown")
async def shutdown_event():
    """Clean up on shutdown"""
    print("Shutting down PowerPoint Auto-Generator...")
    cleanup_old_files()

def cleanup_old_files():
    """Clean up old uploaded and generated files"""
    cleaned_count = 0
    for directory in ["uploads", "generated"]:
        if os.path.exists(directory):
            try:
                for filename in os.listdir(directory):
                    file_path = os.path.join(directory, filename)
                    try:
                        # Remove files older than 1 hour
                        if os.path.getctime(file_path) < (datetime.now().timestamp() - 3600):
                            os.remove(file_path)
                            print(f"Cleaned up old file: {filename}")
                            cleaned_count += 1
                    except Exception as e:
                        print(f"Error cleaning up {filename}: {e}")
            except Exception as e:
                print(f"Error accessing directory {directory}: {e}")
    
    if cleaned_count > 0:
        print(f"Cleanup complete: {cleaned_count} files removed")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
