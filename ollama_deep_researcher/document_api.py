from fastapi import APIRouter, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
import os
import tempfile
import shutil
# Import the export function from document_export
from ollama_deep_researcher.document_export import export_to_word

router = APIRouter()

class ExportRequest(BaseModel):
    markdown: str
    topic: str = "Research Topic"

@router.post("/export-docx")
async def export_to_docx(request: ExportRequest, background_tasks: BackgroundTasks):
    """Export research summary to Word document"""
    # Create temporary file
    temp_dir = tempfile.mkdtemp()
    output_path = os.path.join(temp_dir, f"research-{request.topic[:20].replace(' ', '-')}.docx")
    try:
        # Export to Word
        word_path = export_to_word(
            markdown_content=request.markdown,
            topic=request.topic,
            output_path=output_path
        )
        
        # Schedule cleanup after response is sent
        def cleanup():
            try:
                shutil.rmtree(temp_dir)
            except Exception as e:
                pass
        
        background_tasks.add_task(cleanup)
        
        # Return file for download
        return FileResponse(
            path=word_path, 
            filename=os.path.basename(word_path),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        # Clean up if there's an error
        shutil.rmtree(temp_dir)
        return JSONResponse(
            status_code=500,
            content={"error": str(e)}
        )