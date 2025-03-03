"""API Registration Module for Ollama Deep Researcher."""
from fastapi import APIRouter, FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse
import os

def register_word_routes(app: FastAPI):
    """Register Word Add-in routes with the FastAPI application."""
    from ollama_deep_researcher.word_api import router as word_router
    
    # Include Word API routes
    app.include_router(word_router, prefix="/api", tags=["word"])
    
    # Get the static directory path
    static_dir = os.path.join(os.path.dirname(__file__), "static")
    
    # Mount static files if not already mounted
    if not any(mount.path == "/static" for mount in app.routes):
        app.mount("/static", StaticFiles(directory=static_dir), name="static")

    # Add route for Word Add-in interface
    @app.get("/word-addin", response_class=HTMLResponse)
    async def get_word_addin():
        """Serve the Word Add-in interface."""
        html_path = os.path.join(static_dir, "word-addin.html")
        with open(html_path, "r") as f:
            content = f.read()
        return HTMLResponse(content=content)
    
    # Add route to download the installation batch file
    @app.get("/download/install-word-addin.bat", response_class=FileResponse)
    async def download_batch():
        """Download the Word Add-in installation batch file."""
        batch_path = os.path.join(static_dir, "install-word-addin.bat")
        return FileResponse(
            path=batch_path,
            filename="install-word-addin.bat",
            media_type="application/octet-stream"
        )
    
    # Add route to serve the manifest file
    @app.get("/static/manifest.xml", response_class=HTMLResponse)
    async def get_manifest():
        """Serve the Word Add-in manifest file."""
        manifest_path = os.path.join(static_dir, "manifest.xml")
        with open(manifest_path, "r") as f:
            content = f.read()
        return HTMLResponse(content=content, media_type="application/xml")
    
    # Register document export functionality if available
    try:
        from ollama_deep_researcher.document import register_document_routes
        register_document_routes(app)
    except ImportError:
        print("Document export functionality not available")
    
    print("Word Add-in routes registered successfully")

def register_all_routes(app: FastAPI):
    """Register all API routes for Ollama Deep Researcher."""
    register_word_routes(app)
    
    # Additional routes can be registered here
    
    # Add health check endpoint
    @app.get("/api/health")
    async def health_check():
        """Health check endpoint"""
        return {"status": "healthy", "service": "ollama-deep-researcher"}
