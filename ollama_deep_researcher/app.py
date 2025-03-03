"""
Standalone API server for Word Add-in without LangGraph Studio
Useful for distribution to users who only need the Word integration
"""

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import os
import uvicorn

# Create FastAPI app
app = FastAPI(
    title="Ollama Deep Researcher - Word Add-in",
    description="Microsoft Word Add-in powered by Ollama",
    version="1.0.0"
)

# Create static directory if it doesn't exist
static_dir = os.path.join(os.path.dirname(__file__), "static")
os.makedirs(static_dir, exist_ok=True)

# Mount static files
app.mount("/static", StaticFiles(directory=static_dir), name="static")

# Configure CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Register Word API routes
from ollama_deep_researcher.api import register_word_routes
register_word_routes(app)

# Add exception handler
from ollama_deep_researcher.word_error_handlers import WordAPIError, word_api_exception_handler
app.add_exception_handler(WordAPIError, word_api_exception_handler)

# Add health check endpoint
@app.get("/health")
async def health_check():
    """API health check endpoint"""
    return {"status": "healthy", "service": "word-addin"}

def start_server(host="127.0.0.1", port=2024):
    """Start the API server"""
    uvicorn.run(app, host=host, port=port)

if __name__ == "__main__":
    start_server()
