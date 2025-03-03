from typing import Dict, List, Optional, Any
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from langgraph.api import create_app
from langchain_core.messages import HumanMessage, SystemMessage, AIMessage
import os

# Import the graph to expose
from ollama_deep_researcher.graph import get_graph
# Import the document router
from ollama_deep_researcher.document_api import router as document_router
# Add these imports
from ollama_deep_researcher.word_api import router as word_router

# Create graph instance
graph = get_graph()

# Configure FastAPI application
app = create_app(
    graph, 
    config={
        "title": "Ollama Deep Researcher", 
        "description": "A local web research assistant powered by Ollama",
        "custom_ui_script": "/custom-ui.js"  # Add this line
    }
)

# Create static directory if it doesn't exist
static_dir = os.path.join(os.path.dirname(__file__), "static")
os.makedirs(static_dir, exist_ok=True)

# Mount static files
app.mount("/static", StaticFiles(directory=static_dir), name="static")

# Add route to inject our custom UI extensions
@app.get("/custom-ui.js")
async def get_custom_ui():
    js_content = """
    // Load our custom export button
    const script = document.createElement('script');
    script.src = '/static/export_button.js';
    document.head.appendChild(script);
    """
    return HTMLResponse(content=js_content, media_type="application/javascript")

# Configure CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Include document API router
app.include_router(document_router, prefix="/api/document", tags=["document"])
# Add this line where other routers are included
app.include_router(word_router, prefix="/api", tags=["word"])

# Add route to serve the export HTML page
@app.get("/export", response_class=HTMLResponse)
async def get_export_page():
    with open(os.path.join(os.path.dirname(__file__), "../export.html"), "r") as f:
        html_content = f.read()
    return HTMLResponse(content=html_content)

# Add route to get the current state
@app.get("/api/state")
async def get_current_state():
    """Get the current graph state for export"""
    try:
        # Get the latest thread run
        threads = graph.threads
        if not threads:
            return JSONResponse(status_code=404, content={"error": "No threads found"})
            
        latest_thread = sorted(threads, key=lambda t: t.created_at, reverse=True)[0]
        state = latest_thread.get_state()
        
        # Extract relevant information
        export_data = {
            "topic": state.get("topic", "Research Topic"),
            "final_summary": state.get("final_summary", ""),
            "sources": state.get("sources", []),
            "queries": state.get("queries", [])
        }
        
        return JSONResponse(content=export_data)
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

# Add this route after the other route definitions

@app.get("/word-addin", response_class=HTMLResponse)
async def get_word_addin_page():
    """Serve the Word Add-in page"""
    with open(os.path.join(os.path.dirname(__file__), "static/word-addin.html"), "r") as f:
        html_content = f.read()
    return HTMLResponse(content=html_content)

# Add this route to serve the manifest file
@app.get("/static/manifest.xml", response_class=HTMLResponse)
async def get_manifest():
    """Serve the Word add-in manifest file"""
    with open(os.path.join(os.path.dirname(__file__), "static/manifest.xml"), "r") as f:
        manifest_content = f.read()
    return HTMLResponse(content=manifest_content, media_type="application/xml")

# Add this route for the installation guide
@app.get("/word-install", response_class=HTMLResponse)
async def get_word_install_guide():
    """Serve the Word add-in installation guide"""
    html_content = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Word Add-in Installation Guide</title>
        <style>
            body { font-family: 'Segoe UI', sans-serif; line-height: 1.6; max-width: 800px; margin: 0 auto; padding: 20px; }
            h1 { color: #0078d7; }
            .step { margin-bottom: 30px; }
            .step-number { background: #0078d7; color: white; display: inline-block; width: 30px; height: 30px; text-align: center; 
                       line-height: 30px; border-radius: 50%; margin-right: 10px; }
            code { background: #f0f0f0; padding: 2px 5px; border-radius: 3px; }
            button { padding: 8px 16px; background: #0078d7; color: white; border: none; cursor: pointer; }
            .note { background: #ffffd0; padding: 10px; border-left: 4px solid #ffd700; margin: 15px 0; }
        </style>
    </head>
    <body>
        <h1>Install Ollama Deep Researcher for Word</h1>
        
        <div class="note">
            <strong>Note:</strong> This add-in requires Ollama Deep Researcher to be running on your local machine.
        </div>
        
        <div class="step">
            <span class="step-number">1</span>
            <strong>Download the Word Add-in Manifest</strong>
            <p>
                <a href="/static/manifest.xml" download="ollama-researcher-manifest.xml">
                    <button>Download Manifest</button>
                </a>
            </p>
        </div>
        
        <div class="step">
            <span class="step-number">2</span>
            <strong>Install in Word</strong>
            <p>Open Microsoft Word and follow these steps:</p>
            <ol>
                <li>Go to the <strong>Insert</strong> tab in the ribbon</li>
                <li>Click <strong>Add-ins</strong> → <strong>My Add-ins</strong></li>
                <li>Select <strong>Upload My Add-in</strong> at the bottom</li>
                <li>Browse to the location where you saved the manifest file and select it</li>
                <li>Click <strong>Open</strong></li>
            </ol>
        </div>
        
        <div class="step">
            <span class="step-number">3</span>
            <strong>Using the Add-in</strong>
            <p>Once installed:</p>
            <ol>
                <li>Click on <strong>Ollama Deep Researcher</strong> in your add-ins list</li>
                <li>The research panel will open on the right side of Word</li>
                <li>Enter a research topic or select text and use it as context</li>
                <li>Click <strong>Research</strong> to generate content</li>
            </ol>
        </div>
        
        <div class="note">
            <strong>Troubleshooting:</strong> Make sure the Ollama Deep Researcher service is running at 
            <code>http://localhost:2024</code> before using the add-in.
        </div>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)

# Add this route for the Word integration landing page
@app.get("/word", response_class=HTMLResponse)
async def word_integration_page():
    """Main page for Word integration"""
    html_content = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>Ollama Deep Researcher - Word Integration</title>
        <style>
            body { font-family: 'Segoe UI', sans-serif; line-height: 1.6; max-width: 800px; margin: 0 auto; padding: 20px; }
            h1 { color: #0078d7; }
            .card { border: 1px solid #ddd; border-radius: 8px; padding: 20px; margin: 20px 0; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
            .button { display: inline-block; padding: 10px 20px; background: #0078d7; color: white; text-decoration: none; border-radius: 4px; }
            .status { padding: 10px; margin: 10px 0; border-radius: 4px; }
            .status.success { background: #d4edda; color: #155724; }
            .status.error { background: #f8d7da; color: #721c24; }
            .feature { margin: 10px 0; }
            .feature:before { content: "✓"; color: #0078d7; margin-right: 10px; }
        </style>
        <script>
            async function checkApiStatus() {
                const statusDiv = document.getElementById('api-status');
                try {
                    const response = await fetch('/api/word/status');
                    if (response.ok) {
                        const data = await response.json();
                        statusDiv.innerHTML = '<div class="status success">Word API is active and ready!</div>';
                    } else {
                        throw new Error('API responded with error');
                    }
                } catch (error) {
                    statusDiv.innerHTML = '<div class="status error">Word API not responding. Make sure the server is running.</div>';
                }
            }
            
            // Check status when page loads
            window.onload = checkApiStatus;
        </script>
    </head>
    <body>
        <h1>Ollama Deep Researcher for Word</h1>
        
        <div class="card">
            <h2>Features</h2>
            <div class="feature">Research topics directly within Word</div>
            <div class="feature">Generate comprehensive summaries with sources</div>
            <div class="feature">Edit and improve content with AI assistance</div>
            <div class="feature">All processing happens locally via Ollama</div>
        </div>
        
        <div class="card">
            <h2>System Status</h2>
            <div id="api-status">Checking API status...</div>
            <p>For the Word Add-in to work properly, this server must be running.</p>
        </div>
        
        <div class="card">
            <h2>Installation</h2>
            <p>Install the Word Add-in to access research capabilities directly from Microsoft Word:</p>
            <a href="/word-install" class="button">Installation Guide</a>
        </div>
        
        <div class="card">
            <h2>Try the Add-in</h2>
            <p>You can preview the add-in interface here before installing:</p>
            <a href="/word-addin" class="button">Preview Add-in</a>
        </div>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)

# Check if this file is being run directly
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=2024)