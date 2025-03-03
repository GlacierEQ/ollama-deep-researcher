"""Ollama Deep Researcher package."""
import os

# Initialize the static directory for Word Add-in files
static_dir = os.path.join(os.path.dirname(__file__), "static")
os.makedirs(static_dir, exist_ok=True)

# Check if icon files exist, if not, create them
icon_files = ["icon-32.png", "icon-80.png"]
missing_icons = any(not os.path.exists(os.path.join(static_dir, icon)) for icon in icon_files)

if missing_icons:
    try:
        from ollama_deep_researcher.static.create_icons import create_icons
        print("Creating icon files for Word Add-in...")
    except ImportError:
        print("Warning: Could not import icon creation script")
