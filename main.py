#!/usr/bin/env python
"""
Main entry point for Ollama Deep Researcher.
This script allows running different utilities for the project.
"""

import argparse
import sys
import os

def setup_word_addin():
    """Set up the Word Add-in integration"""
    try:
        from ollama_deep_researcher.setup_files import setup_word_addin
        setup_word_addin()
        print("Word Add-in setup complete!")
    except ImportError as e:
        print(f"Error importing setup module: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Error setting up Word Add-in: {e}")
        sys.exit(1)

def run_server(port=2024):
    """Run the LangGraph server"""
    try:
        from ollama_deep_researcher.langchain_app import app
        import uvicorn
        print(f"Starting Ollama Deep Researcher server on port {port}...")
        uvicorn.run(app, host="0.0.0.0", port=port)
    except ImportError as e:
        print(f"Error importing server module: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Error running server: {e}")
        sys.exit(1)

def main():
    parser = argparse.ArgumentParser(description="Ollama Deep Researcher management script")
    subparsers = parser.add_subparsers(dest="command", help="Command to run")
    
    # Word Add-in setup command
    word_parser = subparsers.add_parser("setup-word", help="Set up Word Add-in integration")
    
    # Server command
    server_parser = subparsers.add_parser("server", help="Run the LangGraph server")
    server_parser.add_argument("--port", type=int, default=2024, help="Port to run server on (default: 2024)")
    
    # Parse arguments
    args = parser.parse_args()
    
    # Execute requested command
    if args.command == "setup-word":
        setup_word_addin()
    elif args.command == "server":
        run_server(port=args.port)
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
