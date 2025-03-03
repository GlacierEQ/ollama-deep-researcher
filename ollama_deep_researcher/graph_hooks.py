from typing import Dict, Any
from fastapi import Request
from langchain.schema.runnable import RunnableConfig

async def on_research_complete(state: Dict[str, Any], config: RunnableConfig):
    """Hook that triggers when research is complete"""
    # Extract the final summary from the state
    if "final_summary" in state:
        summary = state["final_summary"]
        topic = state.get("topic", "Research Topic")
        
        # Add metadata for easy export
        state["_export_ready"] = True
        
        # Log that the research is ready for export
        print(f"Research on '{topic}' is complete and ready for export.")
    
    return state