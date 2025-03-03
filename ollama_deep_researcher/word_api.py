from fastapi import APIRouter, BackgroundTasks, HTTPException, Body, Request
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field, validator
import uuid
import asyncio
from typing import Dict, Any, Optional, List
import json
import logging
from langchain_core.messages import HumanMessage, SystemMessage

# Import graph components
from ollama_deep_researcher.graph import get_graph

# Import custom error handlers
from ollama_deep_researcher.word_error_handlers import (
    WordAPIError, with_error_handling, validate_model_response,
    format_verification_response
)

router = APIRouter()
logger = logging.getLogger(__name__)

# Store active research threads
active_threads = {}

class ResearchRequest(BaseModel):
    topic: str
    model: str = "llama3.2"
    max_loops: int = 3

class ImproveRequest(BaseModel):
    text: str
    model: str = "llama3.2"

class ExpandRequest(BaseModel):
    text: str
    model: str = "llama3.2"

class WordResearchRequest(BaseModel):
    topic: str = Field(..., description="Research topic")
    context: Optional[str] = Field(None, description="Optional research context")
    
    @validator('topic')
    def topic_not_empty(cls, v):
        if not v or not v.strip():
            raise ValueError('Topic cannot be empty')
        return v.strip()

class WordEditRequest(BaseModel):
    content: str = Field(..., description="Text content to edit")
    instruction: str = Field(..., description="Editing instruction")
    
    @validator('content')
    def content_not_empty(cls, v):
        if not v or not v.strip():
            raise ValueError('Content cannot be empty')
        return v.strip()
    
    @validator('instruction')
    def instruction_not_empty(cls, v):
        if not v or not v.strip():
            raise ValueError('Instruction cannot be empty')
        return v.strip()

class VerifyRequest(BaseModel):
    content: str = Field(..., description="Content to verify")

@router.post("/research")
async def start_research(request: ResearchRequest):
    """Start a research process and return thread ID"""
    try:
        # Generate unique thread ID
        thread_id = str(uuid.uuid4())
        
        # Initialize graph with request parameters
        graph = get_graph(
            ollama_model=request.model,
            max_loops=request.max_loops
        )
        
        # Start research process in background
        active_threads[thread_id] = {
            "status": "running",
            "topic": request.topic,
            "graph": graph,
            "result": None,
            "error": None
        }
        
        # Start in background
        asyncio.create_task(
            run_research(thread_id, request.topic, graph)
        )
        
        return {"thread_id": thread_id, "status": "started"}
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to start research: {str(e)}")

async def run_research(thread_id: str, topic: str, graph):
    """Run the research process asynchronously"""
    try:
        # Execute graph with topic
        config = {"recursion_limit": 100}  # Prevent infinite loops
        result = await graph.aflow({
            "topic": topic
        }, config=config)
        
        # Store results
        active_threads[thread_id]["status"] = "completed"
        active_threads[thread_id]["result"] = result
    
    except Exception as e:
        # Store error
        active_threads[thread_id]["status"] = "failed"
        active_threads[thread_id]["error"] = str(e)
        print(f"Research failed: {str(e)}")

@router.get("/research/{thread_id}/status")
async def get_research_status(thread_id: str):
    """Get status of research thread"""
    if thread_id not in active_threads:
        raise HTTPException(status_code=404, detail="Research thread not found")
    
    thread_data = active_threads[thread_id]
    return {
        "status": thread_data["status"],
        "topic": thread_data["topic"]
    }

@router.get("/research/{thread_id}/results")
async def get_research_results(thread_id: str):
    """Get results of completed research"""
    if thread_id not in active_threads:
        raise HTTPException(status_code=404, detail="Research thread not found")
    
    thread_data = active_threads[thread_id]
    
    if thread_data["status"] == "running":
        return {"status": "running", "message": "Research still in progress"}
    
    if thread_data["status"] == "failed":
        raise HTTPException(status_code=500, detail=f"Research failed: {thread_data['error']}")
    
    # Return the final results
    return {
        "topic": thread_data["topic"],
        "final_summary": thread_data["result"].get("final_summary", "No summary generated"),
        "sources": thread_data["result"].get("sources", [])
    }

@router.post("/improve")
async def improve_text(request: ImproveRequest):
    """Improve the provided text"""
    try:
        from langchain_community.chat_models import ChatOllama
        from langchain.prompts import PromptTemplate
        
        # Set up Ollama client
        ollama_client = ChatOllama(
            model=request.model,
            base_url="http://localhost:11434"
        )
        
        # Set up prompt
        template = """
        You are an expert writing assistant. Improve the following text to make it more clear,
        concise, and professional while maintaining its original meaning and key information.
        
        Original text:
        
        ```
        {text}
        ```
        
        Please provide the improved version only, without any additional explanation or commentary.
        """
        
        prompt = PromptTemplate.from_template(template)
        
        # Run inference
        chain = prompt | ollama_client
        result = chain.invoke({"text": request.text})
        
        # Extract content
        improved_text = result.content
        
        return {"improved_text": improved_text}
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to improve text: {str(e)}")

@router.post("/expand")
async def expand_text(request: ExpandRequest):
    """Expand the provided text with more details and information"""
    try:
        from langchain_community.chat_models import ChatOllama
        from langchain.prompts import PromptTemplate
        
        # Set up Ollama client
        ollama_client = ChatOllama(
            model=request.model,
            base_url="http://localhost:11434"
        )
        
        # Set up prompt
        template = """
        You are an expert writing assistant. Expand the following text with more details,
        examples, and supporting information. Maintain the original tone and style while
        making the content more comprehensive and informative.
        
        Original text:
        
        ```
        {text}
        ```
        
        Please provide the expanded version only, without any additional explanation or commentary.
        """
        
        prompt = PromptTemplate.from_template(template)
        
        # Run inference
        chain = prompt | ollama_client
        result = chain.invoke({"text": request.text})
        
        # Extract content
        expanded_text = result.content
        
        return {"expanded_text": expanded_text}
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to expand text: {str(e)}")

@router.post("/word/research")
@with_error_handling
async def research_for_word(request: WordResearchRequest):
    """Conduct research from Word and return formatted results"""
    try:
        logger.info(f"Starting Word research on topic: {request.topic}")
        
        # Get graph instance
        graph = get_graph()
        
        # Initialize research config
        config = {
            "configurable": {
                "topic": request.topic
            }
        }
        
        # Add existing content as context if provided
        if request.context:
            config["configurable"]["context"] = request.context
        
        # Run the research graph
        logger.info(f"Creating research thread with config: {config}")
        thread = graph.create_thread(config)
        
        logger.info(f"Running research on topic: {request.topic}")
        result = thread.run(config)
        
        # Extract summary and sources
        final_summary = result.get("final_summary", "No research results available.")
        sources = result.get("sources", [])
        
        logger.info(f"Research complete, summary length: {len(final_summary)}")
        
        # Format response for Word
        response = {
            "status": "success",
            "summary": final_summary,
            "sources": sources,
            "queries_used": result.get("queries", [])
        }
        
        return response
    except Exception as e:
        logger.error(f"Research failed: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Research failed: {str(e)}")

@router.post("/word/edit")
@with_error_handling
async def edit_for_word(request: WordEditRequest):
    """Edit content based on instructions"""
    try:
        logger.info(f"Starting content editing with instructions: {request.instruction[:50]}...")
        
        # Import the model here to avoid circular imports
        from ollama_deep_researcher.model import get_llm
        
        llm = get_llm()
        
        system_message = """You are a helpful writing assistant. 
        Edit the provided content according to the instruction.
        Maintain the original meaning unless explicitly told to change it.
        Return only the edited content without explanations or additional text."""
        
        messages = [
            SystemMessage(content=system_message),
            HumanMessage(content=f"""
            CONTENT: {request.content}
            
            INSTRUCTION: {request.instruction}
            
            Edit the content according to the instruction. Return only the edited content.
            """)
        ]
        
        response = llm.invoke(messages)
        edited_content = response.content
        
        return {
            "status": "success",
            "edited_content": edited_content
        }
    except Exception as e:
        logger.error(f"Editing failed: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Failed to edit content: {str(e)}")

@router.get("/word/status")
@with_error_handling
async def check_status():
    """Check if the Word API is active and ready"""
    try:
        # Import the model here to avoid circular imports
        from ollama_deep_researcher.model import get_llm
        
        # Test the model with a simple query
        start_time = time.time()
        llm = get_llm()
        test_response = llm.invoke("Hello")
        response_time = time.time() - start_time
        
        return {
            "status": "active", 
            "message": "Word API is ready", 
            "model_responsive": True,
            "response_time_seconds": round(response_time, 2)
        }
    except Exception as e:
        logger.error(f"Status check failed: {str(e)}")
        return {
            "status": "degraded",
            "message": f"API is running but model may not be available: {str(e)}",
            "model_responsive": False
        }

@router.post("/word/edit")
@with_error_handling
async def edit_content(request: WordEditRequest):
    """Edit content based on instruction"""
    logger.info(f"Processing edit request with instruction: {request.instruction[:50]}...")
    
    # Import the model here to avoid circular imports
    from ollama_deep_researcher.model import get_llm
    
    # Split long content into manageable chunks if necessary
    content = request.content
    instruction = request.instruction
    
    # For very long content, let's use chunking approach
    if len(content) > 8000:
        logger.info(f"Content length exceeds 8000 chars ({len(content)}), using chunked processing")
        return await process_long_edit(content, instruction)
    
    # Normal processing for shorter content
    llm = get_llm()
    
    system_message = """You are a helpful writing assistant. 
    Edit the provided content according to the instruction.
    Maintain the original meaning unless explicitly told to change it.
    Return only the edited content without explanations or additional text.
    Ensure your edits address the instruction completely and effectively."""
    
    messages = [
        SystemMessage(content=system_message),
        HumanMessage(content=f"""
        CONTENT: {content}
        
        INSTRUCTION: {instruction}
        
        Edit the content according to the instruction. Return only the edited content.
        """)
    ]
    
    start_time = time.time()
    response = llm.invoke(messages)
    edited_content = response.content
    processing_time = time.time() - start_time
    
    # Validate the edited content
    validation = validate_model_response(edited_content, content, instruction)
    result = {
        "edited_content": edited_content,
        "processing_time_seconds": round(processing_time, 2)
    }
    
    if validation["warnings"]:
        result["warnings"] = validation["warnings"]
    
    logger.info(f"Edit completed in {processing_time:.2f} seconds")
    return result

async def process_long_edit(content: str, instruction: str):
    """Process very long content by splitting into chunks."""
    from ollama_deep_researcher.model import get_llm
    llm = get_llm()
    
    # Split content into paragraphs
    paragraphs = content.split("\n\n")
    
    # Group paragraphs into chunks of reasonable size
    chunks = []
    current_chunk = []
    current_length = 0
    
    for para in paragraphs:
        if current_length + len(para) > 6000:  # Keep chunks under 6000 chars
            chunks.append("\n\n".join(current_chunk))
            current_chunk = [para]
            current_length = len(para)
        else:
            current_chunk.append(para)
            current_length += len(para)
            
    # Don't forget the last chunk
    if current_chunk:
        chunks.append("\n\n".join(current_chunk))
    
    logger.info(f"Split content into {len(chunks)} chunks for processing")
    
    # Process each chunk
    edited_chunks = []
    for i, chunk in enumerate(chunks):
        logger.info(f"Processing chunk {i+1}/{len(chunks)}")
        
        system_message = f"""You are a helpful writing assistant.
        Edit the provided content according to the instruction.
        This is chunk {i+1} of {len(chunks)} of a larger document.
        Maintain the original meaning unless explicitly told to change it.
        Return only the edited content without explanations or additional text."""
        
        messages = [
            SystemMessage(content=system_message),
            HumanMessage(content=f"""
            CONTENT: {chunk}
            
            INSTRUCTION: {instruction}
            
            Edit the content according to the instruction. Return only the edited content.
            """)
        ]
        
        response = llm.invoke(messages)
        edited_chunks.append(response.content)
    
    # Combine the edited chunks
    edited_content = "\n\n".join(edited_chunks)
    
    return {
        "edited_content": edited_content,
        "info": f"Content was processed in {len(chunks)} chunks due to its length."
    }

@router.post("/word/research")
@with_error_handling
async def research_topic(request: WordResearchRequest):
    """Research a topic and return markdown content"""
    logger.info(f"Starting research on topic: {request.topic}")
    start_time = time.time()
    
    # Import here to avoid circular imports
    from ollama_deep_researcher.graph import get_graph
    
    # Create new thread and run the graph
    graph = get_graph()
    thread = graph.new_thread()
    
    config = {
        "configurable": {
            "topic": request.topic
        }
    }
    
    if request.context:
        config["configurable"]["context"] = request.context
        
    result = thread.run(config)
    
    # Extract the final summary from the result
    if "final_summary" in result:
        summary = result["final_summary"]
        sources = result.get("sources", [])
        
        # Format sources as markdown links if available
        source_links = []
        for i, source in enumerate(sources[:10]):  # Limit to top 10 sources
            title = source.get("title", f"Source {i+1}")
            url = source.get("url", "#")
            source_links.append(f"- [{title}]({url})")
            
        # Add sources to the summary if any exist
        if source_links:
            summary += "\n\n## Sources\n\n" + "\n".join(source_links)
        
        processing_time = time.time() - start_time
        logger.info(f"Research completed in {processing_time:.2f} seconds")
        
        return {
            "markdown_content": summary,
            "processing_time_seconds": round(processing_time, 2),
            "source_count": len(sources)
        }
    else:
        logger.warning("Research completed but no summary was generated")
        raise WordAPIError(
            message="Research completed but no summary was generated",
            status_code=404
        )

@router.post("/word/verify")
@with_error_handling
async def verify_content(request: VerifyRequest):
    """Verify content for issues like grammar, spelling and style"""
    # Import the model here to avoid circular imports
    from ollama_deep_researcher.model import get_llm
    
    content = request.content
    if len(content) > 8000:
        # Truncate for analysis
        content = content[:8000] + "...[content truncated for analysis]"
        
    llm = get_llm()
    
    system_message = """You are a professional editor and writing assistant.
    Analyze the provided text for issues with:
    1. Grammar and spelling
    2. Style and clarity
    3. Structure and flow
    
    Return your analysis in JSON format with these keys:
    {
      "grammar_issues": [list of specific grammar or spelling issues found],
      "style_issues": [list of style or clarity issues found],
      "structure_issues": [list of structure or flow issues found],
      "improvement_suggestions": [list of specific suggestions to improve the writing],
      "overall_assessment": "A brief overall assessment of the writing quality"
    }
    
    If no issues are found in a category, return an empty list.
    Be constructive and helpful in your analysis."""
    
    messages = [
        SystemMessage(content=system_message),
        HumanMessage(content=f"Please analyze this text:\n\n{content}")
    ]
    
    response = llm.invoke(messages)
    
    # Extract JSON from response
    try:
        # Try to extract JSON if it's formatted within backticks
        if "```json" in response.content and "```" in response.content.split("```json", 1)[1]:
            json_str = response.content.split("```json", 1)[1].split("```", 1)[0]
            analysis = json.loads(json_str)
        else:
            # Otherwise assume the whole response is JSON
            analysis = json.loads(response.content)
        
        # Format and enhance the verification response
        formatted_analysis = format_verification_response(analysis)
        return formatted_analysis
        
    except json.JSONDecodeError:
        # If JSON parsing fails, return the raw response
        logger.warning("Failed to parse JSON from verification response")
        return {
            "raw_analysis": response.content,
            "error": "Failed to parse structured analysis"
        }