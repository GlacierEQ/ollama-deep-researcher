"""Error handling utilities for Word Add-in integration."""
import logging
import traceback
from fastapi import Request, HTTPException
from fastapi.responses import JSONResponse
from typing import Dict, Any, Callable, Awaitable
import functools
import time
import json

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("word_api")

class WordAPIError(Exception):
    """Custom exception for Word API errors."""
    def __init__(self, message: str, status_code: int = 500, details: Any = None):
        self.message = message
        self.status_code = status_code
        self.details = details
        super().__init__(message)

async def word_api_exception_handler(request: Request, exc: WordAPIError):
    """Handle Word API exceptions and return appropriate responses."""
    logger.error(f"API error: {exc.message}")
    
    error_response = {
        "error": exc.message,
        "status_code": exc.status_code
    }
    
    if exc.details:
        error_response["details"] = exc.details
        
    return JSONResponse(
        status_code=exc.status_code,
        content=error_response
    )

def with_error_handling(func: Callable) -> Callable:
    """Decorator to add error handling to Word API endpoints."""
    @functools.wraps(func)
    async def wrapper(*args, **kwargs):
        try:
            start_time = time.time()
            result = await func(*args, **kwargs)
            processing_time = time.time() - start_time
            
            # Add processing time if result is a dict and doesn't already have it
            if isinstance(result, dict) and "processing_time_seconds" not in result:
                result["processing_time_seconds"] = round(processing_time, 2)
                
            return result
        except WordAPIError as e:
            logger.error(f"Word API error in {func.__name__}: {e.message}")
            raise e
        except HTTPException as e:
            logger.error(f"HTTP error in {func.__name__}: {e.detail}")
            raise e
        except Exception as e:
            logger.error(f"Unexpected error in {func.__name__}: {str(e)}")
            logger.error(traceback.format_exc())
            raise WordAPIError(
                message=f"An unexpected error occurred: {str(e)}",
                status_code=500,
                details=traceback.format_exc()
            )
    
    return wrapper

def validate_model_response(response: str, original_content: str, instruction: str) -> Dict[str, Any]:
    """Validate the model's response for editing to ensure quality."""
    result = {"valid": True, "warnings": []}
    
    # Check for empty response
    if not response or not response.strip():
        result["valid"] = False
        result["warnings"].append("The model returned an empty response")
        return result
    
    # Check if response is identical to original
    if response.strip() == original_content.strip():
        result["warnings"].append("No changes were made to the original content")
        
    # Check if response is much shorter than original without a reason
    if len(response) < len(original_content) * 0.5 and "summarize" not in instruction.lower() and "shorten" not in instruction.lower():
        result["warnings"].append("The response is significantly shorter than the original content")
    
    # Check if response is much longer than original without a reason
    if len(response) > len(original_content) * 1.5 and "expand" not in instruction.lower() and "elaborate" not in instruction.lower():
        result["warnings"].append("The response is significantly longer than the original content")
    
    return result

def format_verification_response(analysis_data: Dict[str, Any]) -> Dict[str, Any]:
    """Format the verification analysis response for better presentation."""
    result = {}
    
    # Handle case where response isn't properly structured
    if not isinstance(analysis_data, dict):
        try:
            # Try to extract JSON from text response
            json_match = re.search(r'{.*}', analysis_data, re.DOTALL)
            if json_match:
                try:
                    analysis_data = json.loads(json_match.group(0))
                except:
                    return {
                        "error": "Unable to parse analysis results",
                        "raw_analysis": str(analysis_data)
                    }
        except:
            return {
                "error": "Received invalid analysis format",
                "raw_analysis": str(analysis_data)
            }
    
    # Copy over main analysis components with defaults
    result["grammar_issues"] = analysis_data.get("grammar_issues", [])
    result["style_issues"] = analysis_data.get("style_issues", [])
    result["structure_issues"] = analysis_data.get("structure_issues", [])
    result["improvement_suggestions"] = analysis_data.get("improvement_suggestions", [])
    result["overall_assessment"] = analysis_data.get("overall_assessment", "No assessment provided")
    
    # Calculate total issues
    total_issues = len(result["grammar_issues"]) + len(result["style_issues"]) + len(result["structure_issues"])
    result["total_issues"] = total_issues
    
    # Add severity assessment
    if total_issues == 0:
        result["severity"] = "none"
        result["severity_message"] = "No issues found"
    elif total_issues <= 3:
        result["severity"] = "low"
        result["severity_message"] = "Minor improvements suggested"
    elif total_issues <= 8:
        result["severity"] = "medium"
        result["severity_message"] = "Some issues need attention"
    else:
        result["severity"] = "high"
        result["severity_message"] = "Significant revision recommended"
    
    return result
