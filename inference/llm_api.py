from typing import List
from openai import OpenAI
import time
import logging

logger = logging.getLogger(__name__)

def get_llm_response(messages: List[str], opt, max_retries=3, timeout=120):
    """
    Robust LLM API call with retry and timeout handling.
    
    Args:
        messages: List of conversation messages
        opt: Options containing api_key, base_url, model
        max_retries: Maximum number of retry attempts (default: 3)
        timeout: Request timeout in seconds (default: 120)
    
    Returns:
        str: LLM response content
    
    Raises:
        Exception: After all retries exhausted
    """
    client = OpenAI(api_key=opt.api_key, base_url=opt.base_url, timeout=timeout)
    messages = [{"role": "user" if i % 2 == 0 else "assistant", "content": messages[i]} for i in range(len(messages))]
    
    last_error = None
    for attempt in range(max_retries):
        try:
            chat_completion = client.chat.completions.create(
                messages=messages,
                model=opt.model,
            )
            return chat_completion.choices[0].message.content
        except Exception as e:
            last_error = e
            error_type = type(e).__name__
            logger.warning(f"LLM API call failed (attempt {attempt + 1}/{max_retries}): {error_type}: {str(e)}")
            
            if attempt < max_retries - 1:
                # Exponential backoff: 2, 4, 8 seconds
                wait_time = 2 ** (attempt + 1)
                logger.info(f"Retrying in {wait_time} seconds...")
                time.sleep(wait_time)
            else:
                logger.error(f"All {max_retries} LLM API attempts failed. Last error: {error_type}: {str(e)}")
    
    # All retries exhausted
    raise Exception(f"LLM API call failed after {max_retries} attempts. Last error: {last_error}")