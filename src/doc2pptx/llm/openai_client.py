"""
OpenAI client for doc2pptx.

This module provides a client for interacting with the OpenAI API to enhance
presentation generation with AI capabilities.
"""

import logging
from typing import Any, Dict, List, Optional, Union

import instructor
from openai import OpenAI
from pydantic import BaseModel

from doc2pptx.core.settings import settings

# Configure logging
logger = logging.getLogger(__name__)

# Use instructor to enhance the OpenAI client with Pydantic model validation
client = instructor.patch(OpenAI(api_key=settings.openai_api_key))


class OpenAIClient:
    """
    Client for interacting with the OpenAI API.
    
    This class provides methods for making requests to the OpenAI API with
    appropriate error handling and response formatting.
    """
    
    def __init__(self, 
                 model: Optional[str] = None,
                 temperature: Optional[float] = None):
        """
        Initialize an OpenAI client.
        
        Args:
            model: The OpenAI model to use (default from settings)
            temperature: The temperature setting (default from settings)
        """
        self.model = model or settings.openai_model
        self.temperature = temperature or settings.openai_temperature
    
    def chat_completion(self, 
                        messages: List[Dict[str, str]], 
                        response_model: Optional[BaseModel] = None) -> Union[Dict[str, Any], BaseModel]:
        """
        Make a chat completion request to the OpenAI API.
        
        Args:
            messages: List of message dictionaries to send to the API
            response_model: Optional Pydantic model for response validation
            
        Returns:
            Model instance or raw response dict
            
        Raises:
            OpenAIError: If the API request fails
        """
        try:
            if response_model:
                return client.chat.completions.create(
                    model=self.model,
                    temperature=self.temperature,
                    response_model=response_model,
                    messages=messages
                )
            else:
                response = client.chat.completions.create(
                    model=self.model,
                    temperature=self.temperature,
                    messages=messages
                )
                return response.choices[0].message.content
        except Exception as e:
            logger.error(f"Error making OpenAI API request: {e}")
            raise