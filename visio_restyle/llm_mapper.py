"""
LLM integration for mapping old shapes to new masters.
"""

import json
import os
from typing import Dict, List, Any, Optional
from openai import OpenAI
from pydantic import BaseModel


class ShapeMapping(BaseModel):
    """Pydantic model for a single shape mapping."""
    old_shape_id: str
    old_master_name: str
    new_master_name: str
    reason: Optional[str] = None


class MappingResponse(BaseModel):
    """Pydantic model for the complete mapping response."""
    mappings: List[ShapeMapping]


class LLMMapper:
    """Use LLM to map shapes from source diagram to target masters."""

    def __init__(self, api_key: Optional[str] = None, model: Optional[str] = None):
        """Initialize the LLM mapper.
        
        Args:
            api_key: OpenAI API key (if None, reads from OPENAI_API_KEY env var)
            model: Model name to use (if None, reads from LLM_MODEL env var, defaults to gpt-4)
        """
        self.api_key = api_key or os.environ.get("OPENAI_API_KEY")
        if not self.api_key:
            raise ValueError("OpenAI API key must be provided or set in OPENAI_API_KEY environment variable")
        
        # Set model: parameter > LLM_MODEL env var > default
        self.model = model or os.environ.get("LLM_MODEL", "gpt-4")
        
        # Build OpenAI client configuration from environment variables
        client_kwargs = {"api_key": self.api_key}
        
        # Optional: Organization ID
        if os.environ.get("OPENAI_ORG_ID"):
            client_kwargs["organization"] = os.environ.get("OPENAI_ORG_ID")
        
        # Optional: Custom base URL (for Azure OpenAI or compatible APIs)
        if os.environ.get("OPENAI_API_BASE"):
            client_kwargs["base_url"] = os.environ.get("OPENAI_API_BASE")
        
        # Optional: API Version (for Azure OpenAI)
        if os.environ.get("OPENAI_API_VERSION"):
            client_kwargs["default_headers"] = {
                "api-version": os.environ.get("OPENAI_API_VERSION")
            }
        
        # Optional: Timeout
        if os.environ.get("OPENAI_TIMEOUT"):
            try:
                client_kwargs["timeout"] = float(os.environ.get("OPENAI_TIMEOUT"))
            except ValueError:
                pass  # Use default timeout if invalid
        
        # Optional: Max retries
        if os.environ.get("OPENAI_MAX_RETRIES"):
            try:
                client_kwargs["max_retries"] = int(os.environ.get("OPENAI_MAX_RETRIES"))
            except ValueError:
                pass  # Use default max_retries if invalid
        
        self.client = OpenAI(**client_kwargs)
        print(f"LLM Mapper initialized with model: {self.model}, base_url: {client_kwargs.get('base_url', 'default')}")
    
    def create_mapping(
        self,
        diagram_data: Dict[str, Any],
        target_masters: List[Dict[str, str]]
    ) -> Dict[str, str]:
        """Create a mapping from source shapes to target masters using LLM.
        
        Args:
            diagram_data: Extracted diagram data from VisioExtractor
            target_masters: List of available masters in the target style
            
        Returns:
            Dictionary mapping shape IDs to target master names
        """
        # Build the prompt
        prompt = self._build_prompt(diagram_data, target_masters)
        
        # Call the LLM
        response = self._call_llm(prompt)
        
        # Parse the response
        mapping = self._parse_response(response)
        
        return mapping
    
    def _build_prompt(
        self,
        diagram_data: Dict[str, Any],
        target_masters: List[Dict[str, str]]
    ) -> str:
        """Build the prompt for the LLM.
        
        Args:
            diagram_data: Extracted diagram data
            target_masters: List of target masters
            
        Returns:
            Formatted prompt string
        """
        # Extract unique shapes across all pages
        shapes_info = []
        for page in diagram_data.get("pages", []):
            for shape in page.get("shapes", []):
                if shape.get("master_name"):
                    shapes_info.append({
                        "id": shape["id"],
                        "master_name": shape["master_name"],
                        "text": shape.get("text", "")[:100]  # Limit text length
                    })
        
        # Format target masters
        masters_list = "\n".join([
            f"- {m['name']}: {m.get('description', '')}"
            for m in target_masters
        ])
        
        # Format source shapes
        shapes_list = "\n".join([
            f"- ID: {s['id']}, Master: {s['master_name']}, Text: {s['text']}"
            for s in shapes_info
        ])
        
        prompt = f"""You are a Visio diagram converter. Your task is to map shapes from a source diagram to the best matching masters in a target style.

SOURCE SHAPES:
{shapes_list}

AVAILABLE TARGET MASTERS:
{masters_list}

Please analyze each source shape and map it to the most appropriate target master based on:
1. Shape type and purpose (e.g., process, decision, data, etc.)
2. Text content and context
3. Common diagram conventions

Respond with a JSON object in this exact format:
{{
  "mappings": [
    {{
      "old_shape_id": "shape_id",
      "old_master_name": "original_master",
      "new_master_name": "target_master",
      "reason": "brief explanation"
    }}
  ]
}}

IMPORTANT: Only use master names that exist in the AVAILABLE TARGET MASTERS list above.
If no good match exists, map to the closest generic shape available.
"""
        return prompt
    
    def _call_llm(self, prompt: str) -> str:
        """Call the OpenAI API.
        
        Args:
            prompt: The prompt to send
            
        Returns:
            LLM response text
        """
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {
                        "role": "system",
                        "content": "You are a Visio diagram conversion expert. Always respond with valid JSON."
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                temperature=0.3,  # Lower temperature for more consistent mappings
                response_format={"type": "json_object"}
            )
            
            return response.choices[0].message.content
        
        except Exception as e:
            raise RuntimeError(f"Failed to call LLM API: {e}")
    
    def _parse_response(self, response: str) -> Dict[str, str]:
        """Parse the LLM response into a mapping dictionary.
        
        Args:
            response: JSON response from LLM
            
        Returns:
            Dictionary mapping shape IDs to new master names
        """
        try:
            data = json.loads(response)
            
            # Validate and extract mappings
            mapping_response = MappingResponse(**data)
            
            # Convert to simple dict: shape_id -> new_master_name
            mapping = {}
            for m in mapping_response.mappings:
                mapping[m.old_shape_id] = m.new_master_name
            
            return mapping
        
        except json.JSONDecodeError as e:
            raise ValueError(f"Failed to parse LLM response as JSON: {e}")
        except Exception as e:
            raise ValueError(f"Failed to validate LLM response: {e}")
    
    def save_mapping(self, mapping: Dict[str, str], output_path: str) -> None:
        """Save the mapping to a JSON file.
        
        Args:
            mapping: Shape ID to master name mapping
            output_path: Path to save the JSON file
        """
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(mapping, f, indent=2, ensure_ascii=False)
    
    @staticmethod
    def load_mapping(mapping_path: str) -> Dict[str, str]:
        """Load a mapping from a JSON file.
        
        Args:
            mapping_path: Path to the JSON file
            
        Returns:
            Shape ID to master name mapping
        """
        with open(mapping_path, 'r', encoding='utf-8') as f:
            return json.load(f)
