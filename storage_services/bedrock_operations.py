import json
from core.config import bedrock_agent
from typing import Optional, Generator


def invoke_model_with_response_stream(
    model_id: str,
    messages: list[dict],
    system_prompt: Optional[str] = None,
    max_tokens: int = 2048,
    temperature: float = 0.7,
) -> Generator[str, None, None]:
    """Invoke a Bedrock model with streaming response using invoke_model_with_response_stream.
    
    Args:
        model_id: The model ID (e.g., "us.anthropic.claude-haiku-4-5-20251001-v1:0")
        messages: List of message dicts with 'role' and 'content' keys
        system_prompt: Optional system prompt to set model behavior
        max_tokens: Maximum tokens in response
        temperature: Sampling temperature
    
    Yields:
        Text chunks from the model response
    """
    try:
        body = {
            "anthropic_version": "bedrock-2023-05-31",
            "max_tokens": max_tokens,
            "messages": messages,
            "temperature": temperature,
        }
        
        if system_prompt:
            body["system"] = system_prompt
        
        response = bedrock_agent.invoke_model_with_response_stream(
            modelId=model_id,
            body=json.dumps(body),
        )
        
        for event in response["body"]:
            chunk = json.loads(event["chunk"]["bytes"])
            if chunk.get("type") == "content_block_delta":
                delta = chunk.get("delta", {})
                if "text" in delta:
                    yield delta["text"]
    except Exception as e:
        print(f"Error invoking Bedrock model with response stream: {e}")
        raise


def invoke_model_stream(
    model_id: str,
    messages: list[dict],
    system_prompt: Optional[str] = None,
    max_tokens: int = 2048,
    temperature: float = 0.7,
) -> Generator[str, None, None]:
    """Invoke a Bedrock model with streaming response using converse_stream.
    
    Args:
        model_id: The model ID (e.g., "us.anthropic.claude-haiku-4-5-20251001-v1:0")
        messages: List of message dicts with 'role' and 'content' keys
        system_prompt: Optional system prompt to set model behavior
        max_tokens: Maximum tokens in response
        temperature: Sampling temperature
    
    Yields:
        Text chunks from the model response
    """
    try:
        kwargs = {
            "modelId": model_id,
            "messages": messages,
            "inferenceConfig": {
                "maxTokens": max_tokens,
                "temperature": temperature,
            },
        }
        
        if system_prompt:
            kwargs["system"] = [{"text": system_prompt}]
        
        response = bedrock_agent.converse_stream(**kwargs)
        
        for event in response["stream"]:
            if "contentBlockDelta" in event:
                delta = event["contentBlockDelta"]["delta"]
                if "text" in delta:
                    yield delta["text"]
    except Exception as e:
        print(f"Error invoking Bedrock model stream: {e}")
        raise


def invoke_model(
    model_id: str,
    messages: list[dict],
    system_prompt: Optional[str] = None,
    max_tokens: int = 2048,
    temperature: float = 0.7,
    stream: bool = False,
) -> dict | Generator[str, None, None]:
    """Invoke a Bedrock model with the provided messages.
    
    Args:
        model_id: The model ID (e.g., "us.anthropic.claude-haiku-4-5-20251001-v1:0")
        messages: List of message dicts with 'role' and 'content' keys
        system_prompt: Optional system prompt to set model behavior
        max_tokens: Maximum tokens in response
        temperature: Sampling temperature
        stream: If True, returns a generator; if False, returns full response dict
    
    Returns:
        Response dict if stream=False, or Generator[str] if stream=True
    """
    try:
        if stream:
            return invoke_model_stream(
                model_id,
                messages,
                system_prompt=system_prompt,
                max_tokens=max_tokens,
                temperature=temperature,
            )
        
        kwargs = {
            "modelId": model_id,
            "messages": messages,
            "inferenceConfig": {
                "maxTokens": max_tokens,
                "temperature": temperature,
            },
        }
        
        if system_prompt:
            kwargs["system"] = [{"text": system_prompt}]
        
        response = bedrock_agent.converse(**kwargs)
        return response
    except Exception as e:
        print(f"Error invoking Bedrock model: {e}")
        raise


def extract_content_from_response(response: dict) -> Optional[str]:
    """Extract text content from a Bedrock converse response.
    
    Args:
        response: Response dict from model invocation
    
    Returns:
        Extracted text content or None
    """
    try:
        if "output" in response and "message" in response["output"]:
            message = response["output"]["message"]
            if "content" in message and len(message["content"]) > 0:
                content = message["content"][0]
                if "text" in content:
                    return content["text"]
        return None
    except Exception as e:
        print(f"Error extracting content from response: {e}")
        return None


def format_messages_for_model(
    user_message: str,
    conversation_history: Optional[list[dict]] = None,
) -> list[dict]:
    """Format messages for Bedrock model consumption.
    
    Args:
        user_message: The current user message
        conversation_history: Optional list of previous messages
    
    Returns:
        Formatted messages list with role and content
    """
    messages = []
    
    if conversation_history:
        messages.extend(conversation_history)
    
    messages.append({"role": "user", "content": [{"text": user_message}]})
    return messages


def invoke_agent(
    agent_id: str,
    agent_alias_id: str,
    session_id: str,
    input_text: str,
) -> dict:
    """Invoke a Bedrock agent.
    
    Args:
        agent_id: The agent ID
        agent_alias_id: The agent alias ID
        session_id: Session ID for maintaining conversation context
        input_text: User input text
    
    Returns:
        Response dict containing the agent's output
    """
    try:
        response = bedrock_agent.invoke_agent(
            agentId=agent_id,
            agentAliasId=agent_alias_id,
            sessionId=session_id,
            inputText=input_text,
        )
        return response
    except Exception as e:
        print(f"Error invoking Bedrock agent: {e}")
        raise
