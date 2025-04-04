from mcp.server.fastmcp import FastMCP
from groq import Groq
import json
import os
import streamlit as st

mcp = FastMCP("pptgen")
api_key= st.secrets["k"]["api_key"]

client = Groq(api_key=api_key)

@mcp.tool()
async def ppt_content_generator(topic: str, context: str, style: str, slide_count: int) -> str:
    """Generate structured PPT content using LLM"""
    prompt = f"""Create a {slide_count}-slide {style} presentation about {topic}:
    
    Follow these rules:
    1. Title slide with subtitle
    2. Minimum 3 bullet points per slide
    3. Include data points from context
    4. Use professional business language
    
    Context:
    {context}
    
    Output format:
    {{
        "title": "Presentation Title",
        "subtitle": "Presentation Subtitle",
        "slides": [
            {{
                "title": "Slide 1 Title",
                "points": ["Point 1", "Point 2", "Point 3"]
            }}
        ]
    }}"""
    
    response = client.chat.completions.create(
        messages=[{"role": "user", "content": prompt}],
        model="llama-3.3-70b-specdec",
        temperature=0.4,
        max_tokens=2000
    )
    
    return response.choices[0].message.content
