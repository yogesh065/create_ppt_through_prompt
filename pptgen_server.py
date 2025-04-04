from mcp.server.fastmcp import FastMCP
from groq import Groq
import json
import os
import streamlit as st

mcp = FastMCP("pptgen")
api_key= st.secrets["k"]["api_key"]

groq_client = Groq(api_key=api_key)

@mcp.tool()
async def generate_ppt_content(topic: str, context: str) -> str:
    """Generate structured PPT content using Groq"""
    try:
        prompt = f"""Create professional presentation slides about: {topic}
        
        Context from web search:
        {context}
        
        Output Format (JSON):
        {{
            "title": "Presentation Title",
            "subtitle": "Informative Subtitle",
            "slides": [
                {{
                    "title": "Slide Title",
                    "points": ["Bullet 1", "Bullet 2", "Bullet 3"]
                }}
            ]
        }}
        
        Use markdown-style formatting and ensure valid JSON output."""
        
        response = groq_client.chat.completions.create(
            messages=[{"role": "user", "content": prompt}],
            model="llama3-70b-8192",
            temperature=0.4,
            max_tokens=4000
        )
        return response.choices[0].message.content
        
    except Exception as e:
        return json.dumps({"error": str(e)})

if __name__ == "__main__":
    mcp.run()
