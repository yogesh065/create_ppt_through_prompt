# main.py (Streamlit App)
import streamlit as st
from pptx import Presentation
from groq import Groq
from mcp import ClientSession
import asyncio
import json
import os

# Load environment variables
groq_api_key = st.secrets["k"]["api_key"]
client = Groq(api_key=groq_api_key)

st.title("Groq-Powered PPT Generator")

async def groq_generate_content(query, context=None):
    """Generate PPT content using Groq's native API"""
    try:
        messages = [
            {"role": "system", "content": "You are a professional presentation creator."},
            {"role": "user", "content": f"Create 5-7 slides about: {query}"}
        ]
        
        if context:
            messages.append({"role": "assistant", "content": context})

        response = client.chat.completions.create(
            messages=messages,
            model="llama3-70b-8192",
            temperature=0.4,
            max_tokens=3000
        )
        
        return response.choices[0].message.content
        
    except Exception as e:
        st.error(f"Generation error: {str(e)}")
        return None

async def execute_mcp_tool(tool_name, params):
    """Execute MCP tools directly"""
    async with ClientSession() as session:
        return await session.execute_tool(
            tool_name,
            params,
            server_name="websearch",
            command="python",
            args=["websearch_server.py"]
        )

def create_presentation(content):
    """Create PPTX from structured content"""
    try:
        data = json.loads(content)
        prs = Presentation()
        
        # Title Slide
        title_slide = prs.slides.add_slide(prs.slide_layouts[0])
        title_slide.shapes.title.text = data['title']
        title_slide.placeholders[1].text = "Generated with Groq & MCP"
        
        # Content Slides
        for slide in data['slides']:
            content_slide = prs.slides.add_slide(prs.slide_layouts[1])
            content_slide.shapes.title.text = slide['title']
            body = content_slide.shapes.placeholders[1]
            
            for point in slide['points']:
                p = body.text_frame.add_paragraph()
                p.text = point
                p.level = 0
                
        file_path = "presentation.pptx"
        prs.save(file_path)
        return file_path
        
    except Exception as e:
        st.error(f"PPT Creation Error: {str(e)}")
        raise

async def main_async():
    st.title("AI Presentation Generator")
    
    query = st.text_input("Enter presentation topic:", 
                        "Generative AI in Healthcare 2025")
    
    if st.button("Generate Presentation"):
        with st.spinner("Researching and creating slides..."):
            try:
                # Step 1: Web Search
                search_results = await execute_mcp_tool(
                    "web_search",
                    {"query": query, "max_results": 5}
                )
                
                # Step 2: Generate Content
                generated_content = await groq_generate_content(
                    query, 
                    context=search_results
                )
                
                # Step 3: Create PPT
                if generated_content:
                    ppt_file = create_presentation(generated_content)
                    
                    with open(ppt_file, "rb") as f:
                        st.download_button(
                            "Download PPTX",
                            f,
                            file_name=ppt_file,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                    
                    st.success("Presentation generated successfully!")
                    st.json(json.loads(generated_content))
                
            except Exception as e:
                st.error(f"Error: {str(e)}")

def main():
    asyncio.run(main_async())

if __name__ == "__main__":
    main()
