import streamlit as st
from pptx import Presentation
from mcp import ClientSession
from mcp.client.stdio import stdio_client
from mcp.types import StdioServerParameters
import asyncio
import json

# Streamlit app configuration
st.title("MCP-Powered Presentation Generator")

async def generate_presentation(topic: str):
    """Main workflow using MCP protocol"""
    # Configure MCP server parameters
    server_params = StdioServerParameters(
        command="python",
        args=["websearch_server.py"]
    )

    async with stdio_client(server_params) as (read, write):
        async with ClientSession(read, write) as session:
            # Initialize connection
            await session.initialize()

            # Execute web search tool
            search_results = await session.call_tool(
                "web_search",
                {"query": topic, "max_results": 5}
            )

            # Generate PPT content
            content = await session.get_prompt(
                "ppt_generator",
                arguments={
                    "topic": topic,
                    "context": json.dumps(search_results)
                }
            )

            # Create PowerPoint file
            return create_pptx(content.messages[0].content.text)

def create_pptx(content: str):
    """Convert structured content to PowerPoint"""
    data = json.loads(content)
    prs = Presentation()
    
    # Title Slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = data['title']
    title_slide.placeholders[1].text = "Generated with MCP"
    
    # Content Slides
    for slide in data['slides']:
        content_slide = prs.slides.add_slide(prs.slide_layouts[1])
        content_slide.shapes.title.text = slide['title']
        body = content_slide.shapes.placeholders[1]
        
        for point in slide['points']:
            p = body.text_frame.add_paragraph()
            p.text = point
    
    file_path = "presentation.pptx"
    prs.save(file_path)
    return file_path

# Streamlit UI
async def main_async():
    topic = st.text_input("Enter presentation topic:", "AI in Healthcare 2025")
    
    if st.button("Generate Presentation"):
        with st.spinner("Creating your presentation..."):
            try:
                ppt_file = await generate_presentation(topic)
                
                with open(ppt_file, "rb") as f:
                    st.download_button(
                        "Download PPTX",
                        f,
                        file_name=ppt_file,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                
                st.success("Presentation generated successfully!")
                
            except Exception as e:
                st.error(f"Error: {str(e)}")

def main():
    asyncio.run(main_async())

if __name__ == "__main__":
    main()
