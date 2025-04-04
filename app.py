import streamlit as st
from pptx import Presentation
from mcp import ClientSession
from langchain_mcp_adapters import MultiServerMCPClient
import asyncio
import json
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
st.set_page_config(page_title="MCP PPT Maker", layout="wide")

class MCPPPTGenerator:
    def __init__(self):
        self.mcp_client = MultiServerMCPClient()
        self.presentation = Presentation()
        
    async def search_and_generate(self, query, style="professional"):
        """Main workflow: Search -> Generate -> Create PPT"""
        async with self.mcp_client:
            # Get search results
            results = await self._mcp_web_search(query)
            
            # Generate content using MCP
            content = await self._generate_content(query, results, style)
            
            # Create presentation
            self._create_slides(content)
            return "presentation.pptx"

    async def _mcp_web_search(self, query):
        """Enhanced web search with MCP"""
        result = await self.mcp_client.execute_tool(
            "web_search",
            params={"query": query, "max_results": 5}
        )
        return json.loads(result)

    async def _generate_content(self, query, results, style):
        """Generate PPT content using MCP"""
        context = "\n".join([f"{r['title']}: {r['snippet']}" for r in results])
        
        return await self.mcp_client.execute_tool(
            "ppt_content_generator",
            params={
                "topic": query,
                "context": context,
                "style": style,
                "slide_count": 7
            }
        )

    def _create_slides(self, content):
        """Create PPTX from structured content"""
        # Title slide
        title_slide = self.presentation.slides.add_slide(
            self.presentation.slide_layouts[0]
        )
        title_slide.shapes.title.text = json.loads(content)["title"]
        
        # Content slides
        for slide in json.loads(content)["slides"]:
            content_slide = self.presentation.slides.add_slide(
                self.presentation.slide_layouts[1]
            )
            content_slide.shapes.title.text = slide["title"]
            content_body = content_slide.shapes.placeholders[1]
            
            for point in slide["points"]:
                p = content_body.text_frame.add_paragraph()
                p.text = point
                
        self.presentation.save("presentation.pptx")

async def main_async():
    st.title("MCP-Powered Presentation Maker")
    
    # UI Controls
    col1, col2 = st.columns([3, 1])
    
    with col1:
        query = st.text_input("Presentation Topic:", 
                            "Latest AI Developments 2025")
        style = st.selectbox("Presentation Style:", 
                           ["Professional", "Educational", "Marketing"])
        
    with col2:
        st.write("### Settings")
        slide_count = st.slider("Slides", 5, 15, 7)
        
    if st.button("Generate Presentation"):
        generator = MCPPPTGenerator()
        
        with st.spinner("Creating your presentation..."):
            try:
                ppt_file = await generator.search_and_generate(query, style.lower())
                
                with open(ppt_file, "rb") as f:
                    st.download_button(
                        "Download PPTX",
                        f,
                        file_name=ppt_file,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                
                st.success("Presentation generated successfully!")
                st.json(json.loads(generator.content), expanded=False)
                
            except Exception as e:
                st.error(f"Generation failed: {str(e)}")

# Streamlit async wrapper
def main():
    asyncio.run(main_async())

if __name__ == "__main__":
    main()
