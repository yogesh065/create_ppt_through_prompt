# main.py (Streamlit App)
import streamlit as st
from pptx import Presentation
from langchain_mcp_adapters.client import MultiServerMCPClient
from langgraph.prebuilt import create_react_agent
from groq import Groq
import asyncio
import json
import os
import sys

# Load environment variables
groq_api_key = st.secrets["k"]["api_key"]
os.environ["MCP_LOG_LEVEL"] = "INFO"

class MCPPresentationGenerator:
    def __init__(self):
        self.client = MultiServerMCPClient()
        self.agent = None
        self.groq_client = Groq(api_key=groq_api_key)
        
    async def initialize_services(self):
        """Initialize MCP connections"""
        try:
            # Connect to web search service
            await self.client.connect_to_server(
                "websearch",
                command=sys.executable,
                args=["websearch_server.py"],
                transport="stdio"
            )
            
            # Connect to PPT generation service
            await self.client.connect_to_server(
                "pptgen",
                command=sys.executable,
                args=["pptgen_server.py"],
                transport="stdio"
            )
            
            # Create agent with tools
            self.agent = create_react_agent(self.groq_client, self.client.get_tools())
            
        except Exception as e:
            st.error(f"Service initialization failed: {str(e)}")
            raise

    async def generate_presentation(self, query):
        """Generate presentation using Groq LLM"""
        try:
            response = await self.agent.ainvoke({
                "messages": [{
                    "role": "user",
                    "content": f"Create professional slides about: {query}"
                }]
            })
            return self._create_pptx(response['messages'][-1]['content'])
            
        except Exception as e:
            st.error(f"Generation failed: {str(e)}")
            raise

    def _create_pptx(self, content):
        """Convert structured content to PowerPoint"""
        try:
            data = json.loads(content)
            prs = Presentation()
            
            # Title Slide
            title_slide = prs.slides.add_slide(prs.slide_layouts[0])
            title_slide.shapes.title.text = data['title']
            title_slide.placeholders[1].text = data.get('subtitle', 'Generated with Groq')
            
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
    st.title("Groq-Powered Presentation Generator")
    generator = MCPPresentationGenerator()
    
    with st.spinner("Initializing services..."):
        await generator.initialize_services()
    
    query = st.text_input("Enter presentation topic:", 
                        "Large Language Model Applications 2025")
    
    if st.button("Generate Presentation"):
        with st.spinner("Creating your presentation..."):
            try:
                ppt_file = await generator.generate_presentation(query)
                
                with open(ppt_file, "rb") as f:
                    st.download_button(
                        "Download PPTX",
                        f,
                        file_name=ppt_file,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                
                st.success("Presentation generated successfully!")
                st.json(json.loads(open(ppt_file.replace(".pptx", ".json")).read()))
                
            except Exception as e:
                st.error(f"Final Error: {str(e)}")

def main():
    asyncio.run(main_async())

if __name__ == "__main__":
    main()
