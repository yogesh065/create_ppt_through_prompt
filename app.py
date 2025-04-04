import streamlit as st
from pptx import Presentation
from pptx.util import Inches
from groq import Groq
from dotenv import load_dotenv
import os
import requests
from bs4 import BeautifulSoup
import io
import reveal_slides as rs

# Set page configuration
st.set_page_config(
    page_title="Professional Presentation Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better UI
st.markdown("""
<style>
    .reportview-container {
        background: #f7f7f9;
    }
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    h1, h2, h3 {
        color: #1E3A8A;
    }
    .stButton>button {
        background-color: #1E3A8A;
        color: white;
        border-radius: 5px;
        padding: 0.5rem 1rem;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #2563EB;
        border-color: #2563EB;
    }
</style>
""", unsafe_allow_html=True)

# Load environment variables (Groq API Key)
load_dotenv()
try:
    groq_api_key = st.secrets["k"]["api_key"]
except:
    groq_api_key = os.getenv("GROQ_API_KEY")
    
if not groq_api_key:
    st.error("Please set your GROQ_API_KEY in a .env file or in Streamlit secrets.")
    st.stop()

# Initialize Groq client
client = Groq(api_key=groq_api_key)

# Initialize session state for storing generated content
if 'generated_content' not in st.session_state:
    st.session_state.generated_content = None
if 'slide_markdown' not in st.session_state:
    st.session_state.slide_markdown = None
if 'presentation_file' not in st.session_state:
    st.session_state.presentation_file = None

# Function to convert presentation to markdown for reveal.js
def pptx_to_markdown(slide_content):
    markdown = "---\ntheme: white\n---\n\n"
    
    slides = slide_content.split("\n\n")
    for slide_text in slides:
        lines = slide_text.split("\n")
        if len(lines) < 2:
            continue
            
        slide_title = lines[0].replace("Title: ", "").strip()
        bullet_points = [point.strip() for point in lines[1:] if point.strip()]
        
        markdown += f"## {slide_title}\n\n"
        for point in bullet_points:
            markdown += f"- {point}\n"
        
        markdown += "\n---\n\n"
    
    return markdown

# Function to generate slide content using Groq with improved prompts
def groq_generate_content(topic, context):
    try:
        prompt = f"""Create a professional presentation with 5 slides about "{topic}".
The presentation should follow these guidelines:
1. Start with a compelling title slide
2. Include an agenda or overview slide
3. Each content slide should have a clear, concise title
4. Bullet points should be specific, actionable, and data-driven
5. Use the principle of "one idea per slide"
6. Include a strong concluding slide

For each slide, provide:
- A clear title (prefixed with "Title: ")
- 3-5 concise bullet points that elaborate on the title
- Each bullet point should provide valuable information, not just concepts

Additional context: {context}"""
        
        response = client.chat.completions.create(
            messages=[
                {
                    "role": "system",
                    "content": "You are an expert presentation designer who creates well-structured, engaging, and professional slide content."
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            model="llama-3.3-70b-specdec",
            temperature=0.7,
            max_tokens=4024,
        )
        generated_text = response.choices[0].message.content
        return generated_text
    except Exception as e:
        st.error(f"Error generating content with Groq: {e}")
        return None

# Function to create a PowerPoint presentation with enhanced styling
def create_presentation(topic, slide_content):
    prs = Presentation()
    
    # Title Slide
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    
    title.text = topic
    subtitle.text = "Professional Presentation"
    
    # Create content slides
    slides = slide_content.split("\n\n")
    for slide_text in slides:
        lines = slide_text.split("\n")
        if len(lines) < 2:
            continue
        
        slide_title = lines[0].replace("Title: ", "").strip()
        bullet_points = [point.strip() for point in lines[1:] if point.strip()]
        
        # Add content slide
        content_slide_layout = prs.slide_layouts[1]  # Layout with title and content
        slide = prs.slides.add_slide(content_slide_layout)
        
        # Set title
        title = slide.shapes.title
        title.text = slide_title
        
        # Add bullet points
        body = slide.placeholders[1]
        tf = body.text_frame
        tf.text = ""  # Clear any default text
        
        for point in bullet_points:
            p = tf.add_paragraph()
            p.text = point
        
    # Save the presentation to a BytesIO object
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    
    return pptx_io

# Main application UI with tabs
st.title("Professional Presentation Generator")
st.markdown("### Create beautiful, engaging presentations with AI assistance")

# Main content with tabs for better organization
tab1, tab2 = st.tabs(["Create Presentation", "Preview Slides"])

with tab1:
    # Topic input
    topic = st.text_input("Presentation Topic", "", help="Enter the main topic of your presentation")
    
    # Context input
    context = st.text_area(
        "Additional Context (optional)",
        "",
        height=150,
        help="Provide any additional information or specific points to include"
    )
    
    # Generate slides button
    if st.button("Generate Presentation", use_container_width=True):
        if topic:
            with st.spinner("Generating professional slides..."):
                # Generate content
                generated_content = groq_generate_content(topic, context)
                
                if generated_content:
                    st.session_state.generated_content = generated_content
                    
                    # Convert to markdown for preview
                    st.session_state.slide_markdown = pptx_to_markdown(generated_content)
                    
                    # Create PowerPoint file
                    pptx_io = create_presentation(topic, generated_content)
                    st.session_state.presentation_file = pptx_io
                    
                    st.success("Presentation generated successfully! Go to the 'Preview Slides' tab to see your presentation.")
                else:
                    st.error("Failed to generate content. Please try again.")
        else:
            st.warning("Please enter a topic for your presentation.")
    
    # Show the generated content if available
    if st.session_state.generated_content:
        with st.expander("Generated Slide Content", expanded=True):
            edited_content = st.text_area(
                "You can edit this content before finalizing",
                st.session_state.generated_content,
                height=400
            )
            
            if st.button("Update Content", key="update_content"):
                # Update the stored content
                st.session_state.generated_content = edited_content
                
                # Update the markdown for preview
                st.session_state.slide_markdown = pptx_to_markdown(edited_content)
                
                # Update PowerPoint file
                pptx_io = create_presentation(topic, edited_content)
                st.session_state.presentation_file = pptx_io
                
                st.success("Content updated! Go to the 'Preview Slides' tab to see your changes.")
    
    # Download button if presentation is generated
    if st.session_state.presentation_file:
        st.download_button(
            label="Download PowerPoint Presentation",
            data=st.session_state.presentation_file,
            file_name=f"{topic.replace(' ', '_')}_presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )

with tab2:
    st.markdown("### Preview Your Presentation")
    
    if st.session_state.slide_markdown:
        # Display presentation preview
        try:
            st.markdown("#### Interactive Slide Preview")
            rs.slides(st.session_state.slide_markdown, height=500)
        except Exception as e:
            st.error(f"Error displaying slides: {e}")
            st.info("Please install streamlit-reveal-slides using 'pip install streamlit-reveal-slides'")
            
            # Fallback to simple preview
            st.markdown("#### Slide Content Preview")
            st.markdown(st.session_state.slide_markdown)
            
        # Additional download button in preview tab
        if st.session_state.presentation_file:
            st.download_button(
                label="Download PowerPoint Presentation",
                data=st.session_state.presentation_file,
                file_name=f"{topic.replace(' ', '_')}_presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )
    else:
        st.info("Generate a presentation in the 'Create Presentation' tab to see a preview here.")
