import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from groq import Groq
from dotenv import load_dotenv
import os
import requests
from bs4 import BeautifulSoup
import io
import reveal_slides as rs
import concurrent.futures
import json
import base64
from PIL import Image
from io import BytesIO
import re
from urllib.parse import quote

# Set page configuration
st.set_page_config(
    page_title="Advanced Presentation Generator",
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
    .search-result {
        border: 1px solid #e0e0e0;
        border-radius: 5px;
        padding: 10px;
        margin-bottom: 10px;
        background: white;
    }
    .search-result h4 {
        margin-top: 0;
    }
    .theme-preview {
        border: 1px solid #ddd;
        border-radius: 5px;
        padding: 10px;
        text-align: center;
        cursor: pointer;
    }
    .theme-preview.active {
        border: 2px solid #1E3A8A;
        background-color: #f0f4ff;
    }
</style>
""", unsafe_allow_html=True)

# Load environment variables
load_dotenv()
try:
    groq_api_key = st.secrets["k"]["api_key"]
except:
    groq_api_key = os.getenv("GROQ_API_KEY")

# Initialize session state
if 'generated_content' not in st.session_state:
    st.session_state.generated_content = None
if 'slide_markdown' not in st.session_state:
    st.session_state.slide_markdown = None
if 'presentation_file' not in st.session_state:
    st.session_state.presentation_file = None
if 'search_results' not in st.session_state:
    st.session_state.search_results = None
if 'selected_theme' not in st.session_state:
    st.session_state.selected_theme = "professional"
if 'include_images' not in st.session_state:
    st.session_state.include_images = True
if 'num_slides' not in st.session_state:
    st.session_state.num_slides = 5

# Presentation themes
THEMES = {
    "professional": {
        "title_font_size": Pt(36),
        "body_font_size": Pt(18),
        "title_color": RGBColor(31, 58, 138),  # Dark blue
        "accent_color": RGBColor(37, 99, 235),  # Medium blue
        "background_color": RGBColor(255, 255, 255),  # White
    },
    "minimal": {
        "title_font_size": Pt(36),
        "body_font_size": Pt(18),
        "title_color": RGBColor(30, 30, 30),  # Almost black
        "accent_color": RGBColor(100, 100, 100),  # Gray
        "background_color": RGBColor(245, 245, 245),  # Light gray
    },
    "vibrant": {
        "title_font_size": Pt(40),
        "body_font_size": Pt(20),
        "title_color": RGBColor(124, 28, 138),  # Purple
        "accent_color": RGBColor(236, 72, 153),  # Pink
        "background_color": RGBColor(253, 244, 255),  # Very light purple
    },
    "corporate": {
        "title_font_size": Pt(36),
        "body_font_size": Pt(18),
        "title_color": RGBColor(20, 83, 45),  # Dark green
        "accent_color": RGBColor(22, 163, 74),  # Green
        "background_color": RGBColor(240, 253, 244),  # Light green
    },
    "dark": {
        "title_font_size": Pt(38),
        "body_font_size": Pt(18),
        "title_color": RGBColor(226, 232, 240),  # Light gray
        "accent_color": RGBColor(56, 189, 248),  # Light blue
        "background_color": RGBColor(30, 41, 59),  # Dark blue/gray
    }
}

# Function to search the web for information
def search_web(query, num_results=3):
    """Search the web for information related to the query."""
    try:
        # Clean and encode the query
        clean_query = quote(query)
        
        # Make the request to a search engine
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }
        response = requests.get(
            f"https://www.google.com/search?q={clean_query}&num={num_results}", 
            headers=headers
        )
        
        if response.status_code != 200:
            return {"error": f"Failed to retrieve search results: {response.status_code}"}
        
        # Parse the HTML
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Extract search results
        results = []
        search_divs = soup.select('div.g')[:num_results]
        
        for div in search_divs:
            try:
                title_elem = div.select_one('h3')
                if not title_elem:
                    continue
                
                title = title_elem.text
                
                link_elem = div.select_one('a')
                link = link_elem['href'] if link_elem else ""
                
                # Look for snippets
                snippet_elem = div.select_one('div.VwiC3b') or div.select_one('span.aCOpRe')
                snippet = snippet_elem.text if snippet_elem else "No snippet available"
                
                results.append({
                    "title": title,
                    "link": link,
                    "snippet": snippet
                })
            except Exception as e:
                continue
        
        return results
    except Exception as e:
        return {"error": f"Error during web search: {str(e)}"}

# Function to extract content from a webpage
def extract_webpage_content(url):
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }
        response = requests.get(url, headers=headers, timeout=10)
        
        if response.status_code != 200:
            return f"Failed to retrieve page content: {response.status_code}"
        
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Remove script and style elements
        for script in soup(["script", "style"]):
            script.extract()
        
        # Extract text from paragraphs, headers and lists
        content_elements = soup.select('p, h1, h2, h3, h4, h5, h6, li')
        content = ' '.join([elem.get_text() for elem in content_elements])
        
        # Clean up whitespace
        content = re.sub(r'\s+', ' ', content).strip()
        
        # Truncate if too long
        if len(content) > 5000:
            content = content[:5000] + "..."
            
        return content
    except Exception as e:
        return f"Error extracting content: {str(e)}"

# Function to get images for slides
def get_image_for_topic(topic):
    try:
        search_url = f"https://www.bing.com/images/search?q={quote(topic)}&first=1"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }
        response = requests.get(search_url, headers=headers)
        
        if response.status_code != 200:
            return None
            
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Look for image URLs in the page
        img_urls = []
        for img in soup.select('img.mimg'):
            if 'src' in img.attrs and img['src'].startswith('http'):
                img_urls.append(img['src'])
                
        if not img_urls:
            return None
            
        # Get the first valid image
        for img_url in img_urls[:3]:
            try:
                img_response = requests.get(img_url, headers=headers, timeout=5)
                if img_response.status_code == 200:
                    return img_response.content
            except:
                continue
                
        return None
    except Exception as e:
        return None

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

# Function to gather research data using multiple web searches
def gather_research_data(topic, subtopics=None):
    results = {}
    
    # Main topic search
    main_results = search_web(topic, num_results=3)
    results["main"] = main_results
    
    # Search for subtopics if provided
    if subtopics:
        subtopic_results = {}
        for subtopic in subtopics:
            search_query = f"{topic} {subtopic}"
            subtopic_results[subtopic] = search_web(search_query, num_results=2)
        
        results["subtopics"] = subtopic_results
    
    # Get deeper content from the most relevant page
    if isinstance(main_results, list) and len(main_results) > 0:
        try:
            main_url = main_results[0].get("link", "")
            if main_url and main_url.startswith("http"):
                results["detailed_content"] = extract_webpage_content(main_url)
        except:
            pass
    
    return results

# Function to generate slide content using Groq with research data
def groq_generate_content(topic, context, research_data, num_slides=5):
    if not groq_api_key:
        st.error("Please set your GROQ_API_KEY in a .env file or in Streamlit secrets.")
        return None
        
    # Initialize Groq client
    client = Groq(api_key=groq_api_key)
    
    try:
        # Format the research data for the prompt
        research_summary = "Research findings:\n"
        
        if "main" in research_data and isinstance(research_data["main"], list):
            research_summary += "Main topic search results:\n"
            for i, result in enumerate(research_data["main"][:3]):
                research_summary += f"- {result.get('title', 'No title')}: {result.get('snippet', 'No snippet')}\n"
        
        if "subtopics" in research_data:
            research_summary += "\nSubtopic search results:\n"
            for subtopic, results in research_data["subtopics"].items():
                if isinstance(results, list) and results:
                    research_summary += f"- {subtopic}: {results[0].get('snippet', 'No information')}\n"
        
        if "detailed_content" in research_data and research_data["detailed_content"]:
            content_sample = research_data["detailed_content"]
            if len(content_sample) > 1000:
                content_sample = content_sample[:1000] + "..."
            research_summary += f"\nDetailed content excerpt:\n{content_sample}\n"
        
        prompt = f"""Create a professional presentation with {num_slides} slides about "{topic}".

Use the following research data to make the presentation informative and data-driven:
{research_summary}

Additional context provided by the user: {context}

The presentation should follow these guidelines:
1. Start with a compelling title slide
2. Include an agenda or overview slide
3. Each content slide should have a clear, concise title
4. Bullet points should be specific, actionable, and data-driven using the research
5. Use the principle of "one idea per slide"
6. Include a strong concluding slide with actionable takeaways

For each slide, provide:
- A clear title (prefixed with "Title: ")
- 3-5 concise bullet points that elaborate on the title
- Each bullet point should provide valuable information, not just concepts
- Incorporate relevant statistics, facts, or data from the research

Remember to cite sources where appropriate and maintain a professional tone."""
        
        response = client.chat.completions.create(
            messages=[
                {
                    "role": "system",
                    "content": "You are an expert presentation designer who creates well-structured, engaging, and professional slide content backed by research data."
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
        return response.choices[0].message.content
    except Exception as e:
        st.error(f"Error generating content with Groq: {e}")
        return None

# Function to create a PowerPoint presentation with enhanced styling
def create_presentation(topic, slide_content, theme="professional", include_images=True):
    prs = Presentation()
    
    # Set theme properties
    theme_properties = THEMES.get(theme, THEMES["professional"])
    
    # Title Slide
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    
    # Apply theme to title slide
    title.text = topic
    subtitle.text = "Professional Presentation"
    
    # Apply theme formatting to title slide
    for paragraph in title.text_frame.paragraphs:
        paragraph.font.size = theme_properties["title_font_size"]
        paragraph.font.color.rgb = theme_properties["title_color"]
    
    for paragraph in subtitle.text_frame.paragraphs:
        paragraph.font.size = Pt(24)
        paragraph.font.color.rgb = theme_properties["accent_color"]
    
    # Add a background shape if we're using the dark theme
    if theme == "dark":
        left = top = 0
        width = prs.slide_width
        height = prs.slide_height
        shape = slide.shapes.add_shape(1, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = theme_properties["background_color"]
        shape.line.fill.background()
        shape.z_order = 0  # Send to back
        
        # Make title and subtitle text visible against dark background
        for paragraph in title.text_frame.paragraphs:
            paragraph.font.color.rgb = RGBColor(255, 255, 255)
        for paragraph in subtitle.text_frame.paragraphs:
            paragraph.font.color.rgb = RGBColor(220, 220, 220)
    
    # Create content slides
    slides = slide_content.split("\n\n")
    for slide_index, slide_text in enumerate(slides):
        lines = slide_text.split("\n")
        if len(lines) < 2:
            continue
        
        slide_title = lines[0].replace("Title: ", "").strip()
        bullet_points = [point.strip() for point in lines[1:] if point.strip()]
        
        # Add content slide
        content_slide_layout = prs.slide_layouts[1]  # Layout with title and content
        slide = prs.slides.add_slide(content_slide_layout)
        
        # Apply background for dark theme
        if theme == "dark":
            left = top = 0
            width = prs.slide_width
            height = prs.slide_height
            shape = slide.shapes.add_shape(1, left, top, width, height)
            shape.fill.solid()
            shape.fill.fore_color.rgb = theme_properties["background_color"]
            shape.line.fill.background()
            shape.z_order = 0  # Send to back
        
        # Set title
        title = slide.shapes.title
        title.text = slide_title
        
        # Apply theme formatting to title
        for paragraph in title.text_frame.paragraphs:
            paragraph.font.size = theme_properties["title_font_size"]
            paragraph.font.color.rgb = theme_properties["title_color"] if theme != "dark" else RGBColor(255, 255, 255)
        
        # Add bullet points
        body = slide.placeholders[1]
        tf = body.text_frame
        tf.text = ""  # Clear any default text
        
        for point in bullet_points:
            p = tf.add_paragraph()
            p.text = point
            p.font.size = theme_properties["body_font_size"]
            p.font.color.rgb = theme_properties["accent_color"] if theme != "dark" else RGBColor(220, 220, 220)
        
        # Add an image if enabled and not for the last slide
        if include_images and slide_index > 0 and slide_index < len(slides) - 1:
            try:
                # Get an image related to the slide title
                image_data = get_image_for_topic(slide_title)
                
                if image_data:
                    # Save the image to a BytesIO object
                    image_stream = BytesIO(image_data)
                    
                    # Add the image to the slide
                    left = Inches(7)  # Position on the right side
                    top = Inches(2)
                    width = Inches(3)  # Fixed width
                    
                    # Maintain aspect ratio
                    img = Image.open(image_stream)
                    aspect_ratio = img.height / img.width
                    height = Inches(3 * aspect_ratio)
                    
                    # Add the image to the slide
                    slide.shapes.add_picture(image_stream, left, top, width, height)
            except:
                # Continue without an image if there was an error
                pass
    
    # Save the presentation to a BytesIO object
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    
    return pptx_io

# Main application UI with tabs
st.title("Advanced Presentation Generator")
st.markdown("### Create data-driven presentations with AI assistance and web research")

# Main content with tabs for better organization
tab1, tab2, tab3 = st.tabs(["Create Presentation", "Preview Slides", "Research Data"])

with tab1:
    col1, col2 = st.columns([2, 1])
    
    with col1:
        # Topic input
        topic = st.text_input("Presentation Topic", "", help="Enter the main topic of your presentation")
        
        # Context input
        context = st.text_area(
            "Additional Context (optional)",
            "",
            height=100,
            help="Provide any additional information or specific points to include"
        )
    
    with col2:
        # Presentation settings
        st.subheader("Presentation Settings")
        
        # Number of slides
        st.session_state.num_slides = st.slider("Number of Slides", 3, 10, 5)
        
        # Include images option
        st.session_state.include_images = st.checkbox("Include Images in Slides", value=True)
        
        # Theme selection
        st.subheader("Select Theme")
        theme_cols = st.columns(3)
        
        for i, (theme_name, _) in enumerate(THEMES.items()):
            with theme_cols[i % 3]:
                theme_active = st.session_state.selected_theme == theme_name
                if st.button(
                    theme_name.capitalize(), 
                    key=f"theme_{theme_name}",
                    type="primary" if theme_active else "secondary"
                ):
                    st.session_state.selected_theme = theme_name
    
    # Generate slides button
    if st.button("Generate Presentation with Web Research", use_container_width=True, type="primary"):
        if topic:
            with st.spinner("Researching and generating professional slides..."):
                # Perform web research
                research_data = gather_research_data(topic)
                st.session_state.search_results = research_data
                
                # Generate content using research data
                generated_content = groq_generate_content(
                    topic, 
                    context, 
                    research_data,
                    num_slides=st.session_state.num_slides
                )
                
                if generated_content:
                    st.session_state.generated_content = generated_content
                    
                    # Convert to markdown for preview
                    st.session_state.slide_markdown = pptx_to_markdown(generated_content)
                    
                    # Create PowerPoint file
                    pptx_io = create_presentation(
                        topic, 
                        generated_content, 
                        theme=st.session_state.selected_theme,
                        include_images=st.session_state.include_images
                    )
                    st.session_state.presentation_file = pptx_io
                    
                    st.success("Presentation generated successfully! Go to the 'Preview Slides' tab to see your presentation or check the 'Research Data' tab to view your sources.")
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
            
            update_col1, update_col2 = st.columns(2)
            
            with update_col1:
                if st.button("Update Content", key="update_content", use_container_width=True):
                    # Update the stored content
                    st.session_state.generated_content = edited_content
                    
                    # Update the markdown for preview
                    st.session_state.slide_markdown = pptx_to_markdown(edited_content)
                    
                    # Update PowerPoint file
                    pptx_io = create_presentation(
                        topic, 
                        edited_content,
                        theme=st.session_state.selected_theme,
                        include_images=st.session_state.include_images
                    )
                    st.session_state.presentation_file = pptx_io
                    
                    st.success("Content updated! Go to the 'Preview Slides' tab to see your changes.")
            
            with update_col2:
                # Download button if presentation is generated
                if st.session_state.presentation_file:
                    st.download_button(
                        label="Download PowerPoint Presentation",
                        data=st.session_state.presentation_file,
                        file_name=f"{topic.replace(' ', '_')}_presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True,
                                key="download_button_tab1"  # Use a different unique key here
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
                use_container_width=True,
                        key="download_button_tab2"  # Use a different unique key here
            )
    else:
        st.info("Generate a presentation in the 'Create Presentation' tab to see a preview here.")

with tab3:
    st.markdown("### Research Data Sources")
    
    if st.session_state.search_results:
        research_data = st.session_state.search_results
        
        # Display main search results
        st.subheader("Main Topic Research")
        if "main" in research_data and isinstance(research_data["main"], list):
            for result in research_data["main"]:
                with st.container(border=True):
                    st.markdown(f"#### {result.get('title', 'No title')}")
                    st.markdown(f"**Source:** {result.get('link', 'No link')}")
                    st.markdown(f"{result.get('snippet', 'No snippet available')}")
        
        # Display subtopic results if available
        if "subtopics" in research_data:
            st.subheader("Subtopic Research")
            for subtopic, results in research_data["subtopics"].items():
                st.markdown(f"### {subtopic}")
                if isinstance(results, list):
                    for result in results:
                        with st.container(border=True):
                            st.markdown(f"#### {result.get('title', 'No title')}")
                            st.markdown(f"**Source:** {result.get('link', 'No link')}")
                            st.markdown(f"{result.get('snippet', 'No snippet available')}")
        
        # Display excerpt from detailed content if available
        if "detailed_content" in research_data and research_data["detailed_content"]:
            with st.expander("Detailed Content Excerpt"):
                st.markdown(research_data["detailed_content"][:2000] + "..." if len(research_data["detailed_content"]) > 2000 else research_data["detailed_content"])
    else:
        st.info("Generate a presentation in the 'Create Presentation' tab to see research data here.")

st.markdown("---")
st.write("Made By Yogesh Mane!")