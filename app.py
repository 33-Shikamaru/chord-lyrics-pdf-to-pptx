# ----------------------------------------
# PDF Chord and Lyric to PPTX Converter
# ----------------------------------------

# ----------------- Setup ----------------
# Install dependencies in Terminal window using this command:
# pip3 install streamlit pdfplumber python-pptx

import streamlit as st
import pdfplumber
import re
from pptx import Presentation
import time


# --------------- Utilities --------------
CHORD_PATTERN = r"[A-G][#b]?(maj|min|m|dim|aug|sus)?\d*(/[A-G][#b]?)?"

# Accomodate for chord formats like "F#m", "Bbmaj7", "D/F#", etc.
CHORD_REGEX = re.compile(rf"^({CHORD_PATTERN}\s*)+$")

NOTES = ['C', 'C#', 'D', 'D#', 'E', 'F',
         'F#', 'G', 'G#', 'A', 'A#', 'B']

SECTION_HEADERS = ("INTRO", "VERSE", "CHORUS", "BRIDGE", "TAG", "ENDING",
                   "REFRAIN", "INSTRUMENTAL", "INTERLUDE", "VAMP",
                   "TURNAROUND", "PRE-CHORUS", "POST-CHORUS", "OUTRO")

# Accommodate for formats like "Verse 1", "Chorus 2", etc.
SECTION_REGEX = re.compile(rf"^({'|'.join(SECTION_HEADERS)})(\s*\d+)?$", re.IGNORECASE)

SECTION_COLOR = "#6EC138"
CHORD_COLOR   = "#E2B801"
NOTE_COLOR    = "#AAAAAA"
LYRIC_COLOR   = "#FFFFFF"


def extract_text_from_pdf(file):
    text = ""
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
    return text

def normalize_pdf_text(text):
    # Replace long runs of dots with spaces
    text = re.sub(r"[\.·•]{2,}", lambda m: " " * len(m.group()), text)
    
    # Replace signal dots between characters with single space
    text = re.sub(r"(?<=\w)[\.·•](?=\w)", " ", text)

    # Replace dots next to punctuation with single space
    text = re.sub(r"[\.·•]+(?=[^\w\s])", " ", text)
    text = re.sub(r"(?<=[^\w\s])[\.·•]+", " ", text)

    # Remove dot in dot-only lines
    text = re.sub(r"^[ \t]*[\.·•]+[ \t]*$", "", text, flags=re.MULTILINE)
    
    # Fix special unicode spaces
    text = text.replace("\u00A0", " ")

    # Normalize line endings
    text = text.replace("\r", "\n")

    # Trim trailing spaces on each line
    text = re.sub(r"[ \t]+$", "\n", text, flags=re.MULTILINE)

    # Clean extra blank lines
    text = re.sub(r"\n{3,}", "\n\n", text)

    # Remove page numbers
    text = re.sub(r"(?m)^\s*\d+\s*$\n?", "", text)

    return text

def transpose_chord(chord, steps):
    match = re.match(CHORD_PATTERN, chord)
    if not match:
        return chord
    
    root = match.group()
    root = root.replace('b', '#')  # simplify flats
    
    if root not in NOTES:
        return chord
    
    idx = NOTES.index(root)
    new_root = NOTES[(idx + steps) % 12]
    
    return chord.replace(match.group(), new_root, 1)

def transpose_text(text, steps):
    def repl(match):
        return transpose_chord(match.group(), steps)
    
    return CHORD_REGEX.sub(repl, text)

def convert_to_inline(text):
    lines = text.split("\n")
    result = []
    
    i = 0
    while i < len(lines) - 1:
        chord_line = lines[i]
        lyric_line = lines[i + 1]
        
        # Detect likely chord line (many chord, few long words)
        if re.search(rf"\b{CHORD_PATTERN}", chord_line):
            inline = ""
            chord_positions = []

            # Find chord positions
            for match in re.finditer(rf"\b{CHORD_PATTERN}\b", chord_line):
                chord_positions.append((match.start(), match.group()))

            # Insert chord into lyric line based on position
            offset = 0
            for pos, chord in chord_positions:
                insert_pos = min(pos, len(lyric_line))
                inline = (
                    lyric_line[:insert_pos + offset]
                    + f"[{chord}]"
                    + lyric_line[insert_pos + offset:]
                )
            
            result.append(inline if inline else lyric_line)
            i += 2
        else:
            result.append(lines[i])
            i += 1
    
    if i < len(lines):
        result.append(lines[i])
    
    return "\n".join(result)

def detect_sections(text):
    lines = text.split("\n")
    structured = []
    current_section = None

    for line in lines:
        line_clean = re.sub(r"[\.·•]+", " ", line)

        match = re.match(r"^\[?(%s)(\s*\d+)?\b(.*)$" % "|".join(SECTION_HEADERS), line_clean.strip(), re.IGNORECASE)
        if match:
            if current_section:
                structured.append(current_section)
            
            section_type = match.group(1).upper()
            section_num = (match.group(2) or "").strip()
            section_remainder = match.group(3).strip()

            # Normalize header remainder
            section_remainder = re.sub(r"\s*[:\.]+\s*", ":", section_remainder, count=1)
            
            # Extract MD notes by first colon
            if ":" in section_remainder:
                parts = section_remainder.split(":", 1)
                notes_raw = parts[1].strip()
            else:
                notes_raw = ""
            
            # Clean notes
            notes = notes_raw.strip()
            notes = re.sub(r"\s+", " ", notes)

            current_section = {
                "label": f"{section_type} {section_num}".strip(),
                "notes": notes,
                "lines": []
            }
        else:
            if not current_section:
                current_section = {
                    "label": "VERSE",
                    "notes": "",
                    "lines": []
                }
            current_section["lines"].append(line)

    if current_section:
        structured.append(current_section)

    return structured

def load_css():
    st.markdown("""
    <style>
    textarea {
        font-family: "Courier New", monospace !important;
        font-size: 16px !important;
        line-height: 1.4 !important;
    }

    .preview-container {
        overflow-y: auto;
        padding-right: 10px;
        border: 1px solid #333;
        border-radius: 8px;
    }

    .preview-container::-webkit-scrollbar {
        width: 6px;
    }
                
    .preview-container::-webkit-scrollbar-thumb {
        background: #555;
        border-radius: 4px;
    }
                
    .slide-separator {
        text-align: center;
        color: #888;
        font-size: 12px;
        margin: 20px 0 10px 0;
    }

    .slide-box {
        background-color: black;
        padding: 20px;
        font-size: 16px;
        font-family: monospace;
        line-height: 1.3;
        margin-bottom: 10px;
        border-radius: 6px;
    }    
    
    </style>
    """, unsafe_allow_html=True)
load_css()

def highlight_chords(text):
    def repl(match):
        return f"<span style='color:#E2B801'>{match.group()}</span>"

    return re.sub(rf"\[({CHORD_PATTERN})\]", repl, text)

def split_slides(text):
    # Normalize trailing spaces
    text = re.sub(r"[ \t]+$", "", text, flags=re.MULTILINE)

    # Split on --- lines
    slides = re.split(r"(?m)^\s*---\s*$", text)

    return [s.strip() for s in slides if s.strip()]

def split_sections(text):
    sections = re.split(r"\n\s*===\s*\n", text.strip())
    return [s.strip() for s in sections if s.strip()]

def create_ppt(slides):
    prs = Presentation()
    
    for slide_text in slides:
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        
        slide.shapes.title.text = ""
        slide.placeholders[1].text = slide_text
    
    file_path = "/mnt/data/output.pptx"
    prs.save(file_path)
    return file_path


# --------------- UI / App ---------------
st.set_page_config(layout="wide")

st.title("🎵 Chord & Lyrics PDF → PPTX Slides")

# Step 1: Upload
st.header("Step 1 — Upload")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])
pasted_text = st.text_area("Or paste text")

if uploaded_file:
    raw_text = extract_text_from_pdf(uploaded_file)
elif pasted_text:
    raw_text = pasted_text
else:
    raw_text = ""

# Step 2: Review & Fix
if raw_text:
    st.header("Step 2 — Review & Fix")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Editable Text")

        # Intialize text box once
        if "edited_text" not in st.session_state:
            st.session_state.edited_text = normalize_pdf_text(raw_text)
        
        # Dynamically adjust box height based on content
        lines = st.session_state.edited_text.count("\n") + 1
        height = min(1000, max(400, lines * 24))

        st.info("""
        **Tips**
        - Use `---` for slide breaks 
        - Chord alignment off? Try zooming out or widening the window
        """)

        edited_text = st.text_area(
            "Edit text here:",
            value=st.session_state.edited_text,
            height=height,
            key="editor"
        )

        
    
        st.session_state.edited_text = edited_text

        # Capture time to prevent excessive re-renders during typing
        if "last_edit_time" not in st.session_state:
            st.session_state.last_edit_time = 0

        st.session_state.last_edit_time = time.time()

        # Actions
        #transpose = st.slider("Transpose", -6, 6, 0)

        colA, colB = st.columns(2)

        #with colA:
            # if st.button("Convert to Inline Chords", key="convert"):
            #     st.session_state.edited_text = convert_to_inline(
            #         st.session_state.edited_text
            #     )
            #     st.rerun()

       # with colB:
            
    # Apply transpose
    #processed_text = transpose_text(st.session_state.edited_text, #transpose)

    # Step 3: Preview Slides
    with col2:
        st.subheader("Slide Preview")

        # Get time since last edit
        last_edit_time = st.session_state.get("last_edit_time", 0)
        last_good_text = st.session_state.get("last_good_text", edited_text)

        if time.time() - last_edit_time > 0.3:
            text_to_use = edited_text
            st.session_state["last_good_text"] = edited_text
        else:
            text_to_use = last_good_text

        slides = split_slides(text_to_use)
        slides_html= ""

        for i, slide in enumerate(slides):
            formatted_slide = highlight_chords(slide).replace("\n", "<br>")

            # Create a slide separator
            slides_html += f"""<div class="slide-separator">
            ──────── Slide {i+1} ────────
            </div>
            <div class="slide-box">
            {formatted_slide}
            </div>"""

        # Render Slide Preview
        st.markdown(f"""<div class="preview-container" style="height:{height}px;">
        {slides_html}
        </div>""",
            unsafe_allow_html=True
        )

    # Step 4: Export
    st.header("Step 3 — Export")

    if st.button("Generate PPTX"):
        file_path = create_ppt(slides)
        
        with open(file_path, "rb") as f:
            st.download_button(
                label="Download PPTX",
                data=f,
                file_name="slides.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
