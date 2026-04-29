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

    # Remove dot-only lines
    text = re.sub(r"^\s*([\.·•]\s*)+\s*\n?", "", text, flags=re.MULTILINE)
    
    # Fix special unicode spaces
    text = text.replace("\u00A0", " ")

    # Normalize line endings
    text = text.replace("\r", "\n")

    # Trim trailing spaces on each line
    text = re.sub(r"[ \t]+$", "\n", text, flags=re.MULTILINE)

    # Clean extra blank lines
    text = re.sub(r"\n{3,}", "\n\n", text)

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

def highlight_chords(text):
    def repl(match):
        return f"<span style='color:#E2B801'>{match.group()}</span>"

    return re.sub(rf"\[({CHORD_PATTERN})\]", repl, text)

def split_slides(text, max_lines=4):
    lines = [l for l in text.split("\n") if l.strip()]
    slides = []
    current_slide = []
    
    for line in lines:
        # Detect section headers
        if SECTION_REGEX.match(line):
            if current_slide:
                slides.append("\n".join(current_slide))
                current_slide = []

            # Add section title as its own slide
            slides.append(line.upper())
        else:
            current_slide.append(line)

            if len(current_slide) >= max_lines:
                slides.append("\n".join(current_slide))
                current_slide = []

    if current_slide:
        slides.append("\n".join(current_slide))

    return slides

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

        transpose = st.slider("Transpose", -6, 6, 0)

        # Normalize text
        raw_text = normalize_pdf_text(raw_text)

        # Initialize text
        if "edited_text" not in st.session_state:
            st.session_state.edited_text = raw_text

        # Convert button
        if st.button("Convert to Inline Chords"):
            st.sfession_state.edited_text = convert_to_inline(st.session_state.edited_text)

        # Detect sections
        sections = detect_sections(st.session_state.edited_text)
        updated_sections = []

        # Render UI
        for i, section in enumerate(sections):
            st.markdown(f"### Section {i+1}")

            # Use monospace font in Editable Text area
            st.markdown("""
                <style>
                textarea {
                    font-family: monospace !important;
                    white-space: pre !important;
                }
                </style>
                """, unsafe_allow_html=True
            )
            
            new_label = st.text_input(
                "Section",
                value=section["label"],
                key=f"label_{i}"
            )

            new_notes = st.text_input(
                "MD Notes (optional)",
                value=section.get("notes", ""),
                key=f"notes_{i}"
            )

            content = st.text_area(
                "Content",
                value="\n".join(section["lines"]),
                key=f"content_{i}",
                height=150
            )

            updated_sections.append({
                "label": new_label.strip(),
                "notes": new_notes.strip(),
                "lines": content.split("\n")
            })

        # Add section button
        if st.button("➕ Add Section"):
            next_num = len(sections) + 1
            st.session_state.edited_text += f"\n\nVerse {next_num}\n\n"
            st.rerun()
        
        # Rebuild cleaned text
        cleaned_text = ""
        for section in updated_sections:
            label_line = section["label"]

            if section["notes"]:
                label_line += f"\n{section['notes']}"

            cleaned_text += label_line + "\n"

            cleaned_text += "\n".join([l for l in section["lines"] if l.strip()]) + "\n\n"

        # Save
        st.session_state.edited_text = cleaned_text.strip()

    # Apply transpose
    processed_text = transpose_text(st.session_state.edited_text, transpose)

    # Step 3: Preview Slides
    with col2:
        st.subheader("Slide Preview")

        max_lines = st.selectbox("Lines per slide", [3, 4, 5], index=1)
        
        # Build slides from structured sections
        slides = []
        for section in sections:
            slide_lines = []
            slide_lines.append(section["label"])

            if section["notes"]:
                slide_lines.append(f"{section['notes']}")

            for line in section["lines"]:
                if isinstance(line, dict):
                    slide_lines.append(line["text"])
                else:
                    slide_lines.append(line)
            
            slides.append("\n".join(slide_lines))
        
        # Apply slide splitting
        final_slides = []
        for slide in slides:
            final_slides.extend(split_slides(slide, max_lines=max_lines))

        # Render slides    
        for i, slide in enumerate(slides):
            st.markdown(f"**Slide {i+1}**")
            
            slide_html = highlight_chords(slide).replace("\n", "<br>")

            # Match Editable Text
            st.markdown(
                f"""
                <div style="
                    background-color:black;
                    padding:20px;
                    font-size:16px;
                    font-family: monospace;
                    white-space: pre;
                    line-height: 1.3;
                ">
                {slide_html}
                </div>
                """,
                unsafe_allow_html=True
            )

            # st.markdown(
            #     f"<div style='background-color:black; padding:20px; font-size:16px'>{highlight_chords(slide_html)}</div>",
            #     unsafe_allow_html=True
            # )

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

# last updated Apr 29 11:39am
# added regex variables
