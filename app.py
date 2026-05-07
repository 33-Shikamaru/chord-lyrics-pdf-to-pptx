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
CHORD_PATTERN = r"^[A-G](#|b)?(m|maj|min|sus|dim|aug)?\d*(\/[A-G](#|b)?)?$"

# For chord line detection (is this line mostly chords?)
# Accomodate for chord formats like "F#m", "Bbmaj7", "D/F#", etc
CHORD_LINE_REGEX = re.compile(rf"^({CHORD_PATTERN}\s*)+$")

# For transposing chords
CHORD_TOKEN_REGEX = re.compile(
    r"(?<![/A-Za-z])([A-G](?:#|b)?(?:m|maj|min|sus|dim|aug|add)?\d*(?:/[A-G](?:#|b)?)?)"
)

# For transposing chords (only need notes not entire chord)
ROOT_REGEX = r"^([A-G](?:#|b)?)"

NOTES = ['C', 'C#', 'D', 'D#', 'E', 'F',
         'F#', 'G', 'G#', 'A', 'A#', 'B']

# Flat conversion map
FLAT_MAP = {'Db': 'C#', 
            'Eb': 'D#',
            'Gb': 'F#',
            'Ab': 'G#',
            'Bb': 'A#'}

SECTION_HEADERS = ("INTRO", "VERSE", "CHORUS", "BRIDGE", "TAG", "ENDING",
                   "REFRAIN", "INSTRUMENTAL", "INTERLUDE", "VAMP", "BREAKDOWN",
                   "TURNAROUND", "PRE-CHORUS", "POST-CHORUS", "OUTRO")

# Accommodate for formats like "Verse 1", "Chorus 2", etc.
SECTION_REGEX = re.compile(rf"^({'|'.join(SECTION_HEADERS)})(\s*\d+)?$", re.IGNORECASE)

def extract_text_from_pdf(file):
    text = ""
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
    return text

def normalize_pdf_text(text):
    # Normalize line endings
    text = text.replace("\r", "\n")

    # Remove page numbers
    text = re.sub(r"(?m)^\s*\d+\s*$\n?", "", text)

    # Remove dot-only lines
    text = re.sub(r"(?m)^[ \t]*[\.·•]+[ \t]*$", "", text)

    # Replace leading dots with spaces
    text = re.sub(
        r"(?m)^(\s*)([\.]+)",
        lambda m: m.group(1) + " " * len(m.group(2)),
        text
    )

    # Replace other dot runs with spaces
    text = re.sub(
        r"[\.·•]+",
        lambda m: " " * len(m.group()),
        text
    )

    # Normalize unicode spaces
    text = text.replace("\u00A0", " ")

    # Trim trailing spaces
    text = re.sub(r"[ \t]+$", "", text, flags=re.MULTILINE)

    # Clean excessive blank lines
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text

def transpose_text(text, steps):
    lines = text.split("\n")
    result = []
    for line in lines:
        # Only transpose chord lines
        if detect_chord_line(line):
            
            def transpose_match(match):
                # Original chord text
                chord = match.group()


                # Split slash chords
                if "/" in chord:
                    main, bass = chord.split("/", 1)
                else:
                    main = chord
                    bass = None

                # Extract root + suffix
                root_match = re.match(r'^([A-G](?:#|b)?)(.*)$', main)

                if not root_match:
                    return chord
                root = root_match.group(1)
                suffix = root_match.group(2)

                # Normalize flats
                root = FLAT_MAP.get(root, root)

                # Transpose root
                if root in NOTES:
                    idx = NOTES.index(root)
                    new_root = NOTES[(idx + steps) % 12]
                else:
                    new_root = root

                new_chord = new_root + suffix

                # Transpose bass note
                if bass:
                    bass = FLAT_MAP.get(bass, bass)
                    if bass in NOTES:
                        bass_idx = NOTES.index(bass)
                        new_bass = NOTES[(bass_idx + steps) % 12]
                    else:
                        new_bass = bass

                    new_chord += "/" + new_bass

                return new_chord

            line = CHORD_TOKEN_REGEX.sub(transpose_match, line)
        
        result.append(line)

    return "\n".join(result)

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

def detect_chord_line(text):
    # Contains bar notation
    if "|" in text.strip():
        return True

    # Contains many chord-like tokens (D, G2, Asus4, F#, Bm7, etc.)
    tokens = text.strip().split()
    chord_like = sum(bool(re.match(CHORD_PATTERN, t)) for t in tokens)

    return chord_like >= max(1, len(tokens) // 2) 

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
    st.html("""
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
        font-size: 16px;
        font-family: monospace;
        white-space: pre;
        padding: 20px;
        margin-bottom: 10px;
        border-radius: 6px;
        line-height: 1.3;
        margin: 0; 
    }
                
    .section {
        display: block;
        text-align: left !important;
    }
                
    .section { color: #6EC138; }    
    .chord   { color: #E2B801; }    
    .note    { 
        color: #AAAAAA;             
        font-style: italic; 
    }    
    .lyric   { color: #FFFFFF; }    
            
    </style>
    """)
load_css()

def split_inline(lines):
    result = []
    i = 0

    while i < len(lines):
        line = lines[i]

        # --- Case 1: Chord + lyric pair ---
        if detect_chord_line(line) and i + 1 < len(lines):
            chord_line = line
            lyric_line = lines[i + 1]

            if "||" in chord_line or "||" in lyric_line:
                source = chord_line if "||" in chord_line else lyric_line

                # Find split indices
                split_indices = []
                idx = 0
                while "||" in source[idx:]:
                    pos = source.index("||", idx)
                    split_indices.append(pos)
                    idx = pos + 2

                # Remove markers
                chord_clean = chord_line.replace("||", "")
                lyric_clean = lyric_line.replace("||", "")

                prev = 0
                offset = 0

                for pos in split_indices:
                    adj = pos - offset

                    c_part = chord_clean[prev:adj]
                    l_part = lyric_clean[prev:adj]

                    # Preserve trailing bar if needed
                    c_part = c_part.rstrip()
                    if "|" in c_part and not c_part.endswith("|"):
                        c_part += "|"

                    result.append(c_part)
                    result.append(l_part)

                    prev = adj
                    offset += 2

                # Final segment
                c_part = chord_clean[prev:]
                l_part = lyric_clean[prev:]

                c_part = c_part.rstrip()
                if "|" in c_part and not c_part.endswith("|"):
                    c_part += "|"

                result.append(c_part)
                result.append(l_part)

            else:
                result.append(chord_line)
                result.append(lyric_line)

            i += 2

        # --- Case 2: Single line ---
        else:
            if "||" in line:
                parts = line.split("||")

                for p in parts:
                    p_clean = p.rstrip()

                    if "|" in p_clean and not p_clean.endswith("|"):
                        p_clean += "|"

                    result.append(p_clean)
            else:
                result.append(line)

            i += 1

    return result

def split_slides(text):
    # Normalize trailing spaces
    text = re.sub(r"[ \t]+$", "", text, flags=re.MULTILINE)

    # Split on --- lines
    slides = re.split(r"(?m)^\s*---\s*$", text)

    return [s.strip() for s in slides if s.strip()]

def format_slides(text):
    lines = text.split("\n")
    cleaned_lines = []

    for line in lines:
        # Split indentation from content
        indent = re.match(r"^(\s*)", line).group(1)
        content = line[len(indent):]

        # Section headers
        if re.match(rf"^\s*({'|'.join(SECTION_HEADERS)})(\s*\d+)?\s*:?", content.strip(), re.IGNORECASE):
            formatted_line = indent + f"<span class='section'>{content}</span>"
        
        # Chord lines
        elif detect_chord_line(content):
            formatted_line = indent + f"<span class='chord'>{content}</span>"

        # Lyrics 
        else:
            formatted_line = indent + f"<span class='lyric'>{content}</span>"

        # Notes (apply last to account for inline notes)
        formatted_line = re.sub(
            r"\(.*?\)",
            lambda m: f"<span class='note'>{m.group()}</span>",
            formatted_line
        )

        # Append lines to list
        cleaned_lines.append(formatted_line)

    return "\n".join(cleaned_lines)

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

# ---------------
# Step 1: Upload
# ---------------
st.header("Step 1 — Upload")

uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])
pasted_text = st.text_area("Or paste text")

if uploaded_file:
    raw_text = extract_text_from_pdf(uploaded_file)
elif pasted_text:
    raw_text = pasted_text
else:
    raw_text = ""

# ---------------------
# Step 2: Review & Fix
# ---------------------
if raw_text:
    st.header("Step 2 — Review & Fix")

    # ----------------------
    # Key transpose section
    # ----------------------
    KEY_OPTIONS = [
        "Select key...",
        "C", "C#", "D", "D#", "E", "F",
        "F#", "G", "G#", "A", "A#", "B"
    ]

    col1, col2, col3 = st.columns([1, 0.3, 1])

    with col1:
        original_key = st.selectbox("Original Key", KEY_OPTIONS, index=0)

    with col2:
        st.html("<div style='text-align:center;padding-top:30px;'>→</div>")

    with col3:
        target_key = st.selectbox("Transpose To", KEY_OPTIONS[1:], disabled=(original_key == "Select key..."))
    

    # Normalize keys for calculation
    if original_key == "Select key...":
        steps = 0
    else:
        calc_original = FLAT_MAP.get(original_key, original_key)
        calc_target = FLAT_MAP.get(target_key, target_key)

        steps = (
            NOTES.index(calc_target)
            - NOTES.index(calc_original)
        ) % 12

    # -------------------
    # Textboxes section
    # -------------------
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Editable Text",
                     help="• Line break: || • Slide break: --- (Press ⌘ + Return to apply changes) • Zoom out if chord misaligned")
        # Intialize text box once
        if "edited_text" not in st.session_state:
            st.session_state.edited_text = normalize_pdf_text(raw_text)
        
        # Dynamically adjust box height based on content
        lines = st.session_state.edited_text.count("\n") + 1
        height = min(1000, max(400, lines * 24))
        offset = 26     # to account for the Preview Slides padding  

        # Render Edited Text box
        edited_text = st.text_area(
            "",
            value=st.session_state.edited_text,
            height=height + offset,
            key="editor"
        )

        st.session_state.edited_text = edited_text

        # Capture time to prevent excessive re-renders during typing
        if "last_edit_time" not in st.session_state:
            st.session_state.last_edit_time = 0

        st.session_state.last_edit_time = time.time()

    # -----------------------
    # Step 3: Preview Slides
    # -----------------------
    with col2:
        st.subheader("Slide Preview")

        # Get time since last edit
        last_edit_time = st.session_state.get("last_edit_time", 0)
        last_good_text = st.session_state.get("last_good_text", edited_text)

        if time.time() - last_edit_time > 0.3:
            text_to_use = edited_text
            st.session_state["last_good_text"] = text_to_use
        else:
            text_to_use = last_good_text

        slides = split_slides(text_to_use)
        slides_html= ""

        for i, slide in enumerate(slides):
            # Split into lines
            lines = slide.split("\n")

            # Split lines on ||
            lines = split_inline(lines)

            # Rejoin slide
            processed_slide = "\n".join(lines)

            # Apply transpose
            processed_slide = transpose_text(processed_slide, steps)

            # Format slides
            formatted_slide = format_slides(processed_slide)
            formatted_slide = formatted_slide.replace("\n", "<br>")

            # Create a slide separator
            slides_html += f"""<div class="slide-separator">
            ──────── Slide {i+1} ────────
            </div>
            <div class="slide-box">
            {formatted_slide}
            </div>"""

        # Render Slide Preview
        st.html("<div style='height: 12px;'></div>")
        st.html(f"""<div class="preview-container" style="height:{height}px;">
       {slides_html}
        </div>"""
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
