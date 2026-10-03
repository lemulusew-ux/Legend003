import io
import os
from datetime import date

import streamlit as st
from PyPDF2 import PdfReader
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from openai import OpenAI


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Legend Engineering | Letter Processor",
    page_icon="📄",
    layout="wide",
)


# ============================================================
# OPENAI CLIENT
# ============================================================

def get_api_key():
    """Read the OpenAI API key from Streamlit secrets or environment."""
    try:
        if "OPENAI_API_KEY" in st.secrets:
            return st.secrets["OPENAI_API_KEY"]
    except Exception:
        pass

    return os.getenv("OPENAI_API_KEY")


def get_client():
    api_key = get_api_key()
    if not api_key:
        return None
    return OpenAI(api_key=api_key)


# ============================================================
# PDF PROCESSING
# ============================================================

def extract_text_from_pdf(pdf_file):
    """Extract selectable text from a PDF."""
    reader = PdfReader(pdf_file)
    pages = []
    for page in reader.pages:
        text = page.extract_text() or ""
        pages.append(text)

    result = "\n\n".join(pages).strip()

    if not result:
        raise ValueError(
            "No selectable text was found in this PDF. "
            "It may be a scanned/image-only PDF. OCR is required for such files."
        )

    return result


# ============================================================
# OPENAI FUNCTIONS
# ============================================================

def ask_openai(instructions, input_text, model):
    client = get_client()

    if client is None:
        raise RuntimeError(
            "OpenAI API key is not configured. Add OPENAI_API_KEY "
            "to Streamlit Cloud Secrets."
        )

    response = client.responses.create(
        model=model,
        instructions=instructions,
        input=input_text,
    )

    return response.output_text.strip()


def summarize_incoming_letter(letter_text, model):
    instructions = """
You are an executive technical assistant for an engineering and infrastructure
company.

Analyze the incoming letter and produce a concise, factual executive summary.

Use exactly these headings:

**Sender / Organization:**
**Letter Ref Number / Date:**
**Subject / Core Matter:**
**Key Action Items Required:**
**Deadlines / Urgency Level:**
**Contractual / Technical Issues:**
**Recommended Response Considerations:**

Do not invent missing information. If information is not stated in the letter,
write "Not stated in the incoming letter."

Distinguish clearly between facts stated in the letter and your interpretation.
"""

    return ask_openai(
        instructions,
        f"INCOMING LETTER:\n\n{letter_text}",
        model,
    )


def draft_response_letter(
    incoming_text,
    summary,
    user_instructions,
    company_name,
    sender_name,
    recipient_name,
    recipient_organization,
    letter_date,
    model,
):
    instructions = f"""
You are a Senior Civil Engineer and corporate correspondence specialist
drafting an official response letter on behalf of "{company_name}".

Prepare a formal, precise and contractually careful business letter.

IMPORTANT RULES:
1. Never invent project facts, dates, commitments, quantities, contract clauses,
   names or reference numbers.
2. Use placeholders such as [REFERENCE NO.] or [DATE] where information is absent.
3. Do not make legal conclusions unless explicitly supported by the supplied text.
4. Preserve the user's stated decision and intent.
5. Do not unnecessarily accept liability or responsibility.
6. Keep the tone professional, respectful and firm where appropriate.
7. Generate ONLY the letter itself — no explanation before or after it.

LETTER DATE:
{letter_date}

RECIPIENT:
{recipient_name}
{recipient_organization}

SIGNATORY:
{sender_name}

COMPANY:
{company_name}

INCOMING LETTER SUMMARY:
{summary}

USER'S RESPONSE DIRECTION:
{user_instructions}

INCOMING LETTER:
{incoming_text}

FORMAT:

[DATE]

[REFERENCE NO.]

To:
[ADDRESSEE]
[ORGANIZATION]
[ADDRESS IF KNOWN]

SUBJECT: [CLEAR UPPERCASE SUBJECT]

Dear Sir/Madam,

[Opening acknowledgement and reference to incoming correspondence.]

[Main response, addressing the issues logically.]

[Required actions, submissions, commitments or reservations.]

[Closing paragraph.]

Sincerely,

[NAME]
[TITLE]
{company_name}
"""

    return ask_openai(
        instructions,
        "Draft the official response letter now.",
        model,
    )


# ============================================================
# WORD DOCUMENT GENERATION
# ============================================================

def add_bottom_border(paragraph):
    """Add a subtle bottom border to a paragraph using OOXML."""
    p = paragraph._p
    pPr = p.get_or_add_pPr()

    pbdr = pPr.find(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr"
    )

    if pbdr is None:
        from docx.oxml import OxmlElement
        pbdr = OxmlElement("w:pBdr")
        pPr.append(pbdr)

    from docx.oxml import OxmlElement
    bottom = OxmlElement("w:bottom")
    bottom.set(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val",
        "single",
    )
    bottom.set(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz",
        "8",
    )
    bottom.set(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}space",
        "1",
    )
    bottom.set(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color",
        "666666",
    )
    pbdr.append(bottom)


def generate_word_document(
    letter_body,
    company_name,
    sender_name,
    sender_title,
):
    doc = Document()

    section = doc.sections[0]
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

    # Default font
    styles = doc.styles
    styles["Normal"].font.name = "Calibri"
    styles["Normal"].font.size = Pt(11)

    # Header
    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    run = hp.add_run(company_name.upper())
    run.bold = True
    run.font.name = "Calibri"
    run.font.size = Pt(14)

    hp2 = header.add_paragraph()
    hp2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r2 = hp2.add_run("ENGINEERING • CONSULTANCY • INFRASTRUCTURE")
    r2.font.size = Pt(8)

    add_bottom_border(hp2)

    # Body
    for block in letter_body.split("\n\n"):
        block = block.strip()
        if not block:
            continue

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.line_spacing = 1.15

        # Preserve line breaks within blocks
        lines = block.splitlines()

        for i, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue

            # Bold subject
            if line.upper().startswith("SUBJECT:"):
                run = p.add_run(line)
                run.bold = True
                run.font.name = "Calibri"
                run.font.size = Pt(11)
            else:
                run = p.add_run(line)
                run.font.name = "Calibri"
                run.font.size = Pt(11)

            if i < len(lines) - 1:
                run.add_break()

    # Signature if not already included
    if sender_name.strip():
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(18)

        r = p.add_run(sender_name)
        r.bold = True
        r.font.size = Pt(11)

        p2 = doc.add_paragraph(sender_title)
        p2.paragraph_format.space_after = Pt(0)

        p3 = doc.add_paragraph(company_name)
        p3.paragraph_format.space_after = Pt(0)

    # Footer
    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = fp.add_run(f"{company_name} | Official Correspondence")
    fr.font.size = Pt(8)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    return output


# ============================================================
# SESSION STATE
# ============================================================

defaults = {
    "extracted_text": "",
    "summary": "",
    "drafted_letter": "",
    "uploaded_name": "",
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title("Corporate Settings")

company_name = st.sidebar.text_input(
    "Company Name",
    value="Legend Engineering PLC",
)

sender_name = st.sidebar.text_input(
    "Signatory Name",
    value="",
    placeholder="e.g. Eng. Mulusew Hunegnaw Gebeyehu",
)

sender_title = st.sidebar.text_input(
    "Signatory Title",
    value="General Manager",
)

model = st.sidebar.selectbox(
    "AI Model",
    options=[
        "gpt-5-mini",
        "gpt-5",
    ],
    index=0,
)

st.sidebar.divider()

st.sidebar.caption(
    "API keys are read from Streamlit Secrets and are never stored in this source code."
)

if not get_api_key():
    st.sidebar.warning(
        "OPENAI_API_KEY is not configured."
    )
else:
    st.sidebar.success("OpenAI API key detected.")


# ============================================================
# MAIN UI
# ============================================================

st.title("📄 Letter Processor & Response Drafter")
st.caption(
    "Executive correspondence assistant for engineering and infrastructure projects"
)

st.info(
    "Workflow: Upload/Paste → Extract → Summarize → Review → Draft Response → Edit → Download Word"
)

left, right = st.columns(2)

# ------------------------------------------------------------
# LEFT COLUMN
# ------------------------------------------------------------

with left:
    st.subheader("1. Incoming Letter")

    input_type = st.radio(
        "Input Method",
        ["Upload PDF", "Paste Text"],
        horizontal=True,
    )

    if input_type == "Upload PDF":
        uploaded_file = st.file_uploader(
            "Upload incoming letter",
            type=["pdf"],
        )

        if uploaded_file:
            st.caption(f"Selected: {uploaded_file.name}")

            if st.button(
                "🔎 Extract & Summarize",
                use_container_width=True,
            ):
                try:
                    with st.spinner("Reading PDF and analyzing the letter..."):
                        extracted = extract_text_from_pdf(uploaded_file)

                        st.session_state.extracted_text = extracted
                        st.session_state.uploaded_name = uploaded_file.name

                        st.session_state.summary = summarize_incoming_letter(
                            extracted,
                            model,
                        )

                        st.session_state.drafted_letter = ""

                    st.success("Letter processed successfully.")

                except Exception as e:
                    st.error(f"Processing failed: {e}")

    else:
        pasted_text = st.text_area(
            "Paste the incoming letter here",
            height=300,
            placeholder="Paste the complete incoming correspondence...",
        )

        if st.button(
            "🔎 Analyze Letter",
            use_container_width=True,
        ):
            if not pasted_text.strip():
                st.warning("Please paste the incoming letter first.")
            else:
                try:
                    with st.spinner("Analyzing the letter..."):
                        st.session_state.extracted_text = pasted_text
                        st.session_state.summary = summarize_incoming_letter(
                            pasted_text,
                            model,
                        )
                        st.session_state.drafted_letter = ""

                    st.success("Letter analyzed successfully.")

                except Exception as e:
                    st.error(f"Analysis failed: {e}")

    if st.session_state.summary:
        st.subheader("Executive Summary")
        st.markdown(st.session_state.summary)

        with st.expander("View Extracted Letter Text"):
            st.text_area(
                "Extracted text",
                st.session_state.extracted_text,
                height=300,
                label_visibility="collapsed",
            )


# ------------------------------------------------------------
# RIGHT COLUMN
# ------------------------------------------------------------

with right:
    st.subheader("2. Response Decision")

    recipient_name = st.text_input(
        "Recipient Name",
        placeholder="e.g. Project Manager",
    )

    recipient_organization = st.text_input(
        "Recipient Organization",
        placeholder="e.g. SHEGIZ Consulting Engineers",
    )

    letter_date = st.date_input(
        "Response Letter Date",
        value=date.today(),
    )

    user_directions = st.text_area(
        "Your instructions / response decision",
        height=180,
        placeholder=(
            "Example:\n"
            "Acknowledge receipt of the consultant's letter. "
            "Confirm that the revised drawings will be submitted after "
            "incorporating the agreed comments. Request seven additional "
            "working days for finalization."
        ),
    )

    if st.button(
        "✍️ Generate Official Response",
        type="primary",
        use_container_width=True,
    ):
        if not st.session_state.extracted_text:
            st.error("Please process an incoming letter first.")
        elif not user_directions.strip():
            st.warning("Please enter your response decision/instructions.")
        else:
            try:
                with st.spinner("Drafting formal response letter..."):
                    st.session_state.drafted_letter = draft_response_letter(
                        incoming_text=st.session_state.extracted_text,
                        summary=st.session_state.summary,
                        user_instructions=user_directions,
                        company_name=company_name,
                        sender_name=sender_name,
                        recipient_name=recipient_name,
                        recipient_organization=recipient_organization,
                        letter_date=letter_date.strftime("%d %B %Y"),
                        model=model,
                    )

                st.success("Draft generated.")

            except Exception as e:
                st.error(f"Drafting failed: {e}")


# ============================================================
# GENERATED LETTER
# ============================================================

if st.session_state.drafted_letter:
    st.divider()
    st.subheader("3. Review & Edit")

    edited_draft = st.text_area(
        "Editable response letter",
        value=st.session_state.drafted_letter,
        height=500,
    )

    st.session_state.drafted_letter = edited_draft

    col_a, col_b = st.columns(2)

    with col_a:
        word_file = generate_word_document(
            edited_draft,
            company_name,
            sender_name,
            sender_title,
        )

        st.download_button(
            label="💾 Download Official Word Document",
            data=word_file,
            file_name="Official_Response_Letter.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
        )

    with col_b:
        st.download_button(
            label="📄 Download Letter as TXT",
            data=edited_draft,
            file_name="Official_Response_Letter.txt",
            mime="text/plain",
            use_container_width=True,
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()
st.caption(
    f"{company_name} | Letter Processor & Response Drafter"
)
