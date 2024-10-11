import streamlit as st
import pandas as pd
import io
from openai import OpenAI
from docx import Document

# Testing header to see if the deployment works
st.header("Testing Deployment 11")

# Sidebar inputs for OpenAI API key and file upload
st.sidebar.header("Configuration")
api_key = st.secrets["openai_api_key"]
uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=["xlsx"])

# Hardcoded template path
template_path = "template.docx"

# Initialize session state variables
if 'summary' not in st.session_state:
    st.session_state.summary = ''
if 'summary_locked' not in st.session_state:
    st.session_state.summary_locked = False
if 'edited_content' not in st.session_state:
    st.session_state.edited_content = []
if 'show_preview' not in st.session_state:
    st.session_state.show_preview = False

def update_summary():
    st.session_state.summary = st.session_state.new_summary
    st.session_state.summary_locked = True

def toggle_summary_lock():
    st.session_state.summary_locked = not st.session_state.summary_locked

def update_row(index, value):
    st.session_state.edited_content[index] = value
    st.session_state[f'row_{index+1}_locked'] = not st.session_state[f'row_{index+1}_locked']

def toggle_preview():
    st.session_state.show_preview = not st.session_state.show_preview

if uploaded_file is not None:
    # Initialize the OpenAI client once the API key is provided
    client = OpenAI(api_key=api_key)

    # Read Excel file into DataFrame (starting from the first row)
    df = pd.read_excel(uploaded_file, usecols="A", nrows=10)

    # Initialize edited_content if it's empty
    if not st.session_state.edited_content:
        st.session_state.edited_content = df.iloc[:, 0].tolist()

    # Hardcoded prompt for executive summary
    custom_prompt = "Generate an executive summary based on the following content:"

    # Single button for generating or refreshing the executive summary
    if st.button("Generate/Refresh Executive Summary"):
        with st.spinner("Generating executive summary..."):
            try:
                combined_content = "\n".join(map(str, st.session_state.edited_content))
                final_prompt = f"{custom_prompt}\n{combined_content}"
                response = client.chat.completions.create(
                    model="gpt-4-1106-preview",
                    messages=[
                        {
                            "role": "user",
                            "content": final_prompt
                        }
                    ],
                    temperature=0.7
                )
                st.session_state.summary = response.choices[0].message.content
                st.session_state.summary_locked = False
            except Exception as e:
                st.error(f"Failed to generate summary: {e}")

    # Display and edit executive summary
    st.subheader("Executive Summary:")
    if st.session_state.summary_locked:
        st.write(st.session_state.summary)
        if st.button("Edit Executive Summary"):
            toggle_summary_lock()
    else:
        st.text_area("Edit the Executive Summary", 
                     key="new_summary",
                     value=st.session_state.summary,
                     height=200)
        if st.button("Save Executive Summary"):
            update_summary()

    # Display content of column A up to row 10 in editable fields
    st.write("Edit the content of column A (rows 1-10):")
    for i, content in enumerate(st.session_state.edited_content):
        if f'row_{i+1}_locked' not in st.session_state:
            st.session_state[f'row_{i+1}_locked'] = False

        col1, col2 = st.columns([3, 1])
        with col1:
            if st.session_state[f'row_{i+1}_locked']:
                st.text_area(f"Row {i+1} (Locked)", value=content, key=f"row_{i+1}", disabled=True)
            else:
                st.text_area(f"Row {i+1}", value=content, key=f"row_{i+1}")
        with col2:
            if st.button("Toggle Edit" if st.session_state[f'row_{i+1}_locked'] else "Save", key=f"button_{i+1}"):
                update_row(i, st.session_state[f"row_{i+1}"])

    # Preview button
    if st.button("Toggle Preview"):
        toggle_preview()

    # Show preview
    if st.session_state.show_preview:
        st.subheader("Document Preview:")
        st.write("Executive Summary:")
        st.write(st.session_state.summary)
        st.write("Content:")
        for i, content in enumerate(st.session_state.edited_content):
            st.write(f"Row {i+1}: {content}")

    # Confirm and generate Word document
    if st.button("Confirm and Generate Word Document"):
        try:
            # Load the Word template
            template = Document(template_path)

            # Replace placeholders in the template
            for paragraph in template.paragraphs:
                if "{{Executive_Summary}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{{Executive_Summary}}", st.session_state.summary)

                for i, content in enumerate(st.session_state.edited_content):
                    placeholder = f"{{{{Row_{i+1}}}}}"
                    if placeholder in paragraph.text:
                        paragraph.text = paragraph.text.replace(placeholder, str(content))

            # Save the generated document to a buffer
            buffer = io.BytesIO()
            template.save(buffer)
            buffer.seek(0)

            # Provide the Word document for download
            st.download_button(
                label="Download Word Document",
                data=buffer,
                file_name="generated_report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Failed to generate Word document: {e}")
        