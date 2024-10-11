import streamlit as st
import pandas as pd
import io
from openai import OpenAI
from docx import Document

# Testing header to see if the deployment works
st.header("Testing Deployment 15")

# Sidebar inputs for OpenAI API key, password, and file upload
st.sidebar.header("Configuration")
api_key = st.secrets["openai_api_key"]
password = st.sidebar.text_input("Enter password", type="password")
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

def update_summary():
    st.session_state.summary = st.session_state.new_summary
    st.session_state.summary_locked = True

def unlock_summary():
    st.session_state.summary_locked = False

def update_row(index, value):
    st.session_state.edited_content[index] = value
    st.session_state[f'row_{index+1}_locked'] = True

def unlock_row(index):
    st.session_state[f'row_{index+1}_locked'] = False

def generate_document():
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

        return buffer
    except Exception as e:
        st.error(f"Failed to generate Word document: {e}")
        return None

# Check if the password is correct
if password == "iken":
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
            if st.button("Edit Executive Summary (Double-click to confirm)"):
                unlock_summary()
        else:
            st.text_area("Edit the Executive Summary", 
                         key="new_summary",
                         value=st.session_state.summary,
                         height=200)
            if st.button("Save Executive Summary (Double-click to confirm)"):
                update_summary()

        # Display content of column A up to row 10 in editable fields
        st.write("Edit the content of column A (rows 1-10):")
        for i, content in enumerate(st.session_state.edited_content):
            if f'row_{i+1}_locked' not in st.session_state:
                st.session_state[f'row_{i+1}_locked'] = False

            if st.session_state[f'row_{i+1}_locked']:
                st.write(f"Row {i+1} (Locked):")
                st.write(content)
                if st.button(f"Edit Row {i+1} (Double-click to confirm)"):
                    unlock_row(i)
            else:
                edited_value = st.text_area(f"Row {i+1}", value=content, height=100, key=f"row_{i+1}")
                if st.button(f"Save Row {i+1} (Double-click to confirm)"):
                    update_row(i, edited_value)

        # Add a visual delimiter
        st.markdown("---")

        # Consolidated button for confirmation, generation, and download
        if st.button("Confirm, Generate, and Download Word Document"):
            with st.spinner("Generating document..."):
                doc_buffer = generate_document()
                if doc_buffer:
                    st.download_button(
                        label="Click here to download the generated document",
                        data=doc_buffer,
                        file_name="generated_report.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="download_button"
                    )
                    st.success("Document generated successfully! Click the download button above to save it.")
                else:
                    st.error("Failed to generate the document. Please try again.")

    else:
        st.warning("Please upload an Excel file to proceed.")
else:
    st.error("Incorrect password. Please enter the correct password to access the application.")