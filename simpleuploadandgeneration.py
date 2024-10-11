import streamlit as st
import pandas as pd
import io
from openai import OpenAI
from docx import Document
import base64

# Testing header to see if the deployment works
st.header("Testing Deployment 5")

# Sidebar inputs for OpenAI API key and file upload
st.sidebar.header("Configuration")
api_key = st.secrets["openai_api_key"]
uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=["xlsx"])

# Hardcoded template path
template_path = "template.docx"

if uploaded_file is not None:
    # Initialize the OpenAI client once the API key is provided
    client = OpenAI(api_key=api_key)

    # Read Excel file into DataFrame (starting from the first row)
    df = pd.read_excel(uploaded_file, usecols="A", nrows=10)

    # Hardcoded prompt for executive summary
    custom_prompt = "Generate an executive summary based on the following content:"

    # Single button for generating or refreshing the executive summary
    if st.button("Generate/Refresh Executive Summary"):
        with st.spinner("Generating executive summary..."):
            try:
                combined_content = "\n".join(df.iloc[:, 0].astype(str))
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
                summary = response.choices[0].message.content
                st.session_state['summary'] = summary
                st.session_state['summary_locked'] = False
                st.subheader("Executive Summary:")
                st.write(summary)
            except Exception as e:
                st.error(f"Failed to generate summary: {e}")

    # Editable field for the executive summary with save and edit functionality
    if 'summary_locked' not in st.session_state:
        st.session_state['summary_locked'] = False

    if st.session_state['summary_locked']:
        st.subheader("Executive Summary (Locked):")
        st.write(st.session_state.get('summary', ''))
        if st.button("Edit Executive Summary"):
            st.session_state['summary_locked'] = False
    else:
        summary = st.text_area("Edit the Executive Summary", 
                               value=st.session_state.get('summary', ''),
                               height=200)
        if st.button("Save Executive Summary"):
            st.session_state['summary'] = summary
            st.session_state['summary_locked'] = True

    # Display content of column A up to row 10 in editable fields with save and edit functionality
    st.write("Edit the content of column A (rows 1-10):")
    edited_content = []
    for i in range(len(df)):
        if f'row_{i+1}_locked' not in st.session_state:
            st.session_state[f'row_{i+1}_locked'] = False

        if st.session_state[f'row_{i+1}_locked']:
            st.write(f"Row {i+1} (Locked):")
            st.write(st.session_state.get(f'row_{i+1}', df.iloc[i, 0]))
            if st.button(f"Edit Row {i+1}"):
                st.session_state[f'row_{i+1}_locked'] = False
                st.experimental_rerun()
        else:
            edited_value = st.text_area(f"Row {i+1}", value=st.session_state.get(f'row_{i+1}', df.iloc[i, 0]), height=100)
            if st.button(f"Save Row {i+1}"):
                st.session_state[f'row_{i+1}'] = edited_value
                st.session_state[f'row_{i+1}_locked'] = True
                st.experimental_rerun()
            edited_content.append(st.session_state.get(f'row_{i+1}', edited_value))

    # Confirm and generate Word document
    summary = st.session_state.get('summary', '')
    if st.button("Confirm and Generate Word Document"):
        try:
            # Load the Word template
            template = Document(template_path)

            # Replace placeholders in the template
            for paragraph in template.paragraphs:
                if "{{Executive_Summary}}" in paragraph.text:
                    paragraph.text = paragraph.text.replace("{{Executive_Summary}}", summary)

                for i, content in enumerate(edited_content):
                    placeholder = f"{{{{Row_{i+1}}}}}"
                    if placeholder in paragraph.text:
                        paragraph.text = paragraph.text.replace(placeholder, content)

            # Save the generated document to a buffer
            buffer = io.BytesIO()
            template.save(buffer)
            buffer.seek(0)

            # Display the Word document
            st.subheader("Generated Word Document Preview:")
            st.write("Word document generated successfully. Use the download button below to view or save the document.")

            # Provide the Word document for download
            st.download_button(
                label="Download Word Document",
                data=buffer,
                file_name="generated_report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"Failed to generate Word document: {e}")