import streamlit as st
import pandas as pd
import io
from openai import OpenAI
from docx import Document
import base64

# Testing header to see if the deployment works
st.header("Testing Deployment 3")

# Sidebar inputs for OpenAI API key and file upload
st.sidebar.header("Configuration")
api_key = st.sidebar.text_input("Enter your OpenAI API Key", type="password")
uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=["xlsx"])

# Hardcoded template path
template_path = "template.docx"

if api_key and uploaded_file is not None:
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
                st.subheader("Executive Summary:")
                st.write(summary)
            except Exception as e:
                st.error(f"Failed to generate summary: {e}")

    # Editable field for the executive summary
    summary = st.text_area("Edit the Executive Summary", 
                           value=st.session_state.get('summary', ''),
                           height=200)

    # Display content of column A up to row 10 in editable fields
    st.write("Edit the content of column A (rows 1-10):")
    edited_content = []
    for i in range(len(df)):
        edited_content.append(st.text_area(f"Row {i+1}", value=df.iloc[i, 0], height=100))

    # Confirm and generate Word document
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