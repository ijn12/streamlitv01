import streamlit as st
import pandas as pd
import io
from docx import Document
import openai

# Step 1: Input for OpenAI API key
api_key = st.text_input("Enter your OpenAI API Key", type="password")

if api_key:
    # Step 2: Upload XLSX file
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

    if uploaded_file is not None:
        # Step 3: Read Excel file into DataFrame
        df = pd.read_excel(uploaded_file, usecols="A", nrows=10)

        # Step 4: Display content of column A up to row 10 in editable fields
        st.write("Edit the content of column A (rows 1-10):")

        edited_content = []
        for i in range(len(df)):
            edited_content.append(st.text_input(f"Row {i+1}", value=df.iloc[i, 0]))

        # Combine the content of all rows to send to GPT API
        combined_content = "\n".join(edited_content)

        # Step 5: Generate Executive Summary with GPT-4
        if st.button("Generate Executive Summary"):
            with st.spinner("Generating executive summary..."):
                try:
                    openai.api_key = api_key
                    response = openai.completions.create(  # Updated method
                        model="gpt-4o-mini",
                        messages=[
                            {
                                "role": "user",
                                "content": f"Please generate an executive summary for the following content:\n{combined_content}"
                            }
                        ],
                        temperature=0.7
                    )
                    # Extract the summary from the API response
                    summary = response["choices"][0]["message"]["content"]
                    st.subheader("Executive Summary:")
                    st.write(summary)
                except Exception as e:
                    st.error(f"Failed to generate summary: {e}")

        # Step 6: Provide a refresh button for regenerating the summary after edits
        if st.button("Refresh Executive Summary"):
            with st.spinner("Regenerating executive summary..."):
                try:
                    openai.api_key = api_key
                    response = openai.completions.create(  # Updated method
                        model="gpt-4o-mini",
                        messages=[
                            {
                                "role": "user",
                                "content": f"Please generate an updated executive summary for the following updated content:\n{combined_content}"
                            }
                        ],
                        temperature=0.7
                    )
                    # Extract the updated summary from the API response
                    updated_summary = response["choices"][0]["message"]["content"]
                    st.subheader("Updated Executive Summary:")
                    st.write(updated_summary)
                except Exception as e:
                    st.error(f"Failed to regenerate summary: {e}")

        # Step 7: Confirm and generate DOCX file
        if st.button("Confirm and Generate Word File"):
            # Create a Word document and add the content
            doc = Document()
            doc.add_heading("Generated Content", level=1)
            for i, content in enumerate(edited_content):
                doc.add_paragraph(f"Row {i+1}: {content}")

            # Add the Executive Summary if it exists
            if 'summary' in locals():
                doc.add_heading("Executive Summary", level=2)
                doc.add_paragraph(summary)
            if 'updated_summary' in locals():
                doc.add_heading("Updated Executive Summary", level=2)
                doc.add_paragraph(updated_summary)

            # Save the document to a BytesIO buffer
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            # Provide the Word file for download
            st.download_button(
                label="Download Word File",
                data=buffer,
                file_name="generated_report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
