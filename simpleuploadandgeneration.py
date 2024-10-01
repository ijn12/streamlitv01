import streamlit as st
import pandas as pd
import io
from docx import Document
from openai import OpenAI

# Step 1: Input for OpenAI API key
api_key = st.text_input("Enter your OpenAI API Key", type="password")

if api_key:
    # Initialize the OpenAI client once the API key is provided
    client = OpenAI(api_key=api_key)

    # Step 2: Upload XLSX file
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

    if uploaded_file is not None:
        # Step 3: Read Excel file into DataFrame
        df = pd.read_excel(uploaded_file, usecols="A", nrows=10)

        # Step 4: Custom prompt input for the executive summary
        custom_prompt = st.text_area("Enter your custom prompt for the executive summary")

        # Step 5: Display content of column A up to row 10 in editable fields
        st.write("Edit the content of column A (rows 1-10):")

        edited_content = []
        for i in range(len(df)):
            edited_content.append(st.text_input(f"Row {i+1}", value=df.iloc[i, 0]))

        # Combine the content of all rows to send to GPT API
        combined_content = "\n".join(edited_content)

        # Step 6: Single button for generating or refreshing the executive summary
        if st.button("Generate/Refresh Executive Summary"):
            with st.spinner("Generating executive summary..."):
                try:
                    # Combine the custom prompt and the content of the rows
                    final_prompt = f"{custom_prompt}\n{combined_content}"
                    response = client.chat.completions.create(
                        model="gpt-4o-mini",  # Updated model
                        messages=[
                            {
                                "role": "user",
                                "content": final_prompt
                            }
                        ],
                        temperature=0.7
                    )
                    # Extract the summary from the API response
                    summary = response.choices[0].message.content
                    st.subheader("Executive Summary:")
                    st.write(summary)
                except Exception as e:
                    st.error(f"Failed to generate summary: {e}")

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
