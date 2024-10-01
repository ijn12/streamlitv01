import streamlit as st
import pandas as pd
import io
from docx import Document

# Step 1: Upload XLSX file
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Step 2: Read Excel file into DataFrame
    df = pd.read_excel(uploaded_file, usecols="A", nrows=10)
    
    # Step 3: Display content of column A up to row 10 in editable fields
    st.write("Edit the content of column A (rows 1-10):")
    
    edited_content = []
    for i in range(len(df)):
        edited_content.append(st.text_input(f"Row {i+1}", value=df.iloc[i, 0]))

    # Step 4: Confirm and generate DOCX file
    if st.button("Confirm and Generate Word File"):
        # Step 5: Create a Word document and add the content
        doc = Document()
        doc.add_heading("Generated Content", level=1)
        for i, content in enumerate(edited_content):
            doc.add_paragraph(f"Row {i+1}: {content}")
        
        # Save the document to a BytesIO buffer
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        # Step 6: Provide the Word file for download
        st.download_button(
            label="Download Word File",
            data=buffer,
            file_name="generated_report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
