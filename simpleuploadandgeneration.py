import streamlit as st
import pandas as pd
import io
from docx import Document
from openai import OpenAI
from PyPDF2 import PdfReader, PdfWriter  # Updated import
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# Testing header to see if the deployment works
st.header("Testing Deployment 2")

# Sidebar inputs for OpenAI API key and file upload
st.sidebar.header("Configuration")
api_key = st.sidebar.text_input("Enter your OpenAI API Key", type="password")
uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=["xlsx"])

if api_key and uploaded_file is not None:
    # Initialize the OpenAI client once the API key is provided
    client = OpenAI(api_key=api_key)

    # Step 3: Read Excel file into DataFrame (starting from the first row)
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
    summary = ""
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
                st.session_state['summary'] = summary  # Store summary in session state
                st.subheader("Executive Summary:")
                st.write(summary)
            except Exception as e:
                st.error(f"Failed to generate summary: {e}")

    # Editable field for the executive summary
    if 'summary' in st.session_state:
        summary = st.text_area("Edit the Executive Summary", value=st.session_state['summary'])
    else:
        summary = st.text_area("Edit the Executive Summary", value="")

    # Step 7: Confirm and generate PDF file
    if st.button("Confirm and Generate PDF"):
        # Create a PDF document and add the content
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=letter)
        width, height = letter

        # Add title
        c.setFont("Helvetica-Bold", 16)
        c.drawString(50, height - 50, "Generated Content")

        # Add Executive Summary
        if summary:
            c.setFont("Helvetica-Bold", 14)
            c.drawString(50, height - 80, "Executive Summary")
            c.setFont("Helvetica", 12)
            text_object = c.beginText(50, height - 100)
            for line in summary.split('\n'):
                text_object.textLine(line)
            c.drawText(text_object)

        # Add edited content
        y_position = height - 200
        for i, content in enumerate(edited_content):
            if y_position < 50:
                c.showPage()
                y_position = height - 50
            c.setFont("Helvetica", 12)
            c.drawString(50, y_position, f"Row {i+1}: {content}")
            y_position -= 20

        c.save()
        buffer.seek(0)

        # Display the PDF
        st.subheader("Generated PDF Preview:")
        st.write(buffer)

        # Provide the PDF file for download
        st.download_button(
            label="Download PDF File",
            data=buffer,
            file_name="generated_report.pdf",
            mime="application/pdf"
        )

    # Add a PDF viewer
    if 'pdf_buffer' in st.session_state:
        st.subheader("PDF Preview:")
        st.write(st.session_state['pdf_buffer'])
