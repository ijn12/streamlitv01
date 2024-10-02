import streamlit as st
import pandas as pd
import io
from openai import OpenAI
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import base64

# Testing header to see if the deployment works
st.header("Testing Deployment 3")

# Sidebar inputs for OpenAI API key and file upload
st.sidebar.header("Configuration")
api_key = st.sidebar.text_input("Enter your OpenAI API Key", type="password")
uploaded_file = st.sidebar.file_uploader("Upload your Excel file", type=["xlsx"])

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

    # Confirm and generate PDF file
    if st.button("Confirm and Generate PDF"):
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
            c.drawString(50, y_position, f"Row {i+1}:")
            text_object = c.beginText(70, y_position - 15)
            for line in content.split('\n'):
                text_object.textLine(line)
            c.drawText(text_object)
            y_position -= 20 + (len(content.split('\n')) * 15)

        c.save()
        buffer.seek(0)

        # Display the PDF
        st.subheader("Generated PDF Preview:")
        st.write("PDF generated successfully. Use the download button below to view or save the PDF.")

        # Provide the PDF file for download
        st.download_button(
            label="Download PDF File",
            data=buffer,
            file_name="generated_report.pdf",
            mime="application/pdf"
        )

        # Display PDF preview using base64 encoding
        base64_pdf = base64.b64encode(buffer.getvalue()).decode('utf-8')
        pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="700" height="1000" type="application/pdf"></iframe>'
        st.markdown(pdf_display, unsafe_allow_html=True)
