import re
import pandas as pd
import pdfplumber
import streamlit as st
from io import BytesIO

def extract_pdf_text(file_stream):
    """
    Extracts text from the given PDF file stream using pdfplumber.
    """
    all_text = ""
    with pdfplumber.open(file_stream) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text += text + "\n"
    return all_text

def parse_proposals(text):
    """
    Parses the AGM proposal data from the extracted text.
    
    Expected fields for each proposal:
      - Proposal Proxy Year
      - Resolution Outcome (Approved if For votes > Against votes)
      - Proposal Text (extracted text enclosed in double quotes)
      - Mgmt. Proposal Category (left blank)
      - Vote Results - For
      - Vote Results - Against
      - Vote Results - Abstained
      - Vote Results - Withheld
      - Vote Results - Broker Non-Votes (if 'Nil' or '-' then consider as zero)
      - Proposal Vote Results Total (left blank)
    
    Adjust regex patterns as needed to match your PDF's structure.
    """
    proposals = []
    # Assuming each proposal starts with "Proposal Proxy Year:" (case-insensitive)
    proposal_blocks = re.split(r'Proposal\s+Proxy\s+Year:', text, flags=re.IGNORECASE)
    
    for block in proposal_blocks[1:]:
        proposal = {}
        # Extract Proposal Proxy Year (assuming a 4-digit year at the beginning)
        m_year = re.match(r'\s*(\d{4})', block)
        proposal['Proposal Proxy Year'] = m_year.group(1) if m_year else ""
        
        # Extract Vote Results - For
        m_for = re.search(r'For\s*votes\s*[:\-]?\s*([\d,]+)', block, flags=re.IGNORECASE)
        proposal['Vote Results - For'] = m_for.group(1) if m_for else ""
        
        # Extract Vote Results - Against
        m_against = re.search(r'Against\s*votes\s*[:\-]?\s*([\d,]+)', block, flags=re.IGNORECASE)
        proposal['Vote Results - Against'] = m_against.group(1) if m_against else ""
        
        # Extract Vote Results - Abstained
        m_abstained = re.search(r'Abstained\s*votes\s*[:\-]?\s*([\d,]+)', block, flags=re.IGNORECASE)
        proposal['Vote Results - Abstained'] = m_abstained.group(1) if m_abstained else ""
        
        # Extract Vote Results - Withheld
        m_withheld = re.search(r'Withheld\s*votes\s*[:\-]?\s*([\d,]+)', block, flags=re.IGNORECASE)
        proposal['Vote Results - Withheld'] = m_withheld.group(1) if m_withheld else ""
        
        # Extract Vote Results - Broker Non-Votes (treat "Nil" or "-" as zero)
        m_broker = re.search(r'Broker\s*Non[-\s]*Votes\s*[:\-]?\s*([\d,]+|Nil|-)', block, flags=re.IGNORECASE)
        if m_broker:
            val = m_broker.group(1)
            proposal['Vote Results - Broker Non-Votes'] = "0" if val in ["Nil", "-"] else val
        else:
            proposal['Vote Results - Broker Non-Votes'] = ""
        
        # Extract Proposal Text (assuming it is enclosed in double quotes)
        m_text = re.search(r'Proposal\s*Text\s*[:\-]?\s*"([^"]+)"', block, flags=re.IGNORECASE)
        proposal['Proposal Text'] = m_text.group(1) if m_text else ""
        
        # Calculate Resolution Outcome: Approved if For votes > Against votes.
        try:
            for_votes = int(proposal['Vote Results - For'].replace(",", "")) if proposal['Vote Results - For'] else 0
        except Exception:
            for_votes = 0
        try:
            against_votes = int(proposal['Vote Results - Against'].replace(",", "")) if proposal['Vote Results - Against'] else 0
        except Exception:
            against_votes = 0
        if for_votes > against_votes:
            proposal['Resolution Outcome'] = f"Approved ({proposal['Vote Results - For']} For > {proposal['Vote Results - Against']} Against)"
        else:
            proposal['Resolution Outcome'] = ""
        
        # Mgmt. Proposal Category (left blank)
        proposal['Mgmt. Proposal Category'] = ""
        # Proposal Vote Results Total (left blank)
        proposal['Proposal Vote Results Total'] = ""
        
        proposals.append(proposal)
    return proposals

def parse_directors(text):
    """
    Parses the director election results from the extracted text.
    
    Expected fields for each director:
      - Director Election Year (fixed as 2024)
      - Individual (director name)
      - Director Votes For
      - Director Votes Against (blank if not available)
      - Director Votes Abstained (blank if not available)
      - Director Votes Withheld
      - Director Votes Broker-Non-Votes (if 'Nil' or '-' then consider as zero)
    
    Adjust regex patterns as needed to match your PDF's structure.
    """
    directors = []
    # Assuming each director block starts with "Individual:" (case-insensitive)
    director_blocks = re.split(r'Individual:', text, flags=re.IGNORECASE)
    
    for block in director_blocks[1:]:
        director = {}
        director['Director Election Year'] = "2024"
        
        # Extract the director's name (up to the first newline)
        m_name = re.match(r'\s*([^\n]+)', block)
        director['Individual'] = m_name.group(1).strip() if m_name else ""
        
        # Extract Director Votes For
        m_for = re.search(r'Director\s*Votes\s*For\s*[:\-]?\s*([\d,]+)', block, flags=re.IGNORECASE)
        director['Director Votes For'] = m_for.group(1) if m_for else ""
        
        # Extract Director Votes Against
        m_against = re.search(r'Director\s*Votes\s*Against\s*[:\-]?\s*([\d,]+)', block, flags=re.IGNORECASE)
        director['Director Votes Against'] = m_against.group(1) if m_against else ""
        
        # Extract Director Votes Abstained
        m_abstained = re.search(r'Director\s*Votes\s*Abstained\s*[:\-]?\s*([\d,]+)', block, flags=re.IGNORECASE)
        director['Director Votes Abstained'] = m_abstained.group(1) if m_abstained else ""
        
        # Extract Director Votes Withheld
        m_withheld = re.search(r'Director\s*Votes\s*Withheld\s*[:\-]?\s*([\d,]+)', block, flags=re.IGNORECASE)
        director['Director Votes Withheld'] = m_withheld.group(1) if m_withheld else ""
        
        # Extract Director Votes Broker-Non-Votes (treat "Nil" or "-" as zero)
        m_broker = re.search(r'Director\s*Votes\s*Broker[-\s]*Non[-\s]*Votes\s*[:\-]?\s*([\d,]+|Nil|-)', block, flags=re.IGNORECASE)
        if m_broker:
            val = m_broker.group(1)
            director['Director Votes Broker-Non-Votes'] = "0" if val in ["Nil", "-"] else val
        else:
            director['Director Votes Broker-Non-Votes'] = ""
        
        directors.append(director)
    return directors

def save_to_excel(proposals, directors):
    """
    Saves the proposals and director election data to an Excel file with two sheets.
    Returns a BytesIO stream containing the Excel file.
    """
    proposals_df = pd.DataFrame(proposals, columns=[
        'Proposal Proxy Year',
        'Resolution Outcome',
        'Proposal Text',
        'Mgmt. Proposal Category',
        'Vote Results - For',
        'Vote Results - Against',
        'Vote Results - Abstained',
        'Vote Results - Withheld',
        'Vote Results - Broker Non-Votes',
        'Proposal Vote Results Total'
    ])
    
    directors_df = pd.DataFrame(directors, columns=[
        'Director Election Year',
        'Individual',
        'Director Votes For',
        'Director Votes Against',
        'Director Votes Abstained',
        'Director Votes Withheld',
        'Director Votes Broker-Non-Votes'
    ])
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        proposals_df.to_excel(writer, sheet_name='Proposal Sheet', index=False)
        directors_df.to_excel(writer, sheet_name='Non-Proposal Sheet', index=False)
    output.seek(0)
    return output

# Streamlit interface
st.title("AGM Data Extractor")

st.markdown("""
Upload your AGM PDF result document. The script will extract proposal and director election data and generate an Excel file with two sheets.
""")

uploaded_file = st.file_uploader("Upload PDF file", type="pdf")

if uploaded_file is not None:
    st.info("Processing the PDF file...")
    # Extract text from the uploaded PDF file
    pdf_text = extract_pdf_text(uploaded_file)
    
    # Parse proposals and directors from the extracted text
    proposals = parse_proposals(pdf_text)
    directors = parse_directors(pdf_text)
    
    if not proposals:
        st.warning("No proposals were found in the document.")
    if not directors:
        st.warning("No director election results were found in the document.")
    
    # Save the data to an Excel file in memory
    excel_data = save_to_excel(proposals, directors)
    
    st.success("Data extraction complete!")
    st.download_button(
        label="Download Excel File",
        data=excel_data,
        file_name="agm_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
