import os
import time
import pandas as pd
from pathlib import Path
from pydantic import BaseModel, Field
from typing import Literal
from google import genai
from google.genai import types

class ExtractionMatrix(BaseModel):
    final_inclusion_status: Literal["Include", "Exclude"] = Field(
        description="After reading the full text, decide if the paper truly meets ALL criteria: (1) Intermittent/delay-tolerant edge network, AND (2) Swarm Intelligence/Metaheuristic algorithm, AND (3) Routing/scheduling action. If it violates any criteria (e.g., assumes continuous connection, uses Deep Learning/DRL instead of Swarm), mark as Exclude."
    )
    exclusion_reason: str = Field(
        description="If 'Exclude', state the exact reason (e.g., 'Uses Deep Reinforcement Learning instead of Swarm Intelligence', 'Assumes continuous connectivity'). If 'Include', write 'N/A'."
    )
    citekey: str = Field(description="Generate a unique citation key using the first author's last name and the year of publication (e.g., 'Smith2025').")
    authors: str = Field(description="The full list of authors.")
    year: str = Field(description="The year of publication.")
    publication_type: str = Field(description="The type of publication (e.g., Journal, Conference Proceedings, Workshop).")
    doi: str = Field(description="The exact DOI if present in the text. If not, write 'Not specified'.")
    network_environment: str = Field(description="The specific intermittent environment (e.g., UWSN, FANET, sparse DTN, vehicular edge).")
    swarm_algorithm: str = Field(description="The base swarm algorithm used and any specific modifications.")
    agent_characteristics: str = Field(description="How the swarm agents operate (e.g., lightweight edge nodes vs. centralized master node).")
    operational_action: str = Field(description="What is the algorithm doing? (e.g., routing, task offloading, resource allocation).")
    handling_of_disconnections: str = Field(description="How does the system handle lost connections?")
    key_metrics_and_results: str = Field(description="The main performance metrics evaluated and claimed results.")
    limitations: str = Field(description="Any stated limitations, unrealistic assumptions, or methodology gaps.")

# Initialize the Gemini Client
client = genai.Client(api_key="<ADD API KEY HERE>")
PAPERS_DIR = Path("papers")

def process_pdf(pdf_path, max_retries=5):
    print(f"\nUploading: {pdf_path.name}...")
    try:
        uploaded_pdf = client.files.upload(file=pdf_path)
    except Exception as e:
        print(f"Failed to upload {pdf_path.name}: {e}")
        return None

    prompt = """
    You are an expert academic reviewer. Read this full-text paper and extract the required 
    information to populate my Systematic Literature Review extraction matrix. 
    
    CRITICAL STEP: First, verify if the paper actually meets the inclusion criteria upon full-text review.
    Many papers claim "edge" or "swarm" in the abstract but use Deep Learning or assume continuous 
    connectivity in the methodology. Evaluate this strictly.
    
    Extract the metadata (authors, year, DOI) and the technical matrix fields.
    Be concise but specific. If a detail is not mentioned, write 'Not specified'.
    """
    
    print(f"Extracting and verifying scope for {pdf_path.name}...")
    
    backoff_time = 5
    parsed_data = None
    
    for attempt in range(1, max_retries + 1):
        try:
            response = client.models.generate_content(
                model='gemini-3.7-flash',
                contents=[prompt, uploaded_pdf],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=ExtractionMatrix,
                    temperature=0.1 
                ),
            )
            parsed_data = response.parsed
            break  
        except Exception as e:
            print(f"  [Attempt {attempt}/{max_retries}] API warning: {e}")
            if attempt == max_retries:
                print(f"  [Failed] Max retries reached for {pdf_path.name}.")
                break
            print(f"  [Retrying] Waiting {backoff_time} seconds before trying again...")
            time.sleep(backoff_time)
            backoff_time *= 2
            
    # Clean up the file from Google's servers
    try:
        client.files.delete(name=uploaded_pdf.name)
    except Exception as e:
        print(f"  [Warning] Failed to delete file {uploaded_pdf.name}: {e}")
        
    return parsed_data

def main():
    results = []
    pdf_files = list(PAPERS_DIR.glob("*.pdf"))
    
    if not pdf_files:
        print(f"No PDFs found in the '{PAPERS_DIR}' directory.")
        return

    print(f"Found {len(pdf_files)} papers to extract. Starting processing...\n")

    for pdf_path in pdf_files:
        extracted_data = process_pdf(pdf_path)
        
        if extracted_data:
            row = extracted_data.model_dump()
            row["filename"] = pdf_path.name
            results.append(row)
        
        # A short 1-second pause
        time.sleep(1)

    print("\nFormatting and saving to Excel...")
    df = pd.DataFrame(results)
    
    # Reorder columns to put status and reasons first
    cols = ['filename', 'final_inclusion_status', 'exclusion_reason', 'citekey'] + \
           [col for col in df.columns if col not in ['filename', 'final_inclusion_status', 'exclusion_reason', 'citekey']]
    df = df[cols]
    
    output_file = 'SLR_Extraction_Matrix_Validated.xlsx'
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Extracted Data')
        worksheet = writer.sheets['Extracted Data']
        
        for col_idx, column in enumerate(worksheet.columns, 1):
            worksheet.column_dimensions[column[0].column_letter].width = 30
            for cell in column:
                cell.alignment = cell.alignment.copy(wrapText=True, vertical='top')
                
    print(f"\nSuccess! Open '{output_file}' to see your populated extraction matrix.")

if __name__ == "__main__":
    main()