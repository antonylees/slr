import time
import rispy
import pandas as pd
from pydantic import BaseModel, Field
from typing import Literal
from google import genai
from google.genai import types

class ScreeningDecision(BaseModel):
    status: Literal["Include", "Exclude"] = Field(
        description="The final decision to include or exclude the paper."
    )
    reason_category: str = Field(
        description="Short category, e.g., 'Wrong Mechanism', 'Continuous Connectivity', 'Review Paper', 'Out of Scope', or 'Matches Scope'"
    )
    justification: str = Field(
        description="A single concise sentence justifying the decision based on the criteria."
    )

# Initialize the Gemini Client with your explicit key
client = genai.Client(api_key="<ADD API KEY HERE")

def screen_paper_with_retry(title, abstract, max_retries=5):
    """Sends prompt to Gemini with built-in exponential backoff for 503/rate-limit errors."""
    prompt = f"""
    You are an expert academic researcher performing a Systematic Literature Review.
    Evaluate the following paper for INCLUSION or EXCLUSION based strictly on the criteria below.

    INCLUSION CRITERIA:
    1. Environment: Intermittent, sparse, disruption-tolerant, or mobile networks with unpredictable availability/disconnections (e.g., DTNs, UWSNs, FANETs, UAV swarms).
    2. Mechanism: Decentralized bio-inspired swarm intelligence or metaheuristic optimization algorithms (ACO, PSO, GWO, ABC, AFSA, etc.).
    3. Action: Network-layer routing, data forwarding, or edge task scheduling/compute offloading under intermittent connectivity.

    EXCLUSION CRITERIA (Reject if any apply):
    - Continuous Connectivity: Assumes stable/continuous connection (standard MANETs, connected VANETs, connected IoT/IIoT edge, stationary WSNs).
    - Wrong Mechanism: Uses deep learning, deep reinforcement learning (DRL/DQN), container orchestration (Docker Swarm), game theory, or standard mathematical heuristics rather than a swarm intelligence metaheuristic.
    - Out of Scope Action: Focuses purely on physical UAV trajectory/flight planning, physical smart grid power scheduling, or cache deployment without routing/task offloading.
    - Review/Survey Papers: Exclude any surveys, literature reviews, or tutorials.

    PAPER DETAILS:
    Title: {title}
    Abstract: {abstract}
    """
    
    backoff_time = 5
    for attempt in range(1, max_retries + 1):
        try:
            response = client.models.generate_content(
                model='gemini-3.7-flash',  # Stable current workhorse model
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=ScreeningDecision,
                    temperature=0.1
                ),
            )
            return response.parsed
        except Exception as e:
            print(f"  [Attempt {attempt}/{max_retries}] API warning on '{title[:25]}...': {e}")
            if attempt == max_retries:
                print(f"  [Failed] Max retries reached for this paper.")
                return None
            print(f"  [Retrying] Waiting {backoff_time} seconds before trying again...")
            time.sleep(backoff_time)
            backoff_time *= 2  # Double the wait time on each subsequent failure (exponential backoff)

def main():
    input_file = 'My Library.ris'
    output_file = 'SLR_Screening_Results.xlsx'
    
    # Parse the RIS file
    with open(input_file, 'r', encoding='utf-8') as f:
        entries = list(rispy.load(f))
        
    results = []
    total = len(entries)
    print(f"Loaded {total} records. Starting resilient screening...\n")
    
    for i, entry in enumerate(entries, 1):
        title = entry.get('title', 'Unknown Title')
        abstract = entry.get('abstract', 'No abstract available.')
        
        # Extract authors and DOI
        authors_list = entry.get('authors', [])
        authors = ", ".join(authors_list) if authors_list else "Unknown Author"
        doi = entry.get('doi', 'No DOI')
        
        print(f"[{i}/{total}] Screening: {title[:60]}...")
        
        if abstract == 'No abstract available.':
            results.append({
                "Title": title,
                "Authors": authors,
                "DOI": doi,
                "Status": "Exclude",
                "Reason Category": "No Abstract",
                "Justification": "Cannot screen without an abstract.",
                "Abstract": abstract
            })
            continue

        decision = screen_paper_with_retry(title, abstract)
        
        if decision:
            results.append({
                "Title": title,
                "Authors": authors,
                "DOI": doi,
                "Status": decision.status,
                "Reason Category": decision.reason_category,
                "Justification": decision.justification,
                "Abstract": abstract
            })
        else:
            results.append({
                "Title": title,
                "Authors": authors,
                "DOI": doi,
                "Status": "Error",
                "Reason Category": "API Error",
                "Justification": "Failed to retrieve decision after multiple retries.",
                "Abstract": abstract
            })
            
        # Standard pacing between successful requests (6 seconds for free tier safety)
        # time.sleep(6) 

    # Export to a formatted Excel file
    print("\nFormatting and saving to Excel...")
    df = pd.DataFrame(results)
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Screening')
        worksheet = writer.sheets['Screening']
        
        # Adjust column widths for easy reading
        worksheet.column_dimensions['A'].width = 40  # Title
        worksheet.column_dimensions['B'].width = 25  # Authors
        worksheet.column_dimensions['C'].width = 25  # DOI
        worksheet.column_dimensions['D'].width = 15  # Status
        worksheet.column_dimensions['E'].width = 25  # Category
        worksheet.column_dimensions['F'].width = 50  # Justification
        worksheet.column_dimensions['G'].width = 100 # Abstract
        
        # Enable text wrapping on all cells
        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = cell.alignment.copy(wrapText=True)
                
    print(f"\nSuccess! Open '{output_file}' to review your full dataset.")

if __name__ == "__main__":
    main()