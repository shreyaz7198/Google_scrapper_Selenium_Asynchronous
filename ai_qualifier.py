import os
import sys
import pandas as pd
from google import genai

def ask_ai_agent_to_qualify(business_name, website, phone):
    """An AI core routine that scores lead value using the Gemini API."""
    if not os.environ.get("GEMINI_API_KEY"):
        print("❌ Error: GEMINI_API_KEY is missing from your system environment settings!")
        sys.exit(1)
        
    client = genai.Client()
    
    prompt = f"""
    You are an expert B2B Sales Growth Agent. Analyze this scraped business row:
    - Business Name: {business_name}
    - Website Listed: {website if website and str(website) != 'nan' else 'NONE'}
    - Phone Number Listed: {phone if phone and str(phone) != 'nan' else 'NONE'}
    
    Determine if this business is a prime target for a marketing agency.
    Rules:
    1. Classify as '🔥 HOT TARGET' if they lack a website OR lack a phone line.
    2. Classify as '💤 LOW PRIORITY' if they already have both listed perfectly.
    
    Respond exactly in this layout format:
    VERDICT: [🔥 HOT TARGET or 💤 LOW PRIORITY]
    REASON: [Write a clear, one-sentence explanation of why]
    """
    
    try:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
        )
        return response.text.strip()
    except Exception as e:
        return f"AI Agent Error: {str(e)}"

def run_lead_qualification_pipeline():
    excel_path = os.path.expanduser("~/Downloads/maps_scraped_leads.xlsx")
    
    if not os.path.exists(excel_path):
        print(f"❌ File Error: Could not find your spreadsheet at: {excel_path}")
        print("Please run your scraper first to create the data file!")
        return
        
    print("🤖 AI Qualification Agent active. Reading spreadsheet...")
    df = pd.read_excel(excel_path)
    
    if "Name" not in df.columns:
        print("❌ Formatting Error: Missing the 'Name' column header inside the target spreadsheet template.")
        return
        
    ai_verdicts = []
    
    print("🧠 Auditing the first 3 lead entries as an optimization sequence...")
    for index, row in df.head(3).iterrows():
        name = row.get("Name", "Unknown")
        website = row.get("Website", "")
        phone = row.get("Contact Number", "")
        
        print(f"   👉 AI Agent is checking: '{name}'")
        analysis_result = ask_ai_agent_to_qualify(name, website, phone)
        ai_verdicts.append(analysis_result)
        
    while len(ai_verdicts) < len(df):
        ai_verdicts.append("Not Evaluated")
        
    df["AI Sales Analysis"] = ai_verdicts
    
    output_filename = excel_path.replace(".xlsx", "_qualified_by_ai.xlsx")
    df.to_excel(output_filename, index=False)
    print(f"\n🎉 Task Complete! Smart report saved successfully at:\n👉 {output_filename}")

if __name__ == "__main__":
    run_lead_qualification_pipeline()