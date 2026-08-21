from google import genai
import json
import traceback
import os

def extract_deeds_from_transfer_sheet(api_key, pdf_path):
    if not api_key:
        return [], "No API key provided."
        
    if not os.path.exists(pdf_path):
        return [], f"File not found: {pdf_path}"

    try:
        client = genai.Client(api_key=api_key)
        
        sample_file = client.files.upload(
            file=pdf_path, 
            config={'display_name': 'Transfer Sheet'}
        )
        
        prompt = """
        This is a property transfer sheet or deed card. It contains a list of historical deeds and ownership transfers.
        1. Please extract all Deed/Mortgage references. Look for 'Vol' or 'Volume' and 'Pg' or 'Page'. Also extract the 'Date'.
        2. Look in the top right corner for "OUT LOT NO." or "LOT NO.". Extract the lot number as a string (e.g. "Lot 142" or "Out Lot 5").
        Return the data in a strict JSON object format with two keys:
        {
           "deeds": [{"vol": "...", "pg": "...", "date": "..."}],
           "lot": "Lot 142"
        }
        """
        
        response = client.models.generate_content(
            model='gemini-3.6-flash',
            contents=[sample_file, prompt],
            config={
                'response_mime_type': 'application/json',
            }
        )
        
        # Cleanup file from Google's servers
        try:
            client.files.delete(name=sample_file.name)
        except:
            pass
            
        text = response.text.strip()
        
        data = json.loads(text)
        return data, None
        
    except Exception as e:
        error_details = traceback.format_exc()
        return [], f"Error: {str(e)}\n{error_details}"

def generate_abstracting_summary(api_key, pdf_path):
    if not api_key:
        return "No API key provided."
        
    if not os.path.exists(pdf_path):
        return f"File not found: {pdf_path}"

    try:
        client = genai.Client(api_key=api_key)
        
        sample_file = client.files.upload(
            file=pdf_path, 
            config={'display_name': 'Transfer Sheet'}
        )
        
        prompt = """
        This is a property transfer sheet or deed card. It contains a list of historical deeds and ownership transfers.
        Create a concise abstracting summary of the chain of title based on the document, in this format (or similar):
        [Book Type] [Vol]/[Page]: [Grantor] to [Grantee], Conveyance: [Type], Exceptions: [Flag any O&G / Oil & Gas exceptions if listed].
        
        Return ONLY the multi-line text summary. Do not format as JSON.
        """
        
        response = client.models.generate_content(
            model='gemini-3.6-flash',
            contents=[sample_file, prompt]
        )
        
        # Cleanup file from Google's servers
        try:
            client.files.delete(name=sample_file.name)
        except:
            pass
            
        return response.text.strip()
        
    except Exception as e:
        error_details = traceback.format_exc()
        return f"Error generating summary: {str(e)}\n{error_details}"

def generate_abstracting_summary_multiple(api_key, pdf_paths):
    if not api_key:
        return "No API key provided."
        
    valid_paths = [p for p in pdf_paths if os.path.exists(p)]
    if not valid_paths:
        return "No valid files found."

    client = genai.Client(api_key=api_key)
    uploaded_files = []
    
    try:
        for idx, pdf_path in enumerate(valid_paths):
            f = client.files.upload(
                file=pdf_path, 
                config={'display_name': f'Deed Document {idx}'}
            )
            uploaded_files.append(f)
        
        prompt = """
        These are property transfer sheets or deed documents containing historical deeds and ownership transfers.
        Create a concise abstracting summary of the chain of title based on ALL these documents. 
        For each document, output a note exactly in this format:
        [Date] | [Book Type] [Vol]/[Page] - [Last Name First Name] (Grantor) .... [Last Name First Name] (Grantee)
        Also flag any O&G / Oil & Gas exceptions if listed.
        
        Return ONLY the multi-line text summary. Do not format as JSON. Ensure you read through all provided documents.
        """
        
        contents = uploaded_files + [prompt]
        response = client.models.generate_content(
            model='gemini-3.6-flash',
            contents=contents
        )
        
        text_result = response.text.strip()
        
    except Exception as e:
        error_details = traceback.format_exc()
        text_result = f"Error generating summary: {str(e)}\n{error_details}"
        
    finally:
        # Cleanup files from Google's servers
        for uf in uploaded_files:
            try:
                client.files.delete(name=uf.name)
            except:
                pass
                
    return text_result

def generate_kofile_search_params(api_key, pdf_paths):
    if not api_key:
        return []
        
    valid_paths = [p for p in pdf_paths if os.path.exists(p)]
    if not valid_paths:
        return []

    client = genai.Client(api_key=api_key)
    uploaded_files = []
    
    try:
        for idx, pdf_path in enumerate(valid_paths):
            f = client.files.upload(
                file=pdf_path, 
                config={'display_name': f'Deed Document {idx}'}
            )
            uploaded_files.append(f)
        
        prompt = """
        These are multiple historical property transfer deeds forming a continuous chain of title.
        You MUST extract a comprehensive list of EVERY SINGLE individual or entity who ever owned this property across ALL provided documents.
        There will likely be multiple different owners (Grantees who later become Grantors) across the different documents. Do not miss any of them!
        
        For each person, determine the exact date they acquired the property (when they were the Grantee) and the exact date they transferred it away (when they became the Grantor).
        
        Output MUST be in strict JSON format matching this schema:
        {
            "owners": [
                {
                    "name": "Last First", // Must be Title Case. Omit middle initials! e.g., 'Vanfossen Timothy', NOT 'Vanfossen Timothy N' or 'VANFOSSEN TIMOTHY N'.
                    "acquisition_date": "MM/DD/YYYY", // Or null if unknown. If only year is known, use 01/01/YYYY.
                    "disposal_date": "MM/DD/YYYY" // Or null if unknown. If only year is known, use 12/31/YYYY.
                }
            ]
        }
        
        Output ONLY the raw JSON block without markdown formatting or code blocks.
        """
        
        contents = uploaded_files + [prompt]
        response = client.models.generate_content(
            model='gemini-3.6-flash',
            contents=contents
        )
        
        text_result = response.text.strip()
        if text_result.startswith('```json'):
            text_result = text_result[7:]
        if text_result.endswith('```'):
            text_result = text_result[:-3]
            
        import json
        data = json.loads(text_result.strip())
        return data.get('owners', [])
        
    except Exception as e:
        print(f"Error extracting Kofile parameters: {e}")
        return []
        
    finally:
        for uf in uploaded_files:
            try:
                client.files.delete(name=uf.name)
            except:
                pass


def initialize_sop_chat(api_key, sop_dirs, progress_callback=None):
    """
    Uploads SOP documents and initializes a strict Gemini chat session.
    """
    if not api_key:
        raise ValueError("No Gemini API key provided.")

    client = genai.Client(api_key=api_key)
    
    uploaded_files = []
    
    # Collect all PDFs and text files from the provided directories
    valid_files = []
    for d in sop_dirs:
        if os.path.exists(d):
            for root, dirs, files in os.walk(d):
                for f in files:
                    if f.lower().endswith(('.pdf', '.txt', '.md')) and not f.startswith("._"):
                        valid_files.append(os.path.join(root, f))
                        
    total_files = len(valid_files)
    if progress_callback:
        progress_callback(f"Found {total_files} documents. Starting upload...")
        
    # Upload files to Gemini
    for i, path in enumerate(valid_files):
        try:
            filename = os.path.basename(path)
            # Only upload if it's less than 50MB to be safe
            if os.path.getsize(path) < 50_000_000:
                if progress_callback:
                    progress_callback(f"Uploading {i+1}/{total_files}: {filename}...")
                    
                f = client.files.upload(
                    file=path,
                    config={'display_name': filename}
                )
                uploaded_files.append(f)
        except Exception as e:
            print(f"Failed to upload {path}: {e}")
            if progress_callback:
                progress_callback(f"Failed to upload {filename}: {e}")

    if progress_callback:
        progress_callback("Processing documents on server... Please wait.")
        
    import time
    # Wait for files to be ACTIVE
    active_files = []
    for f in uploaded_files:
        while True:
            file_info = client.files.get(name=f.name)
            state_str = str(getattr(file_info, 'state', ''))
            if 'ACTIVE' in state_str:
                active_files.append(f)
                break
            elif 'FAILED' in state_str:
                print(f"File {f.name} failed to process on server.")
                break
            time.sleep(2)

    system_instruction = """
    You are the Harbinger Land strict QA/QC and SOP Assistant.
    You have been provided with our company's SOPs, manuals, and feedback documents.
    
    CRITICAL RULES:
    1. You MUST ONLY answer questions using the provided documents. You are strictly forbidden from using outside knowledge or hallucinating.
    2. For every claim, rule, or answer you provide, you MUST include a citation in brackets referencing the exact Document Name and Page Number (e.g., [harbinger_preflight_checklist.pdf, Page 2]).
    3. If applicable, quote the exact relevant sentence from the document.
    4. If the answer is not found in the provided documents, you MUST strictly reply: "I do not have documentation for this." Do not attempt to guess or provide general advice.
    """
    
    return client, active_files, system_instruction

def generate_qc_check(api_key, pdf_path, doc_type):
    if not api_key:
        return "No API key provided."
    if not os.path.exists(pdf_path):
        return f"File not found: {pdf_path}"

    try:
        client = genai.Client(api_key=api_key)
        
        sample_file = client.files.upload(
            file=pdf_path, 
            config={'display_name': 'QC Document'}
        )
        
        base_prompt = (
            "Read the entire document extremely carefully word-by-word.\n"
            "You MUST provide your final answer in exactly this format:\n"
            "SUMMARY: Dower: [Yes/No] | Maturity: [Yes/No] | Released: [Yes/No]\n"
            "FULL TEXT: [Explicitly restate the questions first, then provide your detailed reasoning and exact quotes]\n\n"
            "Here are the questions to answer:\n"
        )
        prompt = ""
        if "deed" in doc_type.lower():
            prompt = base_prompt + "Question 1: Is the word 'dower' written anywhere in this document? Question 2: Does this document mention dower being released or retained?"
        elif "mortgage" in doc_type.lower():
            prompt = base_prompt + "Question 1: Is the word 'dower' written anywhere in this document? Question 2: Does this document mention a mortgage maturity date, due date, or a release?"
        else:
            prompt = base_prompt + "Question 1: Is the word 'dower' written anywhere in this document? Question 2: Does this document mention a maturity date?"

        response = client.models.generate_content(
            model='gemini-3.6-flash',
            contents=[sample_file, prompt]
        )
        
        try:
            client.files.delete(name=sample_file.name)
        except Exception:
            pass
            
        if response and response.text:
            return response.text.strip()
        else:
            return "No response from Gemini."

    except Exception as e:
        return f"AI Error: {str(e)}"

def extract_exact_title(api_key, pdf_path):
    from google import genai
    import os
    if not api_key:
        return ""
    if not os.path.exists(pdf_path):
        return ""

    try:
        client = genai.Client(api_key=api_key)
        
        sample_file = client.files.upload(
            file=pdf_path, 
            config={'display_name': 'Title Extraction'}
        )
        
        prompt = "Read the first page of this document. Extract the EXACT Document Title / Instrument Type (e.g. 'Warranty Deed', 'Release Of Mortgage', 'Oil And Gas Lease', 'Affidavit', etc). Reply ONLY with the exact title in Title Case. Do not include any other text, quotes, or explanations whatsoever."

        response = client.models.generate_content(
            model='gemini-3.6-flash',
            contents=[sample_file, prompt]
        )
        
        try:
            client.files.delete(name=sample_file.name)
        except Exception:
            pass
            
        if response and response.text:
            # Clean up potential markdown or weird punctuation
            title = response.text.strip().strip('"').strip("'").strip()
            if "TITLE:" in title.upper():
                title = title.upper().split("TITLE:")[1].strip().title()
            return title
        else:
            return ""

    except Exception as e:
        print(f"AI Title Error: {e}")
        return ""
