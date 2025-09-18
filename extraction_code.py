import pandas as pd
import pdfplumber
import re
import os

def extract_contingent_liabilities_tables(pdf_path):
    """
    Extract contingent liabilities tables from PDF pages 176 and 214
    """
    
    # Page mappings
    pages_to_extract = {
        'consolidated_page': 176,  # Contains logical page 346
        'standalone_page': 214     # Contains logical page 423
    }
    
    logical_pages = {
        'consolidated_page': 346,
        'standalone_page': 423
    }
    
    # Validate PDF
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            print(f"PDF opened successfully. Total pages: {total_pages}")
    except Exception as e:
        print(f"Error opening PDF: {e}")
        return {}
    
    # Check pages exist
    valid_pages = {}
    for section_name, page_num in pages_to_extract.items():
        if page_num <= total_pages:
            valid_pages[section_name] = page_num
        else:
            print(f"Page {page_num} doesn't exist")
    
    if not valid_pages:
        return {}
    
    results = {}
    
    for section_name, page_num in valid_pages.items():
        print(f"Processing {section_name.replace('_', ' ')} (Page {page_num})...")
        
        section_results = {
            'logical_page': logical_pages[section_name],
            'physical_page': page_num,
            'text_found': '',
            'tables': []
        }
        
        try:
            # Extract text from left and right columns separately
            with pdfplumber.open(pdf_path) as pdf:
                page = pdf.pages[page_num - 1]
                
                # Split page into left/right halves
                page_width = page.width
                half_width = page_width / 2
                
                left_text = page.within_bbox((0, 0, half_width, page.height)).extract_text() or ""
                right_text = page.within_bbox((half_width, 0, page_width, page.height)).extract_text() or ""
                
                # Determine target section
                if 'consolidated' in section_name:
                    target_heading = "notes to consolidated financial statement"
                else:
                    target_heading = "notes to standalone financial statement"
                
                # Find which column contains our section
                relevant_text = ""
                left_lower = left_text.lower()
                right_lower = right_text.lower()
                
                if target_heading in left_lower and "contingent liabilit" in left_lower:
                    relevant_text = left_text
                elif target_heading in right_lower and "contingent liabilit" in right_lower:
                    relevant_text = right_text
                elif "contingent liabilit" in left_lower:
                    relevant_text = left_text
                elif "contingent liabilit" in right_lower:
                    relevant_text = right_text
                
                if relevant_text:
                    # Find contingent liabilities section
                    text_lower = relevant_text.lower()
                    patterns = [r'b\)\s*contingent\s+liabilit[ieysn]*', r'contingent\s+liabilit[ieysn]*']
                    
                    for pattern in patterns:
                        match = re.search(pattern, text_lower)
                        if match:
                            start = match.start()
                            
                            # Find section end
                            end_patterns = [r'\bc\)\s*', r'\bd\)\s*', r'\n\s*\d+\.\s*']
                            search_start = match.end() + 50
                            end = len(relevant_text)
                            
                            for end_pattern in end_patterns:
                                end_match = re.search(end_pattern, text_lower[search_start:])
                                if end_match:
                                    end = search_start + end_match.start()
                                    break
                            
                            section_results['text_found'] = relevant_text[start:end].strip()
                            break
                    
                    if not section_results['text_found']:
                        section_results['text_found'] = relevant_text
                
            results[section_name] = section_results
            
        except Exception as e:
            print(f"Failed to process page {page_num}: {e}")
            results[section_name] = section_results
    
    return results

def save_results_to_word(results, output_file="contingent_liabilities_extracted.docx"):
    """
    Save extracted text to Word document
    """
    
    try:
        from docx import Document
        
        doc = Document()
        doc.add_heading('Contingent Liabilities Data', 0)
        
        for section_name, data in results.items():
            section_title = section_name.replace('_', ' ').title()
            
            doc.add_heading(f'{section_title}', level=1)
            doc.add_paragraph(f"Page {data['physical_page']} (Logical page {data['logical_page']})")
            
            text = data['text_found']
            if not text:
                doc.add_paragraph("No text was extracted for this section.")
                continue
            
            doc.add_paragraph(text)
            
            # Add page break between sections
            if section_name != list(results.keys())[-1]:
                doc.add_page_break()
        
        doc.save(output_file)
        print(f"Word document saved: {output_file}")
        return output_file
        
    except ImportError:
        print("python-docx not installed. Install with: pip install python-docx")
        return None
    except Exception as e:
        print(f"Error saving Word document: {e}")
        return None

def main():
    """
    Main function to run extraction
    """
    
    pdf_path = "your_pdf_file.pdf"  # Change this to your PDF path
    
    print("PDF Contingent Liabilities Extractor")
    print("=" * 40)
    
    results = extract_contingent_liabilities_tables(pdf_path)
    
    if not results:
        print("No sections found.")
        return
    
    # Save to Word
    if results:
        save_results_to_word(results)
    
    # Print summary
    print("\nExtraction Complete:")
    for section_name, data in results.items():
        section_title = section_name.replace('_', ' ').title()
        print(f"{section_title}: {len(data['text_found'])} characters extracted")

# Run the extraction
if __name__ == "__main__":
    main()
