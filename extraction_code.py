import pandas as pd
import pdfplumber
import re
import os

def extract_contingent_liabilities_tables(pdf_path):
    """
    Extract contingent liabilities tables from PDF - reads continuously across pages until total is found
    """
    
    # Starting page mappings
    start_pages = {
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
    
    results = {}
    
    for section_name, start_page in start_pages.items():
        print(f"Processing {section_name.replace('_', ' ')} starting from page {start_page}...")
        
        section_results = {
            'logical_page': logical_pages[section_name],
            'physical_page': start_page,
            'text_found': '',
            'tables': []
        }
        
        try:
            # Determine target section
            if 'consolidated' in section_name:
                target_heading = "notes to consolidated financial statement"
            else:
                target_heading = "notes to standalone financial statement"
            
            # Start continuous reading from the starting page
            all_extracted_text = ""
            found_start = False
            found_total = False
            current_page = start_page
            
            with pdfplumber.open(pdf_path) as pdf:
                
                while current_page <= total_pages and not found_total:
                    print(f"  Reading page {current_page}...")
                    
                    page = pdf.pages[current_page - 1]
                    page_width = page.width
                    half_width = page_width / 2
                    
                    # Read left column first, then right column
                    columns = [
                        ("LEFT", page.within_bbox((0, 0, half_width, page.height))),
                        ("RIGHT", page.within_bbox((half_width, 0, page_width, page.height)))
                    ]
                    
                    for column_name, column_area in columns:
                        if found_total:
                            break
                            
                        column_text = column_area.extract_text() or ""
                        if not column_text:
                            continue
                        
                        column_lower = column_text.lower()
                        
                        # If we haven't found the start yet, look for it
                        if not found_start:
                            if (target_heading in column_lower and "contingent liabilit" in column_lower):
                                print(f"    ✓ Found section start in page {current_page} {column_name} column")
                                found_start = True
                                
                                # Find the contingent liabilities section start
                                patterns = [r'b\)\s*contingent\s+liabilit[ieysn]*', r'contingent\s+liabilit[ieysn]*']
                                start_pos = 0
                                
                                for pattern in patterns:
                                    match = re.search(pattern, column_lower)
                                    if match:
                                        start_pos = match.start()
                                        break
                                
                                column_text = column_text[start_pos:]
                        
                        # If we found the start, continue collecting text
                        if found_start:
                            # Check each line for total/grand total
                            lines = column_text.split('\n')
                            
                            for line in lines:
                                all_extracted_text += line + '\n'
                                
                                # Check if this line contains total/grand total with amounts
                                line_lower = line.lower().strip()
                                if (('total' in line_lower or 'grand total' in line_lower) and 
                                    (any(c.isdigit() for c in line) or '₹' in line or 'rs.' in line_lower or 
                                     'crore' in line_lower or 'lakh' in line_lower)):
                                    print(f"    ✓ Found total line in page {current_page} {column_name}: {line.strip()}")
                                    found_total = True
                                    break
                    
                    # Move to next page if total not found
                    if not found_total:
                        current_page += 1
                
                if found_start:
                    section_results['text_found'] = all_extracted_text.strip()
                    print(f"  ✓ Extracted text from pages {start_page} to {current_page}")
                    print(f"  ✓ Total characters: {len(section_results['text_found'])}")
                else:
                    print(f"  ✗ Section not found starting from page {start_page}")
            
            results[section_name] = section_results
            
        except Exception as e:
            print(f"Failed to process section {section_name}: {e}")
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
