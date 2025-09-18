import pandas as pd
import pdfplumber
import tabula
import re

def extract_contingent_liabilities_tables(pdf_path):
    """
    Simple function to extract contingent liabilities tables from PDF
    
    Args:
        pdf_path (str): Path to your PDF file
        
    Returns:
        dict: Simple results with tables and text
    """
    
    # Target pages
    pages_to_extract = {
        'consolidated_page': 346,
        'standalone_page': 423
    }
    
    results = {}
    
    print("Starting extraction...")
    
    for section_name, page_num in pages_to_extract.items():
        print(f"\nProcessing {section_name} (Page {page_num})...")
        
        section_results = {
            'page_number': page_num,
            'text_found': '',
            'tables': []
        }
        
        try:
            # Step 1: Extract text from the page
            with pdfplumber.open(pdf_path) as pdf:
                page = pdf.pages[page_num - 1]  # Convert to 0-indexed
                page_text = page.extract_text()
                
                # Step 2: Find contingent liabilities section
                text_lower = page_text.lower()
                
                # Look for contingent liabilities heading
                patterns = [
                    r'b\)\s*contingent\s+liabilit[ieysn]*',
                    r'contingent\s+liabilit[ieysn]*'
                ]
                
                found_section = False
                for pattern in patterns:
                    match = re.search(pattern, text_lower)
                    if match:
                        found_section = True
                        # Extract text around the match (500 chars before and after)
                        start = max(0, match.start() - 500)
                        end = min(len(page_text), match.end() + 1500)
                        section_results['text_found'] = page_text[start:end].strip()
                        print(f"  ✓ Found contingent liabilities section")
                        break
                
                if not found_section:
                    print(f"  ⚠ Contingent liabilities heading not found, using full page")
                    section_results['text_found'] = page_text
            
            # Step 3: Extract tables using tabula
            try:
                print(f"  Extracting tables...")
                tables = tabula.read_pdf(
                    pdf_path, 
                    pages=page_num,
                    multiple_tables=True,
                    pandas_options={'header': 0}
                )
                
                for i, table in enumerate(tables):
                    if not table.empty:
                        # Simple cleaning
                        cleaned_table = table.dropna(how='all').dropna(axis=1, how='all')
                        cleaned_table = cleaned_table.fillna('')
                        
                        if not cleaned_table.empty:
                            section_results['tables'].append(cleaned_table)
                            print(f"  ✓ Extracted table {i+1}: {cleaned_table.shape}")
                
            except Exception as e:
                print(f"  ⚠ Tabula failed: {e}")
                
                # Fallback: try pdfplumber
                try:
                    with pdfplumber.open(pdf_path) as pdf:
                        page = pdf.pages[page_num - 1]
                        page_tables = page.extract_tables()
                        
                        for i, table_data in enumerate(page_tables):
                            if table_data and len(table_data) > 1:
                                df = pd.DataFrame(table_data[1:], columns=table_data[0])
                                df = df.dropna(how='all').dropna(axis=1, how='all')
                                df = df.fillna('')
                                
                                if not df.empty:
                                    section_results['tables'].append(df)
                                    print(f"  ✓ Extracted table {i+1} with pdfplumber: {df.shape}")
                                    
                except Exception as e2:
                    print(f"  ✗ Both methods failed: {e2}")
            
            results[section_name] = section_results
            print(f"  Done: Found {len(section_results['tables'])} tables")
            
        except Exception as e:
            print(f"  ✗ Failed to process page {page_num}: {e}")
            results[section_name] = section_results
    
    return results

def save_simple_results(results, output_file="contingent_liabilities_extracted.xlsx"):
    """
    Save all results to one Excel file with multiple sheets
    
    Args:
        results (dict): Extraction results
        output_file (str): Output Excel file name
    """
    
    print(f"\nSaving results to {output_file}...")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        
        # Create summary sheet
        summary_data = []
        for section_name, data in results.items():
            summary_data.append({
                'Section': section_name.replace('_', ' ').title(),
                'Page Number': data['page_number'],
                'Text Extracted': 'Yes' if data['text_found'] else 'No',
                'Number of Tables': len(data['tables']),
                'Text Length': len(data['text_found'])
            })
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # Save text content
        text_data = []
        for section_name, data in results.items():
            text_data.append({
                'Section': section_name.replace('_', ' ').title(),
                'Page': data['page_number'],
                'Extracted Text': data['text_found']
            })
        
        text_df = pd.DataFrame(text_data)
        text_df.to_excel(writer, sheet_name='Extracted Text', index=False)
        
        # Save all tables
        table_counter = 1
        for section_name, data in results.items():
            for i, table in enumerate(data['tables']):
                sheet_name = f"{section_name.split('_')[0].title()}_Table_{i+1}"
                # Limit sheet name to 31 characters (Excel limit)
                if len(sheet_name) > 31:
                    sheet_name = f"Table_{table_counter}"
                
                table.to_excel(writer, sheet_name=sheet_name, index=False)
                table_counter += 1
    
    print(f"✓ Results saved successfully!")

def main():
    """
    Simple main function to run everything
    """
    
    # Set your PDF path here
    pdf_path = "your_pdf_file.pdf"  # Change this to your actual PDF path
    
    print("PDF Contingent Liabilities Extractor")
    print("=" * 40)
    
    # Extract data
    results = extract_contingent_liabilities_tables(pdf_path)
    
    # Save to Excel
    save_simple_results(results)
    
    # Print simple summary
    print("\n" + "=" * 40)
    print("EXTRACTION COMPLETE")
    print("=" * 40)
    
    for section_name, data in results.items():
        section_title = section_name.replace('_', ' ').title()
        print(f"{section_title}:")
        print(f"  Page: {data['page_number']}")
        print(f"  Tables found: {len(data['tables'])}")
        print(f"  Text extracted: {len(data['text_found'])} characters")
        
        if data['tables']:
            print("  Table shapes:")
            for i, table in enumerate(data['tables']):
                print(f"    Table {i+1}: {table.shape[0]} rows × {table.shape[1]} columns")
        print()
    
    print("Output file: contingent_liabilities_extracted.xlsx")
    print("✓ All done!")

# Run the extraction
if __name__ == "__main__":
    main()
