import pandas as pd
import pdfplumber
import re
import os

def extract_contingent_liabilities_tables(pdf_path):
    """
    Simple function to extract contingent liabilities tables from PDF
    Using the correct physical page mappings:
    - Logical page 346 → Physical page 176
    - Logical page 423 → Physical page 214
    
    Args:
        pdf_path (str): Path to your PDF file
        
    Returns:
        dict: Simple results with tables and text
    """
    
    # Correct page mappings as provided
    pages_to_extract = {
        'consolidated_page': 176,  # Contains logical page 346
        'standalone_page': 214     # Contains logical page 423
    }
    
    logical_pages = {
        'consolidated_page': 346,
        'standalone_page': 423
    }
    
    print(f"Using correct page mappings:")
    print(f"Logical page 346 → Physical page 176")
    print(f"Logical page 423 → Physical page 214")
    
    # First, check PDF and get total pages
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            print(f"PDF opened successfully. Total physical pages: {total_pages}")
    except Exception as e:
        print(f"✗ Error opening PDF: {e}")
        return {}
    
    # Validate pages exist
    valid_pages = {}
    for section_name, page_num in pages_to_extract.items():
        if page_num <= total_pages:
            valid_pages[section_name] = page_num
            print(f"✓ Will extract {section_name.replace('_', ' ')} from physical page {page_num}")
        else:
            print(f"✗ Physical page {page_num} doesn't exist (PDF has {total_pages} pages)")
    
    if not valid_pages:
        print("✗ No valid pages found.")
        return {}
    
    results = {}
    
    print("Starting extraction...")
    
    for section_name, page_num in valid_pages.items():
        print(f"\nProcessing {section_name.replace('_', ' ')} (Physical page {page_num})...")
        
        section_results = {
            'logical_page': logical_pages[section_name],
            'physical_page': page_num,
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
            
            # Step 3: Extract tables from the relevant section
            try:
                print(f"  Extracting tables from the identified section...")
                print(f"  Note: Tables have row lines but no column lines - using text-based extraction")
                
                # For tables with only row lines (no column lines), text extraction works better
                text_to_analyze = section_results['text_found']
                
                # Method 1: Look for structured financial data in text
                lines = text_to_analyze.split('\n')
                potential_table_lines = []
                
                print(f"  Analyzing {len(lines)} lines for tabular data...")
                
                for line in lines:
                    line = line.strip()
                    # Look for lines that are likely table rows:
                    # - Have financial amounts (₹, Rs., crore, lakh)
                    # - Have multiple numbers
                    # - Have consistent spacing patterns
                    # - Are not too short or too long
                    if (10 < len(line) < 200 and 
                        (line.count('₹') >= 1 or 
                         line.count('Rs.') >= 1 or
                         'crore' in line.lower() or
                         'lakh' in line.lower() or
                         # Lines with multiple numbers (likely amounts)
                         len([word for word in line.split() if word.replace(',', '').replace('.', '').replace('(', '').replace(')', '').isdigit()]) >= 1)):
                        
                        # Additional check: skip lines that are clearly headers or notes
                        line_lower = line.lower()
                        if not any(skip_word in line_lower for skip_word in 
                                 ['note', 'contingent', 'liabilit', 'statement', 'financial', 'as on', 'previous year']):
                            potential_table_lines.append(line)
                
                if potential_table_lines:
                    print(f"  ✓ Found {len(potential_table_lines)} potential financial data lines")
                    
                    # Method 2: Smart column detection for row-only tables
                    table_data = []
                    
                    # Analyze spacing patterns to detect columns
                    for line in potential_table_lines:
                        # Try different splitting strategies for tables without column lines:
                        
                        # Strategy 1: Split by significant spacing (2+ spaces)
                        row1 = re.split(r'\s{2,}', line.strip())
                        
                        # Strategy 2: Split by currency symbols and numbers
                        # Keep currency symbols with their numbers
                        row2 = re.split(r'(?<=\d)\s+(?=[A-Za-z])|(?<=[a-zA-Z])\s+(?=[\d₹])', line.strip())
                        
                        # Strategy 3: Split by positions (if there's a consistent pattern)
                        # Look for patterns like: Description [spaces] Amount [spaces] Amount
                        row3 = re.findall(r'([^₹\d]*(?:₹[\d,\.]+|[\d,\.]+\s*(?:crore|lakh)?|[\d,\.]+))', line)
                        
                        # Choose the best split (usually the one with reasonable number of columns)
                        best_row = row1
                        if 2 <= len(row2) <= 6 and len(row2) != len(row1):
                            best_row = row2
                        elif 2 <= len(row3) <= 6:
                            best_row = [item.strip() for item in row3 if item.strip()]
                        
                        # Clean up the row
                        cleaned_row = []
                        for cell in best_row:
                            cell = cell.strip()
                            if cell:  # Only add non-empty cells
                                cleaned_row.append(cell)
                        
                        if len(cleaned_row) >= 2:  # At least description + amount
                            table_data.append(cleaned_row)
                    
                    # Create DataFrame from extracted table data
                    if table_data:
                        # Determine maximum number of columns
                        max_cols = max(len(row) for row in table_data)
                        
                        # Pad shorter rows with empty strings
                        for row in table_data:
                            while len(row) < max_cols:
                                row.append('')
                        
                        # Create DataFrame with intelligent column names
                        if max_cols == 2:
                            col_names = ['Description', 'Amount']
                        elif max_cols == 3:
                            col_names = ['Description', 'Current Year', 'Previous Year']
                        elif max_cols == 4:
                            col_names = ['Description', 'Current Year', 'Previous Year', 'Notes']
                        else:
                            col_names = [f'Column_{i+1}' for i in range(max_cols)]
                        
                        df = pd.DataFrame(table_data, columns=col_names)
                        
                        # Final cleaning
                        df = df.replace('', pd.NA).dropna(how='all').fillna('')
                        
                        if not df.empty:
                            section_results['tables'].append(df)
                            print(f"  ✓ Created table from text analysis: {df.shape[0]} rows × {df.shape[1]} columns")
                            print(f"  ✓ Column structure: {list(df.columns)}")
                        else:
                            print(f"  ⚠ Table data found but became empty after cleaning")
                    else:
                        print(f"  ⚠ Could not structure the data into table format")
                
                # Method 3: Also try pdfplumber's table extraction (sometimes works even without column lines)
                try:
                    with pdfplumber.open(pdf_path) as pdf:
                        page = pdf.pages[page_num - 1]
                        page_width = page.width
                        half_width = page_width / 2
                        
                        # Focus on the correct half based on where we found the text
                        if found_section:
                            if "LEFT column" in str(locals().get('found_section', '')):
                                crop_area = (0, 0, half_width, page.height)
                                print(f"  Trying pdfplumber extraction on LEFT half...")
                            elif "RIGHT column" in str(locals().get('found_section', '')):
                                crop_area = (half_width, 0, page_width, page.height)
                                print(f"  Trying pdfplumber extraction on RIGHT half...")
                            else:
                                crop_area = None
                                print(f"  Trying pdfplumber extraction on full page...")
                            
                            if crop_area:
                                cropped_page = page.within_bbox(crop_area)
                                page_tables = cropped_page.extract_tables()
                            else:
                                page_tables = page.extract_tables()
                            
                            if page_tables:
                                print(f"  ✓ pdfplumber found {len(page_tables)} structured tables")
                                for i, table_data in enumerate(page_tables):
                                    if table_data and len(table_data) > 1:
                                        df = pd.DataFrame(table_data[1:], columns=table_data[0])
                                        df = df.dropna(how='all').dropna(axis=1, how='all').fillna('')
                                        
                                        if not df.empty:
                                            section_results['tables'].append(df)
                                            print(f"  ✓ Added pdfplumber table {i+1}: {df.shape}")
                            else:
                                print(f"  ⚠ pdfplumber found no structured tables (expected for row-only tables)")
                        
                except Exception as e:
                    print(f"  ⚠ pdfplumber extraction failed: {e}")
                
                if not section_results['tables']:
                    print(f"  ℹ No tables extracted - all financial data will be available in the text section")
                
            except Exception as e:
                print(f"  ✗ Table extraction failed: {e}")
                print(f"  ℹ Continuing with text extraction only...")
            
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

def search_for_section(pdf_path, section_type, total_pages):
    """
    Search for contingent liabilities sections in the PDF
    
    Args:
        pdf_path (str): Path to PDF
        section_type (str): 'consolidated' or 'standalone'
        total_pages (int): Total pages in PDF
        
    Returns:
        int: Page number if found, None otherwise
    """
    
    search_terms = {
        'consolidated': ['consolidated financial statement', 'consolidated financial statements'],
        'standalone': ['standalone financial statement', 'standalone financial statements']
    }
    
    print(f"  Searching for {section_type} section...")
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Search in reasonable range (avoid searching entire PDF for large files)
            start_page = max(1, total_pages - 100)  # Search last 100 pages
            end_page = total_pages
            
            for page_num in range(start_page, end_page + 1):
                try:
                    page = pdf.pages[page_num - 1]
                    text = page.extract_text()
                    if text:
                        text_lower = text.lower()
                        
                        # Check for section type and contingent liabilities
                        section_found = any(term in text_lower for term in search_terms[section_type])
                        contingent_found = 'contingent liabilit' in text_lower
                        
                        if section_found and contingent_found:
                            print(f"  ✓ Found {section_type} section on page {page_num}")
                            return page_num
                            
                except Exception as e:
                    continue  # Skip problematic pages
                    
    except Exception as e:
        print(f"  ✗ Search failed: {e}")
    
    print(f"  ✗ {section_type} section not found")
    return None

def main():
    """
    Simple main function to run everything
    """
    
    # Set your PDF path here - CHANGE THIS TO YOUR ACTUAL PDF PATH
    pdf_path = input("Enter the path to your PDF file: ").strip().strip('"').strip("'")
    
    # Alternative: uncomment and set your PDF path directly
    # pdf_path = r"C:\path\to\your\pdf\file.pdf"  # Windows
    # pdf_path = "/path/to/your/pdf/file.pdf"     # Mac/Linux
    
    if not pdf_path or not os.path.exists(pdf_path):
        print("❌ PDF file not found. Please check the path.")
        print("Example paths:")
        print("  Windows: C:\\Documents\\your_file.pdf")
        print("  Mac/Linux: /Users/username/Documents/your_file.pdf")
        return
    
    print("PDF Contingent Liabilities Extractor")
    print("=" * 40)
    
    # Extract data using the correct physical page numbers
    results = extract_contingent_liabilities_tables(pdf_path)
    
    if not results:
        print("No sections found. Let me help you find the right pages...")
        
        # Show some sample pages to help user identify correct pages
        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                print(f"\nYour PDF has {total_pages} pages.")
                print("Let me show you some pages that mention 'contingent liabilities':\n")
                
                found_pages = []
                for page_num in range(1, min(total_pages + 1, 50)):  # Check first 50 pages
                    try:
                        page = pdf.pages[page_num - 1]
                        text = page.extract_text()
                        if text and 'contingent liabilit' in text.lower():
                            found_pages.append(page_num)
                            print(f"Page {page_num}: Found 'contingent liabilities'")
                            # Show first 200 chars to help identify
                            start_pos = text.lower().find('contingent liabilit')
                            if start_pos >= 0:
                                snippet = text[max(0, start_pos-50):start_pos+150]
                                print(f"  Context: ...{snippet}...")
                            print()
                            
                            if len(found_pages) >= 5:  # Limit output
                                break
                    except:
                        continue
                
                if found_pages:
                    print(f"Found 'contingent liabilities' on pages: {found_pages}")
                    print("\nTo extract specific pages, modify the code like this:")
                    print("results = extract_contingent_liabilities_tables(pdf_path,")
                    print(f"                                               consolidated_page={found_pages[0]},")
                    if len(found_pages) > 1:
                        print(f"                                               standalone_page={found_pages[1]})")
                    else:
                        print("                                               standalone_page=None)")
                else:
                    print("No pages found with 'contingent liabilities'. Please check your PDF.")
                    
        except Exception as e:
            print(f"Error analyzing PDF: {e}")
        
        return
    
    # Print simple summary
    print("\n" + "=" * 40)
    print("EXTRACTION COMPLETE")
    print("=" * 40)
    
    for section_name, data in results.items():
        section_title = section_name.replace('_', ' ').title()
        print(f"{section_title}:")
        print(f"  Logical page: {data['logical_page']}")
        print(f"  Physical page: {data['physical_page']}")
        print(f"  Tables found: {len(data['tables'])}")
        print(f"  Text extracted: {len(data['text_found'])} characters")
        
        if data['tables']:
            print("  Table shapes:")
            for i, table in enumerate(data['tables']):
                print(f"    Table {i+1}: {table.shape[0]} rows × {table.shape[1]} columns")
        print()
    
    print("✓ All done!")

# Run the extraction
if __name__ == "__main__":
    main()
