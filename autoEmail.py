import win32com.client
from datetime import datetime, timedelta
import os
import sys
import re

def get_column_letter(col_num):
    """Convert column number to Excel column letter"""
    result = ""
    while col_num > 0:
        col_num -= 1
        result = chr(col_num % 26 + ord('A')) + result
        col_num //= 26
    return result

def column_letter_to_num(letter):
    """Convert Excel column letter to number"""
    result = 0
    for char in letter:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result

def calculate_end_column(start_col, month):
    """Calculate end column based on month"""
    start_num = column_letter_to_num(start_col)
    end_num = start_num + month  # 1月: +1, 2月: +2...
    return get_column_letter(end_num)

def find_row_by_text(worksheet, search_text, search_range):
    """Find row number containing specific text in the given range"""
    try:
        found_range = worksheet.Range(search_range).Find(search_text)
        if found_range:
            return found_range.Row
        return None
    except Exception as e:
        print(f"Error finding text '{search_text}': {e}")
        return None

def get_cell_value(worksheet, row, col_letter):
    """Get cell value from worksheet"""
    try:
        cell_address = f"{col_letter}{row}"
        return worksheet.Range(cell_address).Value
    except Exception as e:
        print(f"Error getting cell value at {col_letter}{row}: {e}")
        return 0

def get_cell_text(worksheet, row, col_letter):
    """Get cell text (display value) from worksheet"""
    try:
        cell_address = f"{col_letter}{row}"
        return worksheet.Range(cell_address).Text
    except Exception as e:
        print(f"Error getting cell text at {col_letter}{row}: {e}")
        return ""

def format_digital_account_number(value):
    """Format number for digital account display (always show as full number with commas)"""
    if value is None:
        return "0"
    if isinstance(value, (int, float)):
        return f"{round(value):,}"
    return str(value)

def format_platform_revenue(value):
    """Format number for platform revenue display (萬元 or 億元) with thousand separator"""
    if value is None:
        return "0"
    if isinstance(value, (int, float)):
        if value >= 100000000:  # 億
            return f"{round(value/100000000, 1)}億元"
        elif value >= 10000:  # 萬
            wan_value = round(value/10000)
            if wan_value >= 1000:  # Add thousand separator for 萬元 values >= 1000
                return f"{wan_value:,}萬元"
            else:
                return f"{wan_value}萬元"
        else:
            return f"{round(value):,}元"
    return str(value)

def format_percentage_from_text(text_value):
    """Format percentage from Excel cell text (extract number and format to 1 decimal place)"""
    if text_value is None or text_value == "":
        return "0%"
    
    # Convert to string if not already
    text_str = str(text_value)
    
    # Extract numeric part using regex
    # Look for numbers (including decimals) followed by % or just numbers
    match = re.search(r'([-]?\d+\.?\d*)', text_str)
    if match:
        try:
            numeric_value = float(match.group(1))
            # Round to 1 decimal place and remove trailing zeros
            rounded_value = round(numeric_value, 1)
            if rounded_value == int(rounded_value):
                return f"{int(rounded_value)}%"
            else:
                return f"{rounded_value}%"
        except ValueError:
            pass
    
    return "0%"

def convert_formulas_to_values(worksheet, start_row, end_row, start_col, end_col):
    """Convert formulas to values in specified range"""
    try:
        # Create range string
        range_str = f"{start_col}{start_row}:{end_col}{end_row}"
        print(f"Converting formulas to values in range: {range_str}")
        
        # Get the range
        range_obj = worksheet.Range(range_str)
        
        # Copy the range
        range_obj.Copy()
        
        # Paste as values only
        range_obj.PasteSpecial(Paste=-4163)  # xlPasteValues = -4163
        
        # Clear clipboard
        worksheet.Application.CutCopyMode = False
        
        print(f"Successfully converted formulas to values in range {range_str}")
        
    except Exception as e:
        print(f"Error converting formulas to values: {e}")

def get_dynamic_values(ws_digital_account, ws_digital_platform, target_month):
    """Get dynamic values from Excel worksheets using target month"""
    
    # Calculate target column (target month offset from base column)
    digital_account_base_col = "Q"  # 數位戶表格起始列
    digital_platform_base_col = "P"  # 數位平台收益表格起始列
    
    # Calculate actual column letters for target month
    digital_account_col_num = column_letter_to_num(digital_account_base_col) + target_month
    digital_platform_col_num = column_letter_to_num(digital_platform_base_col) + target_month
    
    digital_account_col = get_column_letter(digital_account_col_num)
    digital_platform_col = get_column_letter(digital_platform_col_num)
    
    print(f"數位戶 target column: {digital_account_col}")
    print(f"數位平台收益 target column: {digital_platform_col}")
    
    values = {}
    
    try:
        # 數位戶客戶數 values from 數位戶 worksheet
        # Find rows for each metric (search in a reasonable range)
        search_range = "A1:Z100"  # Adjust range as needed
        
        # Find 月目標數 row
        month_target_row = find_row_by_text(ws_digital_account, "月目標數", search_range)
        if month_target_row:
            values['digital_month_target'] = get_cell_value(ws_digital_account, month_target_row, digital_account_col)
        
        # Find 數位戶實績(存戶+卡戶) row
        digital_actual_row = find_row_by_text(ws_digital_account, "數位戶實績(存戶+卡戶)", search_range)
        if digital_actual_row:
            values['digital_actual'] = get_cell_value(ws_digital_account, digital_actual_row, digital_account_col)
        
        # Find 月目標達成率 row - get text instead of value for percentage
        digital_rate_row = find_row_by_text(ws_digital_account, "月目標達成率", search_range)
        if digital_rate_row:
            values['digital_achievement_rate_text'] = get_cell_text(ws_digital_account, digital_rate_row, digital_account_col)
        
        # 數位平台收益 values from 數位平台收益 worksheet
        # Find 月目標數 row
        platform_month_target_row = find_row_by_text(ws_digital_platform, "月目標數", search_range)
        if platform_month_target_row:
            values['platform_month_target'] = get_cell_value(ws_digital_platform, platform_month_target_row, digital_platform_col)
        
        # Find 實際數位平台收益 row
        platform_actual_row = find_row_by_text(ws_digital_platform, "實際數位平台收益", search_range)
        if platform_actual_row:
            values['platform_actual'] = get_cell_value(ws_digital_platform, platform_actual_row, digital_platform_col)
        
        # Find 月目標達成率 row - get text instead of value for percentage
        platform_rate_row = find_row_by_text(ws_digital_platform, "月目標達成率", search_range)
        if platform_rate_row:
            values['platform_achievement_rate_text'] = get_cell_text(ws_digital_platform, platform_rate_row, digital_platform_col)
        
        # Find 累積月目標數 row
        platform_cumulative_target_row = find_row_by_text(ws_digital_platform, "累積月目標數", search_range)
        if platform_cumulative_target_row:
            values['platform_cumulative_target'] = get_cell_value(ws_digital_platform, platform_cumulative_target_row, digital_platform_col)
        
        # Find 累積月實際數 row
        platform_cumulative_actual_row = find_row_by_text(ws_digital_platform, "累積月實際數", search_range)
        if platform_cumulative_actual_row:
            values['platform_cumulative_actual'] = get_cell_value(ws_digital_platform, platform_cumulative_actual_row, digital_platform_col)
        
        # Find 累積月目標達成率 row - get text instead of value for percentage
        platform_cumulative_rate_row = find_row_by_text(ws_digital_platform, "累積月目標達成率", search_range)
        if platform_cumulative_rate_row:
            values['platform_cumulative_rate_text'] = get_cell_text(ws_digital_platform, platform_cumulative_rate_row, digital_platform_col)
        
    except Exception as e:
        print(f"Error getting dynamic values: {e}")
    
    return values

def get_user_input_month():
    """Get target month from user input"""
    while True:
        try:
            print("請輸入要產出報表的月份 (1-12):")
            month_input = input("月份: ").strip()
            target_month = int(month_input)
            
            if 1 <= target_month <= 12:
                return target_month
            else:
                print("請輸入1-12之間的數字")
        except ValueError:
            print("請輸入有效的數字")

def copy_excel_range_to_outlook(worksheet, range_address):
    """Copy Excel range and return it for pasting into Outlook"""
    try:
        range_obj = worksheet.Range(range_address)
        range_obj.Copy()
        return range_obj
    except Exception as e:
        print(f"Error copying range {range_address}: {e}")
        return None

def copy_excel_range_with_deletion(worksheet, range_address, delete_rows_start, delete_rows_end):
    """Copy Excel range and delete specific rows after copying, with formula conversion"""
    try:
        # First copy the entire range
        range_obj = worksheet.Range(range_address)
        range_obj.Copy()
        
        # Create a temporary worksheet to manipulate the data
        temp_ws = worksheet.Parent.Worksheets.Add()
        temp_ws.Name = "TempData"
        
        # Paste the copied data to temporary worksheet (preserve formatting)
        temp_ws.Range("A1").PasteSpecial(Paste=-4104)  # xlPasteAll = -4104 (values + formatting)
        
        # **重要：將公式轉換為數值，避免刪除行後出現 #REF! 錯誤**
        # 獲取貼上的資料範圍
        used_range = temp_ws.UsedRange
        if used_range:
            print("Converting formulas to values in temporary worksheet...")
            # 複製範圍
            used_range.Copy()
            # 貼上為數值（這會將所有公式轉換為計算結果）
            used_range.PasteSpecial(Paste=-4163)  # xlPasteValues = -4163
            # 清除剪貼簿
            worksheet.Parent.Application.CutCopyMode = False
            print("Successfully converted formulas to values")
        
        # Calculate the rows to delete in the temporary worksheet
        # If original range starts at row 11 and we want to delete rows 23-31,
        # in the temp sheet this would be rows 13-21 (23-11+1 to 31-11+1)
        original_start_row = int(range_address.split(':')[0][1:])  # Extract row number from range
        temp_delete_start = delete_rows_start - original_start_row + 1
        temp_delete_end = delete_rows_end - original_start_row + 1
        
        # Delete the unwanted rows in temporary worksheet
        if temp_delete_start > 0 and temp_delete_end > 0:
            delete_range = f"{temp_delete_start}:{temp_delete_end}"
            print(f"Deleting rows {temp_delete_start}-{temp_delete_end} from temporary data...")
            temp_ws.Rows(delete_range).Delete()
            print(f"Successfully deleted rows {temp_delete_start}-{temp_delete_end} from temporary data")
        
        # Copy the modified data with formatting (now all values, no formulas)
        used_range = temp_ws.UsedRange
        if used_range:
            used_range.Copy()
            print("Final data copied successfully (formulas converted to values)")
        else:
            print("Warning: No data found in temporary worksheet")
            return None
        
        # Don't delete the temporary worksheet immediately - let the caller handle it
        # Store reference for later cleanup
        temp_ws.Name = f"TempData_{worksheet.Name}"
        
        return used_range
        
    except Exception as e:
        print(f"Error copying range with deletion {range_address}: {e}")
        # Clean up temporary worksheet if it exists
        try:
            if 'temp_ws' in locals():
                worksheet.Parent.Application.DisplayAlerts = False
                temp_ws.Delete()
                worksheet.Parent.Application.DisplayAlerts = True
                print("Cleaned up temporary worksheet due to error")
        except:
            pass
        return None

def get_word_document_content_with_formatting(word_file_path):
    """Get content from Word document with formatting preserved"""
    try:
        import win32com.client
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        
        # Open the Word document
        word_doc = word_app.Documents.Open(word_file_path)
        
        # Select all content in the document
        word_doc.Content.Select()
        
        # Copy the content with formatting
        word_doc.Content.Copy()
        
        # Get the plain text as backup
        content_text = word_doc.Content.Text
        
        # Close document and Word application
        word_doc.Close(False)
        word_app.Quit()
        
        print(f"Successfully retrieved formatted content from {word_file_path}")
        return content_text.strip(), True  # Return text and flag indicating formatting is copied
        
    except Exception as e:
        print(f"Error getting Word document content with formatting: {e}")
        return None, False

def insert_signature_to_email(word_doc, signature_path):
    """Insert signature from Word document to email with formatting preserved"""
    try:
        # Check if signature file exists
        if not os.path.exists(signature_path):
            print(f"Signature file not found: {signature_path}")
            return False
        
        # Get content from Word document with formatting
        signature_content, has_formatting = get_word_document_content_with_formatting(signature_path)
        
        if not signature_content:
            print("Failed to retrieve content from SIGN.docx")
            return False
        
        print("Adding signature to email with formatting...")
        
        # Method 1: Try using Selection object with paste (preserves formatting)
        try:
            selection = word_doc.Application.Selection
            selection.EndKey(6)  # wdStory = 6 (move to end of document)
            selection.TypeText("\n\n")
            
            if has_formatting:
                # Paste with formatting if we have it in clipboard
                selection.Paste()
                print("✓ Signature added successfully with formatting using Selection.Paste!")
            else:
                # Fallback to plain text
                selection.TypeText(signature_content)
                print("✓ Signature added as plain text using Selection method!")
            return True
        except Exception as e1:
            print(f"Selection paste method failed: {e1}")
        
        # Method 2: Try using Range object with paste
        try:
            content_range = word_doc.Content
            content_range.Collapse(0)  # Collapse to end
            content_range.InsertAfter("\n\n")
            content_range.Collapse(0)  # Collapse to end again
            
            if has_formatting:
                content_range.Paste()
                print("✓ Signature added successfully with formatting using Range.Paste!")
            else:
                content_range.InsertAfter(signature_content)
                print("✓ Signature added as plain text using Range method!")
            return True
        except Exception as e2:
            print(f"Range paste method failed: {e2}")
        
        # Method 3: Try direct content manipulation (plain text only)
        try:
            current_content = word_doc.Content.Text
            word_doc.Content.Text = current_content + "\n\n" + signature_content
            print("✓ Signature added as plain text using direct content method!")
            return True
        except Exception as e3:
            print(f"Direct content method failed: {e3}")
        
        return False
        
    except Exception as e:
        print(f"Error inserting signature: {e}")
        return False

def main():
    excel = None
    outlook = None
    workbook = None
    
    try:
        # Get user input for target month
        target_month = get_user_input_month()
        print(f"選擇的月份: {target_month}月")
        
        # Get current year for file path
        current_date = datetime.now()
        current_year = current_date.year
        
        # Format month as MM
        formatted_month = f"{target_month:02d}"
        
        # Use target month as the formatted date for email content
        formatted_date = f"{target_month:02d}"
        
        # Construct file name and path
        file_name = f"{current_year}統計({current_year}{formatted_month})"
        file_path = f"\\\\X.X.X.X\\{file_name}.xlsx"
        
        # Check if file exists
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            input("Press Enter to exit...")
            return
        
        print(f"Processing file: {file_name}")
        
        # Create Excel application
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Open workbook
        workbook = excel.Workbooks.Open(file_path)
        
        # Find worksheets
        ws_digital_account = None
        ws_digital_platform = None
        
        for sheet in workbook.Worksheets:
            if sheet.Name == "數位戶":
                ws_digital_account = sheet
            elif sheet.Name == "數位平台收益":
                ws_digital_platform = sheet
        
        if ws_digital_account is None:
            print("Worksheet '數位戶' not found!")
            return
        
        if ws_digital_platform is None:
            print("Worksheet '數位平台收益' not found!")
            return
        
        # Get dynamic values from Excel using target month
        dynamic_values = get_dynamic_values(ws_digital_account, ws_digital_platform, target_month)
        
        # Calculate dynamic column ranges using target month
        # 數位戶: Q50 to (Q+month)60
        digital_account_end_col = calculate_end_column("Q", target_month)
        range1 = f"Q50:{digital_account_end_col}60"
        
        # 數位平台收益: P11 to (P+month)41 (整個範圍，稍後刪除23-31行)
        digital_platform_end_col = calculate_end_column("P", target_month)
        range2 = f"P11:{digital_platform_end_col}41"
        
        print(f"Target Month: {target_month}")
        print(f"數位戶 range: {range1}")
        print(f"數位平台收益 range: {range2} (will delete rows 23-31)")
        
        # Create Outlook application
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # olMailItem = 0
        
        # Set email properties
        mail.Subject = "(週報)績效數字統計"
        
        # Create email body with dynamic values
        mail_body = f"""Dear all,

至{formatted_date}績效數字統計及說明如下，謝謝。

(1) 數位戶客戶數: 年目標為1,564,000戶，月目標{format_digital_account_number(dynamic_values.get('digital_month_target', 0))}戶，目前實際數為{format_digital_account_number(dynamic_values.get('digital_actual', 0))}戶，
月目標達成率為{format_percentage_from_text(dynamic_values.get('digital_achievement_rate_text', '0%'))}。
網行銀客戶數(具有網行銀會員身分之存戶+卡戶)

[TABLE1_PLACEHOLDER]

(2)數位平台收益: 年目標為4億元，月目標{format_platform_revenue(dynamic_values.get('platform_month_target', 0))}，目前實際數為{format_platform_revenue(dynamic_values.get('platform_actual', 0))}，月目標達成率為{format_percentage_from_text(dynamic_values.get('platform_achievement_rate_text', '0%'))}。
     累積月目標數{format_platform_revenue(dynamic_values.get('platform_cumulative_target', 0))}，累積月實際數{format_platform_revenue(dynamic_values.get('platform_cumulative_actual', 0))}，累積月目標達成率為{format_percentage_from_text(dynamic_values.get('platform_cumulative_rate_text', '0%'))}。
     榮譽累積月目標數1.63億元，榮譽累積月目標達成率為140.6%。
數位平台收益

[TABLE2_PLACEHOLDER]
"""
        
        print(f"Email content generated for {target_month}月 data")
        print(f"Using month date: {formatted_date}")
        
        # Set initial body
        mail.Body = mail_body
        
        # Display the email first
        mail.Display()
        
        # Get the mail item's inspector and word editor
        inspector = mail.GetInspector
        word_doc = inspector.WordEditor
        
        # Find and replace placeholders with actual Excel data
        # Copy first table (數位戶)
        print("Copying 數位戶 data...")
        range_obj1 = copy_excel_range_to_outlook(ws_digital_account, range1)
        if range_obj1:
            # Find the placeholder text and replace with copied range
            find_range = word_doc.Content
            find_range.Find.Text = "[TABLE1_PLACEHOLDER]"
            if find_range.Find.Execute():
                find_range.Paste()
                print("數位戶 table pasted successfully")
                
                # Make the text above the table bold
                find_range = word_doc.Content
                find_range.Find.Text = "網行銀客戶數(具有網行銀會員身分之存戶+卡戶)"
                if find_range.Find.Execute():
                    find_range.Font.Bold = True
        
        # Copy second table (數位平台收益) - copy range 11-41 and delete rows 23-31
        print("Copying 數位平台收益 data with deletion of rows 23-31...")
        print("⚠️  Converting formulas to values to prevent #REF! errors...")
        
        # Find the placeholder text first
        find_range = word_doc.Content
        find_range.Find.Text = "[TABLE2_PLACEHOLDER]"
        if find_range.Find.Execute():
            # Copy range with deletion of rows 23-31 (formulas will be converted to values)
            print("Preparing data with row deletion and formula conversion...")
            range_obj2 = copy_excel_range_with_deletion(ws_digital_platform, range2, 23, 31)
            if range_obj2:
                # Clear the placeholder text
                find_range.Text = ""
                # Paste the table data
                find_range.Paste()
                print("✅ 數位平台收益 table pasted successfully (with rows 23-31 deleted, formulas converted to values)")
                
                # Clean up temporary worksheet
                try:
                    temp_sheet_name = f"TempData_{ws_digital_platform.Name}"
                    for sheet in workbook.Worksheets:
                        if sheet.Name == temp_sheet_name:
                            excel.DisplayAlerts = False
                            sheet.Delete()
                            excel.DisplayAlerts = True
                            print("Temporary worksheet cleaned up")
                            break
                except Exception as cleanup_error:
                    print(f"Warning: Could not clean up temporary worksheet: {cleanup_error}")
            else:
                print("❌ Failed to copy 數位平台收益 table with deletion")
                # Fallback: try the original method
                print("Attempting fallback method...")
                try:
                    range_obj2_fallback = copy_excel_range_to_outlook(ws_digital_platform, f"P11:{digital_platform_end_col}22")
                    if range_obj2_fallback:
                        find_range.Text = ""
                        find_range.Paste()
                        print("Fallback: First part of table pasted")
                        
                        # Add second part
                        find_range.Collapse(0)
                        find_range.TypeText("\n\n")
                        range_obj2_fallback2 = copy_excel_range_to_outlook(ws_digital_platform, f"P32:{digital_platform_end_col}41")
                        if range_obj2_fallback2:
                            find_range.Paste()
                            print("Fallback: Second part of table pasted")
                except Exception as fallback_error:
                    print(f"Fallback method also failed: {fallback_error}")
        else:
            print("Could not find [TABLE2_PLACEHOLDER] in email content")
        
        # Make the second occurrence of "數位平台收益" text bold (skip the first one)
        print("Making the second occurrence of '數位平台收益' text bold...")
        find_range = word_doc.Content
        find_range.Find.ClearFormatting()
        find_range.Find.Text = "數位平台收益"
        
        # Search for all instances and find the second standalone one
        occurrence_count = 0
        while find_range.Find.Execute():
            # Check if this is a standalone line (not part of a longer paragraph)
            paragraph_text = find_range.Paragraphs(1).Range.Text.strip()
            if paragraph_text == "數位平台收益":
                occurrence_count += 1
                if occurrence_count == 2:  # Apply bold to the second occurrence only
                    find_range.Font.Bold = True
                    print("Applied bold formatting to the second occurrence of '數位平台收益' text")
                    break
            # Move to next occurrence
            find_range.Start = find_range.End
        
        # Apply underline formatting to text between parentheses and colon
        # Format: (1) 數位戶客戶數: and (2)數位平台收益:
        print("Applying underline formatting to section headers...")
        
        # Find and format "(1) 數位戶客戶數:"
        find_range = word_doc.Content
        find_range.Find.Text = "(1) 數位戶客戶數:"
        if find_range.Find.Execute():
            find_range.Font.Underline = True
        
        # Find and format "(2)數位平台收益:"
        find_range = word_doc.Content
        find_range.Find.Text = "(2)數位平台收益:"
        if find_range.Find.Execute():
            find_range.Font.Underline = True
        
        # Add signature from Word document
        print("Adding signature from Word document...")
        signature_path = r"C:\Users\Documents\SIGN.docx"
        
        signature_success = insert_signature_to_email(word_doc, signature_path)
        if not signature_success:
            print("⚠ Warning: Signature could not be added automatically.")
            print("Please manually add the signature content to the email.")
        
        print("✅ Email draft created successfully!")
        print("Tables have been inserted with original Excel formatting.")
        print(f"Dynamic values have been updated based on {target_month}月 data.")
        print("✅ 數位平台收益 table: formulas converted to values, rows 23-31 deleted - no #REF! errors!")
        
    except Exception as e:
        print(f"An error occurred: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        # Clean up - but don't close Excel as requested
        try:
            if workbook:
                print("Workbook remains open as requested")
                # workbook.Close(False)  # Commented out to keep Excel open
            # if excel:
            #     excel.Quit()  # Commented out to keep Excel open
        except:
            pass
    
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()