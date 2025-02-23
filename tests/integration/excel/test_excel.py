"""
Simple script to test Excel connectivity.
"""
import xlwings as xw
import sys

def test_excel_connection():
    print("Testing Excel connection...")
    try:
        # Try to get the active Excel application
        app = xw.apps.active
        print("✓ Successfully connected to Excel application")
        
        # List all open workbooks
        print("\nOpen workbooks:")
        for book in app.books:
            print(f"- {book.name}")
            
        # Try to get our specific workbook
        try:
            wb = app.books['NEON_ML.xlsm']
            print("\n✓ Found NEON_ML.xlsm")
            
            # List all sheets
            print("\nWorksheets in NEON_ML.xlsm:")
            for sheet in wb.sheets:
                print(f"- {sheet.name}")
                
            # Try to access specific sheet
            try:
                sheet = wb.sheets['AH NEON']
                print("\n✓ Found 'AH NEON' worksheet")
                
                # Try to read a test cell
                test_cell = sheet.range('A1').value
                print(f"\nTest read from A1: {test_cell}")
                
            except Exception as e:
                print(f"\n✗ Error accessing 'AH NEON' sheet: {str(e)}")
                
        except Exception as e:
            print(f"\n✗ Error accessing NEON_ML.xlsm: {str(e)}")
            
    except Exception as e:
        print(f"\n✗ Error connecting to Excel: {str(e)}")
        print("\nPlease ensure:")
        print("1. Excel is running")
        print("2. NEON_ML.xlsm is open")
        print("3. You have necessary permissions")
        return False
        
    return True

if __name__ == '__main__':
    success = test_excel_connection()
    sys.exit(0 if success else 1) 