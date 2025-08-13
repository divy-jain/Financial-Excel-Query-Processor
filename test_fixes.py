#!/usr/bin/env python3
"""
Test script to verify the specific fixes implemented
"""

from excel_query_system import process_file, excel_query

def test_specific_fixes():
    print("🔧 Testing Specific Fixes")
    print("=" * 40)
    
    # Test the Excel file processing
    try:
        print("📁 Processing Excel file...")
        file_rep = process_file("Consolidated Plan 2023-2024.xlsm")
        print("✅ File processed successfully")
        
        # Test Query 4: Percentage calculation fix
        print("\n📝 Testing Query 4: Percentage calculation fix")
        query4 = "What percent of MXDs costs are indirect? Which month had the highest percentage?"
        answer4 = excel_query(query4, file_rep)
        print(f"Query: {query4}")
        print(f"Answer: {answer4}")
        
        # Test Query 6: Monthly granularity fix
        print("\n📝 Testing Query 6: Monthly granularity fix")
        query6 = "What is Branch's advertising forecasts for each month in 2024?"
        answer6 = excel_query(query6, file_rep)
        print(f"Query: {query6}")
        print(f"Answer: {answer6}")
        
        print("\n✅ Test completed successfully!")
        
    except Exception as e:
        print(f"❌ Error during testing: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_specific_fixes()
