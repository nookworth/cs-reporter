# CS Reporter
# Chart utilities for adding charts to PowerPoint presentations.
# Test V3 chart generation with demo data.
# **Date:** 2026-02-20
# **Status:** ✅ Completed
# **Component:** src


from src.config_v2 import load_config
from src.excel_reader_v2 import ExcelReader
from src.ppt_writer_v3 import PowerPointWriterV3

def test_v3_charts():
   
    print("Testing V3 Chart Generation...")
    print("="*60)

    current_file = "demo_data/current_month.xlsx"
    previous_file = "demo_data/previous_month.xlsx"
    
    config = load_config("config/demo_mapping_v2.yaml")
    print("[OK] Config loaded\n")
    
    reader = ExcelReader(current_file, config, previous_file)
    data = reader.extract_data()
    print(f"\n[OK] Extracted {len(data)} fields\n")
    
    print("Generating PowerPoint with charts...")
    writer = PowerPointWriterV3(config)
    output_path = writer.generate_report(data)
    
    print(f"\n[SUCCESS] Report with charts generated: {output_path}")
    print("\nCharts included:")
    print("  - Month-over-month comparison (bar chart)")
    print("  - Category distribution (pie chart)")
    print("  - Top organizations (horizontal bar chart)")

if __name__ == "__main__":
    try:
        test_v3_charts()
    except Exception as e:
        print(f"\n[ERROR] {e}")
        import traceback
        traceback.print_exc()
