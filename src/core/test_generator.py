import sys
from pathlib import Path

# Add project root to sys.path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from src.core.code_generator import CodeGenerator
from src.core.placement_generator import PlacementResult, PlacementCommand, LineElement

def main():
    generator = CodeGenerator()
    
    # Dummy data
    mock_placement = list([
        PlacementCommand(category="text_outside", r1=1, c1=1, r2=2, c2=5, value="Test", font_size=12, font_bold=True, alignment="center")
    ])
    mock_lines = [
        LineElement(orientation="horizontal", col_start=1, col_end=5, row_start=3, row_end=3)
    ]
    
    result = PlacementResult(commands=mock_placement, warnings=[], line_elements=mock_lines)
    
    code = generator.generate(
        placement_result=result,
        grid_cols=120,
        grid_rows=150,
        col_width=0.8,
        row_height=12.0,
        page_count=1,
        output_filename="test_output.xlsx",
        pdf_name="test.pdf"
    )
    
    print("=== Generated Code ===")
    print(code)

if __name__ == "__main__":
    main()
