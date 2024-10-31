Yo. Program is slow but this is why -
Validates the following:

1. Checks for empty Excel file
2. Validates presence of required columns:
  - Variant SKU
  - Title
  - Variant Position
  - Variant Price
  - Variant Cost
3. Price Validations
  - Validates all price fields are:
  - Present (not empty)
  - Numeric values
  - Greater than 0
  - Properly formatted (handles $ symbol)
  - Enforces price relationship hierarchy:
  - Variant Price must be greater than Trader Price
  - Trader Price must be greater than Dealer Price
  - Dealer Price must be greater than Variant Cost
4. Dealer Price Formula Validation - Basically checks if you missed VAT on top.
  - Validates Dealer Price follows the formula:
  - Dealer Price must not exceed: (Variant Price / 1.2) * 0.9
5. Image Validations
  - Proper URL format
  - Valid domain (resolvable)
  - Valid image extensions:
  - .jpg
  - .jpeg
  - Handles multiple image URLs (semicolon-separated)
  - Parses image URLs from HTML src attributes
  - Validates that images are either:
  - Exactly 825x825 pixels
  - OR maintain a 1:1 aspect ratio
  - Checks actual image dimensions by downloading and analyzing the image
6. Variant Order Validations ( i dont think it does this rn, needs tweaking )
  - Validates variant positions:
  - Must be sequential (1, 2, 3, etc.)
  - No gaps in sequence
  - Starts from position 1
7. Title Format
  - Groups products by base title (removing trailing numbers)
  - Validates title formatting:
  - First variant: Base title only
  - Subsequent variants: Base title + space + position number
  - Ensures title numbering matches position sequence
8. Inventory Validations
  - Validates 'Variant Inventory Qty':
  - Must be numeric
  - Cannot be negative
  - Handles zero inventory
9. HTML Content Validations
  - Checks HTML content for:
  - Well-formed HTML structure
  - Complete/valid anchor tags (no missing href attributes)
  - Complete/valid image tags (no missing src attributes)
  - Proper list structure (ul/ol must contain li elements)
  - Uses html5lib parser for strict validation
10. Output and Reporting - This is what is returned as categories:
  - Inventory issues
  - Price issues
  - Image issues
  - HTML issues
  - Variant Order issues
  - Other issues
  - Generates Excel report with:
  - Detailed Issues sheet (all findings)
  - Summary sheet (issue counts by category)
  - Includes SKU references for all issues
  - Timestamps output files
11. It also does progress tracking:
  - Provides real-time progress indication:
  - Current SKU being processed
  - Progress bar showing completion percentage
  - Status messages for different validation stages
  - Displays final summary in GUI message box