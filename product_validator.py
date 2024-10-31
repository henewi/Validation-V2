import pandas as pd
import os
from datetime import datetime
import requests
from PIL import Image
from io import BytesIO
from bs4 import BeautifulSoup
import html5lib
from urllib.parse import urlparse
import socket
import logging
import re

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def is_valid_url(url):
    """Validate URL format and domain existence"""
    try:
        result = urlparse(url.strip())
        if not all([result.scheme, result.netloc]):
            return False, "Invalid URL format"
            
        # Basic domain check
        try:
            socket.gethostbyname(result.netloc)
        except socket.gaierror:
            return False, "Domain not resolvable"
            
        valid_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.webp']
        path_lower = result.path.lower()
        if not any(path_lower.endswith(ext) for ext in valid_extensions):
            return False, "Invalid image extension"
            
        return True, ""
    except Exception as e:
        return False, str(e)

def validate_image_dimensions(url, timeout=10):
    """Validate image dimensions are either 825x825 or maintain 1:1 ratio"""
    try:
        response = requests.get(url, timeout=timeout)
        img = Image.open(BytesIO(response.content))
        width, height = img.size
        
        if width == 825 and height == 825:
            return True, ""
            
        if width != height:
            return False, f"Image dimensions {width}x{height} do not maintain 1:1 ratio"
            
        return True, ""
    except requests.exceptions.Timeout:
        return False, "Request timed out while fetching image"
    except requests.exceptions.RequestException as e:
        return False, f"Error fetching image: {str(e)}"
    except Exception as e:
        return False, f"Error validating image dimensions: {str(e)}"

def validate_image_urls(urls_str):
    """Validate multiple image URLs separated by semicolons"""
    if not urls_str or pd.isna(urls_str):
        return False, "Missing image URL"
    
    urls = []
    for url_part in str(urls_str).split(';'):
        # Extract URLs from possible HTML src attributes
        matches = re.findall(r'(?:src=[\'"])(.*?)(?:[\'"])', url_part)
        if matches:
            urls.extend(matches)
        else:
            urls.append(url_part.strip())
    
    urls = [url for url in urls if url.strip()]
    if not urls:
        return False, "No valid URLs found"
    
    issues = []
    for url in urls:
        # Check URL validity
        url_valid, url_message = is_valid_url(url)
        if not url_valid:
            issues.append(f"Invalid URL {url}: {url_message}")
            continue
            
        # Check dimensions
        dim_valid, dim_message = validate_image_dimensions(url)
        if not dim_valid:
            issues.append(f"Invalid dimensions for {url}: {dim_message}")
    
    if issues:
        return False, "; ".join(issues)
    return True, ""

def validate_price(value):
    """Validate price/cost values"""
    if pd.isna(value):
        return False, "Missing value"
    try:
        float_val = float(str(value).replace('$', '').strip())
        if float_val <= 0:
            return False, "Value must be greater than 0"
        return True, float_val
    except ValueError:
        return False, "Invalid numeric value"

def validate_price_hierarchy(row):
    """Validate price relationships"""
    try:
        price_valid, variant_price = validate_price(row['Variant Price'])
        if not price_valid:
            return False, f"Invalid Variant Price: {variant_price}"
            
        cost_valid, variant_cost = validate_price(row['Variant Cost'])
        if not cost_valid:
            return False, f"Invalid Variant Cost: {variant_cost}"

        # Optional price fields
        trader_price = row.get('Variant Metafield:product.trader-price [single_line_text_field]')
        dealer_price = row.get('Variant Metafield:product.dealer-price [single_line_text_field]')
        
        issues = []
        
        if trader_price:
            trader_valid, trader_price = validate_price(trader_price)
            if trader_valid and variant_price <= trader_price:
                issues.append("Variant Price must be greater than Trader Price")
                
        if dealer_price:
            dealer_valid, dealer_price = validate_price(dealer_price)
            if dealer_valid:
                if trader_price and trader_valid and trader_price <= dealer_price:
                    issues.append("Trader Price must be greater than Dealer Price")
                if dealer_price <= variant_cost:
                    issues.append("Dealer Price must be greater than Variant Cost")
                    
        return len(issues) == 0, "; ".join(issues) if issues else ""
        
    except Exception as e:
        return False, f"Price validation error: {str(e)}"

def validate_product_data(file_path):
    """Validate product data and return issues"""
    logger.info(f"Starting validation of file: {file_path}")
    
    try:
        # Read Excel file
        df = pd.read_excel(file_path)
        logger.info(f"Successfully loaded Excel file with {len(df)} rows")
        
        if df.empty:
            logger.error("The Excel file is empty")
            return []
            
        # Check required columns
        required_columns = ['Variant SKU', 'Title', 'Variant Position', 'Variant Price', 'Variant Cost']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            logger.error(f"Missing required columns: {missing_columns}")
            return []
        
        issues = []
        total_rows = len(df)
        
        # Process row-by-row validations
        for index, row in df.iterrows():
            sku = row['Variant SKU']
            logger.info(f"Processing row {index + 1}/{total_rows}, SKU: {sku}")
            
            # Price hierarchy validation
            price_valid, price_message = validate_price_hierarchy(row)
            if not price_valid:
                issues.append({
                    'Variant SKU': sku,
                    'Message': f'Price hierarchy issue: {price_message}'
                })
            
            # Image validation if present
            if 'Image Src' in df.columns and not pd.isna(row['Image Src']):
                img_valid, img_message = validate_image_urls(row['Image Src'])
                if not img_valid:
                    issues.append({
                        'Variant SKU': sku,
                        'Message': f'Image issue: {img_message}'
                    })
        
        return issues
        
    except Exception as e:
        logger.error(f"Error processing file: {str(e)}")
        return []

def save_validation_results(validation_issues, input_file_path):
    """Save validation results to Excel file"""
    if not validation_issues:
        logger.info("No validation issues to save")
        return None
        
    try:
        # Create output filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.dirname(input_file_path)
        output_file = os.path.join(
            output_dir,
            f'validation_issues_{timestamp}.xlsx'
        )
        
        # Create DataFrame
        issues_df = pd.DataFrame(validation_issues)
        
        # Add categories
        issues_df['Category'] = issues_df['Message'].apply(lambda x: 
            'Price' if 'price' in x.lower() else
            'Image' if 'image' in x.lower() else
            'Other'
        )
        
        # Create summary
        summary = issues_df['Category'].value_counts().reset_index()
        summary.columns = ['Category', 'Count']
        
        logger.info(f"Writing results to {output_file}")
        
        # Save to Excel
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            issues_df.to_excel(writer, sheet_name='Detailed Issues', index=False)
            summary.to_excel(writer, sheet_name='Summary', index=False)
            
        logger.info(f"Results successfully saved to: {output_file}")
        return output_file
        
    except Exception as e:
        logger.error(f"Error saving results: {str(e)}")
        return None

def main():
    """Main function"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Validate product data file')
    parser.add_argument('input_file', help='Path to the Excel file to validate')
    args = parser.parse_args()
    
    if not os.path.exists(args.input_file):
        logger.error(f"Input file not found: {args.input_file}")
        return
    
    # Run validation
    validation_issues = validate_product_data(args.input_file)
    
    # Save results if there are issues
    if validation_issues:
        output_file = save_validation_results(validation_issues, args.input_file)
        if output_file:
            print(f"\nValidation complete. Found {len(validation_issues)} issues.")
            print(f"Results saved to: {output_file}")
    else:
        print("\nValidation complete. No issues found!")

if __name__ == "__main__":
    main()