import pandas as pd
from collections import defaultdict
import re
from urllib.parse import urlparse
import os
import socket
import ssl
import requests
from requests.exceptions import RequestException
import tkinter as tk
from tkinter import ttk  # Added ttk import
from tkinter import filedialog, messagebox
import logging
from datetime import datetime
from PIL import Image
from io import BytesIO
from bs4 import BeautifulSoup
import html5lib

def is_valid_domain(domain):
    """Check if domain exists and is resolvable"""
    try:
        socket.gethostbyname(domain)
        return True
    except socket.gaierror:
        return False

def is_valid_url(url):
    """Validate URL format, domain existence, and image extension"""
    try:
        result = urlparse(url.strip())
        
        if not all([result.scheme, result.netloc]):
            return False
            
        domain = result.netloc
        if not is_valid_domain(domain):
            return False
        
        valid_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.webp']
        path_lower = result.path.lower()
        if not any(path_lower.endswith(ext) for ext in valid_extensions):
            return False
        
        return True, ""
    except Exception as e:
        return False, str(e)

def validate_image_dimensions(url):
    """Validate image dimensions are either 825x825 or maintain 1:1 ratio"""
    try:
        response = requests.get(url, timeout=10)
        img = Image.open(BytesIO(response.content))
        width, height = img.size
        
        # Check if exactly 825x825
        if width == 825 and height == 825:
            return True, ""
            
        # Check if maintains 1:1 ratio
        if width != height:
            return False, f"Image dimensions {width}x{height} do not maintain 1:1 ratio"
            
        return True, ""
    except Exception as e:
        return False, f"Error validating image dimensions: {str(e)}"

def validate_image_urls(urls_str):
    """Validate multiple image URLs separated by semicolons"""
    if not urls_str or pd.isna(urls_str):
        return False, "Missing image URL"
    
    urls = []
    for url_part in str(urls_str).split(';'):
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
    """
    Validate price relationships:
    Variant Price > Trader Price > Dealer Price > Variant Cost
    """
    try:
        # Validate and convert all prices
        price_valid, variant_price = validate_price(row['Variant Price'])
        if not price_valid:
            return False, f"Invalid Variant Price: {variant_price}"
            
        trader_price_valid, trader_price = validate_price(
            row.get('Variant Metafield:product.trader-price [single_line_text_field]', 0))
        if not trader_price_valid and row.get('Variant Metafield:product.trader-price [single_line_text_field]'):
            return False, f"Invalid Trader Price: {trader_price}"
            
        dealer_price_valid, dealer_price = validate_price(
            row.get('Variant Metafield:product.dealer-price [single_line_text_field]', 0))
        if not dealer_price_valid and row.get('Variant Metafield:product.dealer-price [single_line_text_field]'):
            return False, f"Invalid Dealer Price: {dealer_price}"
            
        cost_valid, variant_cost = validate_price(row['Variant Cost'])
        if not cost_valid:
            return False, f"Invalid Variant Cost: {variant_cost}"

        issues = []
        
        # Check price hierarchy
        if trader_price_valid and variant_price <= trader_price:
            issues.append("Variant Price must be greater than Trader Price")
        if trader_price_valid and dealer_price_valid and trader_price <= dealer_price:
            issues.append("Trader Price must be greater than Dealer Price")
        if dealer_price_valid and dealer_price <= variant_cost:
            issues.append("Dealer Price must be greater than Variant Cost")
            
        # Check dealer price formula
        if dealer_price_valid:
            required_dealer_price = variant_price / 1.2 * 0.9
            if dealer_price > required_dealer_price:
                issues.append(f"Dealer Price Issue: Not less than {required_dealer_price:.2f} (Variant Price/1.2*0.9)")
            
        return len(issues) == 0, "; ".join(issues)
    except Exception as e:
        return False, f"Price validation error: {str(e)}"

def validate_inventory(row):
    """Validate inventory quantities are non-negative"""
    try:
        inventory = float(row.get('Variant Inventory Qty', 0))
        if inventory < 0:
            return False, f"Negative inventory quantity found: {inventory}"
        return True, ""
    except ValueError:
        return False, "Invalid inventory quantity format"

def validate_html_content(html_content):
    """Validate HTML content for well-formedness and basic structure"""
    if not html_content or pd.isna(html_content):
        return False, "Missing HTML content"
        
    try:
        # Parse HTML using html5lib for strict validation
        soup = BeautifulSoup(html_content, 'html5lib')
        
        # Check for common issues
        issues = []
        
        # Check for broken links
        broken_links = [a for a in soup.find_all('a') if not a.get('href')]
        if broken_links:
            issues.append("Found links without href attributes")
            
        # Check for broken images
        broken_images = [img for img in soup.find_all('img') if not img.get('src')]
        if broken_images:
            issues.append("Found images without src attributes")
            
        # Check for invalid nested lists
        invalid_lists = soup.find_all(['ul', 'ol'], recursive=False)
        for lst in invalid_lists:
            if not lst.find_all('li', recursive=False):
                issues.append(f"Found {lst.name} without li elements")
                
        return len(issues) == 0, "; ".join(issues)
    except Exception as e:
        return False, f"Invalid HTML content: {str(e)}"

def validate_variant_order(df):
    """Validate variant positions and title ordering"""
    issues = []
    
    # Group by base title (removing any trailing numbers)
    df['Base Title'] = df['Title'].str.replace(r'\s+\d+$', '', regex=True)
    grouped = df.groupby('Base Title')
    
    for base_title, group in grouped:
        if len(group) > 1:
            # Sort by Variant Position
            sorted_group = group.sort_values('Variant Position')
            
            # Check position sequence
            positions = sorted_group['Variant Position'].tolist()
            expected_positions = list(range(1, len(positions) + 1))
            
            if sorted(positions) != expected_positions:
                for _, row in sorted_group.iterrows():
                    issues.append({
                        'Variant SKU': row['Variant SKU'],
                        'Message': f'Incorrect position sequence. Expected {expected_positions}, got {positions}'
                    })
            
            # Check title numbering
            for i, (_, row) in enumerate(sorted_group.iterrows(), 1):
                expected_title = f"{base_title} {i}" if i > 1 else base_title
                if row['Title'] != expected_title:
                    issues.append({
                        'Variant SKU': row['Variant SKU'],
                        'Message': f'Incorrect title format. Expected "{expected_title}", got "{row["Title"]}"'
                    })
    
    return issues

class ProgressWindow:
    def __init__(self, title="Progress"):
        self.root = tk.Tk()
        self.root.title(title)
        self.root.geometry("400x150")
        
        # Center the window
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width/2) - (400/2)
        y = (screen_height/2) - (150/2)
        self.root.geometry(f'+{int(x)}+{int(y)}')
        
        # Progress label
        self.label = tk.Label(self.root, text="Starting validation...", padx=20, pady=10)
        self.label.pack()
        
        # Progress bar
        self.progress = tk.ttk.Progressbar(self.root, length=300, mode='determinate')
        self.progress.pack(padx=20, pady=10)
        
        # Status label
        self.status = tk.Label(self.root, text="", padx=20)
        self.status.pack()
        
        self.root.update()
    
    def update_progress(self, value, total, message):
        progress = (value / total) * 100
        self.progress['value'] = progress
        self.label['text'] = f"Processing... ({value}/{total})"
        self.status['text'] = message
        self.root.update()
    
    def close(self):
        self.root.destroy()

def validate_product_data(file_path):
    """Enhanced product data validation with progress reporting"""
    progress_window = ProgressWindow()
    
    try:
        progress_window.update_progress(0, 100, "Loading Excel file...")
        df = pd.read_excel(file_path)
        
        if df.empty:
            messagebox.showerror("Error", "The Excel file is empty!")
            progress_window.close()
            return []
            
        required_columns = ['Variant SKU', 'Title', 'Variant Position', 'Variant Price', 'Variant Cost']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            messagebox.showerror("Error", f"Missing required columns: {', '.join(missing_columns)}")
            progress_window.close()
            return []
        
        issues = []
        total_rows = len(df)
        
        # Process variant order validation first
        progress_window.update_progress(0, total_rows, "Validating variant ordering...")
        variant_order_issues = validate_variant_order(df)
        issues.extend(variant_order_issues)
        
        # Process row-by-row validations
        for index, row in df.iterrows():
            sku = row['Variant SKU']
            current_message = f"Processing SKU: {sku}"
            progress_window.update_progress(index + 1, total_rows, current_message)
            
            # 1. Inventory validation
            inv_valid, inv_message = validate_inventory(row)
            if not inv_valid:
                issues.append({
                    'Variant SKU': sku,
                    'Message': f'Inventory issue: {inv_message}'
                })
            
            # 2. Price hierarchy validation
            price_valid, price_message = validate_price_hierarchy(row)
            if not price_valid:
                issues.append({
                    'Variant SKU': sku,
                    'Message': f'Price hierarchy issue: {price_message}'
                })
            
            # 3. Image validation
            if 'Image Src' in df.columns and not pd.isna(row['Image Src']):
                img_valid, img_message = validate_image_urls(row['Image Src'])
                if not img_valid:
                    issues.append({
                        'Variant SKU': sku,
                        'Message': f'Image issue: {img_message}'
                    })
            
            # 4. HTML validation
            if 'Body HTML' in df.columns and not pd.isna(row['Body HTML']):
                html_valid, html_message = validate_html_content(row['Body HTML'])
                if not html_valid:
                    issues.append({
                        'Variant SKU': sku,
                        'Message': f'HTML content issue: {html_message}'
                    })
        
        progress_window.close()
        return issues
        
    except Exception as e:
        progress_window.close()
        error_message = f"Error processing file:\n{str(e)}"
        messagebox.showerror("Error", error_message)
        return []

def main():
    """Enhanced main function with error handling"""
    root = tk.Tk()
    root.withdraw()
    
    try:
        file_path = filedialog.askopenfilename(
            title="Select Excel File to Validate",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
            
        if not os.path.exists(file_path):
            messagebox.showerror("Error", "Selected file does not exist!")
            return
            
        print(f"Processing file: {file_path}")  # Debug output
        
        validation_issues = validate_product_data(file_path)
        
        if validation_issues is None:
            return
            
        if validation_issues:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(
                os.path.dirname(file_path),
                f'validation_issues_{timestamp}.xlsx'
            )
            
            # Create DataFrame with categorized issues
            issues_df = pd.DataFrame(validation_issues)
            
            # Add category column
            issues_df['Category'] = issues_df['Message'].apply(lambda x: 
                'Inventory' if 'inventory' in x.lower() else
                'Price' if 'price' in x.lower() else
                'Image' if 'image' in x.lower() else
                'HTML' if 'html' in x.lower() else
                'Variant Order' if 'position' in x.lower() or 'title' in x.lower() else
                'Other'
            )
            
            try:
                # Save to Excel with summary sheet
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    issues_df.to_excel(writer, sheet_name='Detailed Issues', index=False)
                    
                    # Create summary sheet
                    summary = issues_df['Category'].value_counts().reset_index()
                    summary.columns = ['Category', 'Count']
                    summary.to_excel(writer, sheet_name='Summary', index=False)
                
                # Show detailed message box
                summary_text = "\n".join([
                    f"- {row['Category']}: {row['Count']} issues"
                    for _, row in summary.iterrows()
                ])
                
                message = (
                    f"Found {len(validation_issues)} total issues:\n\n"
                    f"{summary_text}\n\n"
                    f"Results saved to:\n{output_file}"
                )
                
                messagebox.showinfo("Validation Complete", message)
            except Exception as e:
                messagebox.showerror("Error", f"Error saving results: {str(e)}")
        else:
            messagebox.showinfo("Validation Complete", "No issues found!")
            
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        print(f"Error details: {str(e)}")  # Debug output

if __name__ == "__main__":
    main()