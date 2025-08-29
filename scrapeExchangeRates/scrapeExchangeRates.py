import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import re
from datetime import datetime
import argparse
import sys

class XRatesScraper:
    def __init__(self, base_url="https://www.x-rates.com/average/"):
        self.base_url = base_url
        self.session = requests.Session()
        # Add headers to avoid being blocked
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Accept-Encoding': 'gzip, deflate',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
        })
    
    def get_year_data(self, year, from_currency="USD", to_currency="INR", amount=1):
        """
        Fetch exchange rate data for a specific year
        """
        url = f"{self.base_url}?from={from_currency}&to={to_currency}&amount={amount}&year={year}"
        
        try:
            print(f"Fetching data for year {year}...")
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            
            # Wait a bit to be respectful to the server
            time.sleep(1)
            
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Find the OutputLinksAvg class
            output_links = soup.find('ul', class_='OutputLinksAvg')
            
            if not output_links:
                print(f"Warning: Could not find OutputLinksAvg class for year {year}")
                return None
            
            # Wait a bit more to ensure all content is loaded
            time.sleep(0.5)
            
            year_data = {}
            
            # Parse each month's data
            for li in output_links.find_all('li'):
                month_span = li.find('span', class_='avgMonth')
                rate_span = li.find('span', class_='avgRate')
                
                if month_span and rate_span:
                    month = month_span.text.strip()
                    rate_text = rate_span.text.strip()
                    
                    # Extract numeric value from rate
                    rate_match = re.search(r'[\d.]+', rate_text)
                    if rate_match:
                        rate = float(rate_match.group())
                        year_data[month] = rate
            
            print(f"Successfully fetched {len(year_data)} months for year {year}")
            return year_data
            
        except requests.exceptions.RequestException as e:
            print(f"Error fetching data for year {year}: {e}")
            return None
        except Exception as e:
            print(f"Error parsing data for year {year}: {e}")
            return None
    
    def scrape_multiple_years(self, start_year, end_year=None, from_currency="USD", to_currency="INR"):
        """
        Scrape data for multiple years and return as DataFrame
        """
        if end_year is None:
            end_year = datetime.now().year
        
        all_data = {}
        
        for year in range(start_year, end_year + 1):
            year_data = self.get_year_data(year, from_currency, to_currency)
            
            if year_data:
                # Add year data to the main dictionary
                for month, rate in year_data.items():
                    if month not in all_data:
                        all_data[month] = {}
                    all_data[month][year] = rate
            
            # Be respectful to the server
            time.sleep(2)
        
        # Convert to DataFrame
        df = pd.DataFrame(all_data).T
        
        # Reorder months properly
        month_order = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                      'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        
        # Only include months that exist in the data
        existing_months = [month for month in month_order if month in df.index]
        df = df.reindex(existing_months)
        
        # Sort columns (years) in ascending order
        df = df.sort_index(axis=1)
        
        return df
    
    def save_to_excel(self, df, filename=None, from_currency="USD", to_currency="INR"):
        """
        Save DataFrame to Excel file
        """
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"exchange_rates_{from_currency}_to_{to_currency}_{timestamp}.xlsx"
        
        try:
            # Create Excel writer with formatting
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                # Write the main data
                df.to_excel(writer, sheet_name='Exchange Rates', index_label='Month')
                
                # Get the workbook and worksheet
                workbook = writer.book
                worksheet = writer.sheets['Exchange Rates']
                
                # Add some basic formatting
                from openpyxl.styles import Font, PatternFill, Alignment
                
                # Header formatting
                header_font = Font(bold=True, color="FFFFFF")
                header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                
                # Format header row
                for cell in worksheet[1]:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal="center")
                
                # Format month column
                for row in range(2, len(df) + 2):
                    cell = worksheet[f'A{row}']
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal="center")
                
                # Auto-adjust column widths
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 15)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            print(f"Data saved successfully to {filename}")
            return filename
            
        except Exception as e:
            print(f"Error saving to Excel: {e}")
            return None

def main():
    parser = argparse.ArgumentParser(description='Scrape exchange rates from x-rates.com')
    parser.add_argument('--start-year', type=int, default=2015, 
                       help='Start year for data collection (default: 2015)')
    parser.add_argument('--end-year', type=int, default=None,
                       help='End year for data collection (default: current year)')
    parser.add_argument('--from-currency', default='USD',
                       help='Source currency (default: USD)')
    parser.add_argument('--to-currency', default='INR',
                       help='Target currency (default: INR)')
    parser.add_argument('--output', default=None,
                       help='Output filename (default: auto-generated)')
    parser.add_argument('--current-year-only', action='store_true',
                       help='Fetch only current year data')
    
    args = parser.parse_args()
    
    # Initialize scraper
    scraper = XRatesScraper()
    
    if args.current_year_only:
        current_year = datetime.now().year
        print(f"Fetching data for current year only: {current_year}")
        df = scraper.scrape_multiple_years(current_year, current_year, 
                                         args.from_currency, args.to_currency)
    else:
        end_year = args.end_year or datetime.now().year
        print(f"Fetching data from {args.start_year} to {end_year}")
        df = scraper.scrape_multiple_years(args.start_year, end_year, 
                                         args.from_currency, args.to_currency)
    
    if df is not None and not df.empty:
        print(f"\nData Summary:")
        print(f"Shape: {df.shape}")
        print(f"Years: {list(df.columns)}")
        print(f"Months: {list(df.index)}")
        
        # Display preview
        print(f"\nPreview of data:")
        print(df.head())
        
        # Save to Excel
        filename = scraper.save_to_excel(df, args.output, args.from_currency, args.to_currency)
        
        if filename:
            print(f"\n✓ Successfully saved exchange rate data to: {filename}")
        else:
            print("\n✗ Failed to save data to Excel")
    else:
        print("\n✗ No data was collected. Please check the website and try again.")

if __name__ == "__main__":    
    main()
