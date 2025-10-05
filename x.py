"""
Redfin Property Scraper - Interactive User Control
Allows user to select all filters manually
"""

import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import os
import subprocess
from datetime import datetime

class RedfinScraperInteractive:
    def __init__(self, excel_file="redfin_properties.xlsx"):
        self.excel_file = excel_file
        self.driver = None
        self.base_url = "https://www.redfin.com/county/1974/NY/Nassau-County"
        self.listing_status = None
    
    def kill_chrome_processes(self):
        """Kill any existing Chrome/ChromeDriver processes"""
        try:
            # Kill chromedriver
            subprocess.run(['taskkill', '/F', '/IM', 'chromedriver.exe'], 
                         stderr=subprocess.DEVNULL, stdout=subprocess.DEVNULL)
            time.sleep(1)
            print("Closed existing ChromeDriver processes")
        except:
            pass
        
    def setup_driver(self):
        """Initialize Chrome driver"""
        chrome_options = Options()
        chrome_options.add_argument('--start-maximized')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_argument('--no-first-run')
        chrome_options.add_argument('--no-service-autorun')
        chrome_options.add_argument('--password-store=basic')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--remote-debugging-port=9222')
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # Add user agent to avoid detection
        chrome_options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        
        self.driver = webdriver.Chrome(options=chrome_options)
        self.wait = WebDriverWait(self.driver, 20)
        
        # Execute CDP commands to avoid detection
        self.driver.execute_cdp_cmd('Network.setUserAgentOverride', {
            "userAgent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        })
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        print("Browser opened")
        
    def start_and_wait_for_user(self):
        """Open browser and let user apply filters manually"""
        print("\n" + "="*60)
        print("REDFIN SCRAPER - MANUAL FILTER SELECTION")
        print("="*60)
        
        self.driver.get(self.base_url)
        print(f"\nOpened: {self.base_url}")
        
        print("\n" + "-"*60)
        print("INSTRUCTIONS:")
        print("-"*60)
        print("1. Click on 'Filters' button")
        print("2. Select 'Sold' or 'For Sale'")
        print("3. If Sold: Choose time period (Last week, month, 3 months, etc.)")
        print("4. Select property types (House, Townhouse, Multi-family)")
        print("5. Select 'Time on Redfin' if needed")
        print("6. Click 'See X homes' or 'Apply' button")
        print("7. Wait for results to load")
        print("-"*60)
        
        input("\nPress ENTER when you have applied all filters and can see the property listings...")
        
        # Ask user what they selected
        print("\n" + "-"*60)
        status = input("Did you select 'Sold' or 'For Sale'? (sold/sale): ").strip().lower()
        self.listing_status = 'sold' if status == 'sold' else 'for-sale'
        print(f"Listing Status: {self.listing_status}")
        
        # Ask about starting page
        start_page_input = input("\nWhat page should we start from? (default: 1): ").strip()
        start_page = int(start_page_input) if start_page_input else 1
        
        if start_page > 1:
            print(f"\nNavigating to page {start_page}...")
            # Get current URL and modify it
            current_url = self.driver.current_url
            if '/page-' in current_url:
                # Replace existing page number
                base_url = current_url.split('/page-')[0]
                new_url = f"{base_url}/page-{start_page}"
            else:
                new_url = f"{current_url}/page-{start_page}"
            
            self.driver.get(new_url)
            time.sleep(3)
            print(f"On page {start_page}")
            input("\nPress ENTER to start scraping from this page...")
        
        print("\n" + "="*60)
        print("STARTING SCRAPING...")
        print("="*60 + "\n")
    
    def close_popup_if_exists(self):
        """Close any popup that appears"""
        try:
            close_button = WebDriverWait(self.driver, 3).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.bp-CloseButton"))
            )
            close_button.click()
            time.sleep(1)
            print("✓ Popup closed")
            return True
        except:
            return False
    
    def extract_property_details(self, property_url):
        """Extract detailed property information from property page"""
        # Open in new tab with the URL directly
        self.driver.execute_script(f"window.open('{property_url}', '_blank');")
        time.sleep(2)
        self.driver.switch_to.window(self.driver.window_handles[-1])
        
        property_data = {
            'url': property_url,
            'listing_status': self.listing_status,
            'scrape_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        try:
            # Navigate if not already on the page
            if self.driver.current_url != property_url:
                self.driver.get(property_url)
                time.sleep(3)
            
            # Wait for page to load properly
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            time.sleep(2)
            
            # Close popup if exists
            self.close_popup_if_exists()
            
            # Get sold date (only for sold properties)
            if self.listing_status == 'sold':
                try:
                    sold_banner = WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'div.ListingStatusBannerSection'))
                    )
                    sold_text = sold_banner.text
                    if 'SOLD ON' in sold_text:
                        property_data['sold_date'] = sold_text.split('SOLD ON')[1].strip()
                    else:
                        property_data['sold_date'] = sold_text.replace('SOLD', '').strip()
                except:
                    property_data['sold_date'] = 'N/A'
            else:
                property_data['sold_date'] = 'N/A'
            
            # Get address - try multiple selectors
            try:
                # Try main address
                address_elem = self.driver.find_element(By.CSS_SELECTOR, 'h1.full-address')
                street = address_elem.text.split(',')[0].strip()
                property_data['street_address'] = street
                
                # Get city, state, zip
                city_state_zip_elem = self.driver.find_element(By.CSS_SELECTOR, 'span.bp-cityStateZip')
                city_state_zip = city_state_zip_elem.text
                
                parts = city_state_zip.split(',')
                property_data['city'] = parts[0].strip() if len(parts) > 0 else ''
                
                if len(parts) > 1:
                    state_zip = parts[1].strip().split()
                    property_data['state'] = state_zip[0] if len(state_zip) > 0 else ''
                    property_data['zip_code'] = state_zip[1] if len(state_zip) > 1 else ''
                else:
                    property_data['state'] = ''
                    property_data['zip_code'] = ''
                    
                property_data['full_address'] = f"{street}, {city_state_zip}"
            except Exception as e:
                print(f"  ⚠ Error extracting address: {e}")
                property_data['street_address'] = 'N/A'
                property_data['city'] = 'N/A'
                property_data['state'] = 'N/A'
                property_data['zip_code'] = 'N/A'
                property_data['full_address'] = 'N/A'
            
            # Get price
            try:
                price_elem = self.driver.find_element(By.CSS_SELECTOR, 'div.statsValue')
                property_data['price'] = price_elem.text.strip()
            except:
                try:
                    price_elem = self.driver.find_element(By.CSS_SELECTOR, 'div.price')
                    property_data['price'] = price_elem.text.strip()
                except:
                    property_data['price'] = 'N/A'
            
            # Get beds
            try:
                beds_elem = self.driver.find_element(By.CSS_SELECTOR, 'div.beds-section .statsValue')
                property_data['beds'] = beds_elem.text.strip()
            except:
                property_data['beds'] = 'N/A'
            
            # Get baths
            try:
                baths_elem = self.driver.find_element(By.CSS_SELECTOR, 'div.baths-section .statsValue')
                property_data['baths'] = baths_elem.text.strip()
            except:
                property_data['baths'] = 'N/A'
            
            # Get sqft
            try:
                sqft_elem = self.driver.find_element(By.CSS_SELECTOR, 'div.sqft-section .statsValue')
                property_data['sqft'] = sqft_elem.text.strip()
            except:
                property_data['sqft'] = 'N/A'
            
            # Get property type
            try:
                # Look for property type in key details
                prop_type_elem = self.driver.find_element(By.XPATH, '//span[text()="Property Type"]/preceding-sibling::span[@class="valueText"]')
                property_data['property_type'] = prop_type_elem.text.strip()
            except:
                property_data['property_type'] = 'N/A'
            
            # Scroll down to find Interior section
            try:
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight/2);")
                time.sleep(1)
            except:
                pass
            
            # Find and click Interior section
            property_data['has_oil_heating'] = 'No'
            property_data['heating_type'] = 'N/A'
            property_data['cooling_type'] = 'N/A'
            
            try:
                # Scroll to property details section first
                try:
                    details_section = self.driver.find_element(By.ID, 'property-details-scroll')
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'start'});", details_section)
                    time.sleep(1.5)
                except:
                    pass
                
                # Find the clickable Interior header
                interior_header = None
                
                try:
                    interior_header = self.driver.find_element(By.XPATH, 
                        '//svg[contains(@class, "lightbulb-shine")]/ancestor::div[@class="sectionHeaderContainer"]')
                    print("  ✓ Found Interior header (lightbulb)")
                except:
                    try:
                        interior_header = self.driver.find_element(By.XPATH, 
                            '//h3[contains(., "Interior")]/ancestor::div[@class="sectionHeaderContainer"]')
                        print("  ✓ Found Interior header (h3)")
                    except Exception as e:
                        print(f"  ✗ Could not find Interior header: {e}")
                
                if interior_header:
                    # Scroll to it
                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", interior_header)
                    time.sleep(1)
                    
                    # Check if collapsed
                    parent_expandable = interior_header.find_element(By.XPATH, './ancestor::div[contains(@class, "expandableSection")]')
                    parent_classes = parent_expandable.get_attribute('class')
                    
                    print(f"  ℹ Parent classes: {parent_classes}")
                    
                    # ALWAYS CLICK regardless of state - it might be showing wrong state
                    print("  → Force clicking Interior section...")
                    try:
                        interior_header.click()
                        time.sleep(3)
                        print("  ✓ Clicked Interior section")
                    except Exception as e:
                        print(f"  ⚠ Click failed, trying JavaScript click: {e}")
                        self.driver.execute_script("arguments[0].click();", interior_header)
                        time.sleep(3)
                        print("  ✓ JavaScript clicked Interior section")
                    
                    # Wait for content
                    time.sleep(2)
                    
                    # Debug: Check if "Heating" text exists anywhere on page
                    page_source = self.driver.page_source
                    page_text = self.driver.find_element(By.TAG_NAME, 'body').text
                    
                    print(f"  ℹ 'Heating' in page source: {'Heating' in page_source}")
                    print(f"  ℹ 'Heating:' in page text: {'Heating:' in page_text}")
                    print(f"  ℹ 'Cooling' in page source: {'Cooling' in page_source}")
                    
                    # Try to find heating section in HTML
                    heating_found = False
                    
                    # Method 1: Direct HTML search
                    if 'Heating:' in page_source or 'Heating &amp;' in page_source:
                        print("  ℹ Found Heating in HTML source")
                        
                        # Try to get all li elements with entryItem class
                        try:
                            all_items = self.driver.find_elements(By.CSS_SELECTOR, 'li.entryItem')
                            print(f"  ℹ Found {len(all_items)} li.entryItem elements")
                            
                            for item in all_items:
                                item_text = item.text
                                # print(f"    - Item text: {item_text[:50]}...")  # Debug each item
                                
                                if 'Heating:' in item_text or 'Heating :' in item_text:
                                    print(f"  ✓ Found heating item: {item_text}")
                                    heating_text = item_text.split('Heating')[1].replace(':', '').strip()
                                    
                                    if heating_text:
                                        property_data['heating_type'] = heating_text
                                        heating_found = True
                                        
                                        # Check for Oil (case insensitive)
                                        if 'oil' in heating_text.lower():
                                            property_data['has_oil_heating'] = 'Yes'
                                            print(f"  🔥 OIL HEATING FOUND: {heating_text}")
                                        else:
                                            print(f"  ℹ Heating type: {heating_text} (No oil)")
                                        break
                                        
                                if 'Cooling:' in item_text or 'Cooling :' in item_text:
                                    cooling_text = item_text.split('Cooling')[1].replace(':', '').strip()
                                    if cooling_text:
                                        property_data['cooling_type'] = cooling_text
                        except Exception as e:
                            print(f"  ⚠ Error iterating items: {e}")
                    
                    # Method 2: Text-based fallback
                    if not heating_found and 'Heating:' in page_text:
                        print("  ℹ Trying text-based extraction...")
                        lines = page_text.split('\n')
                        for i, line in enumerate(lines):
                            if 'Heating:' in line:
                                heating_text = line.replace('Heating:', '').strip()
                                
                                if not heating_text and i + 1 < len(lines):
                                    heating_text = lines[i + 1].strip()
                                
                                if heating_text:
                                    property_data['heating_type'] = heating_text
                                    heating_found = True
                                    
                                    if 'oil' in heating_text.lower():
                                        property_data['has_oil_heating'] = 'Yes'
                                        print(f"  🔥 OIL HEATING FOUND: {heating_text}")
                                    else:
                                        print(f"  ℹ Heating type: {heating_text} (No oil)")
                                    break
                    
                    if not heating_found:
                        print("  ⚠ Could not extract heating information after all methods")
                        # Print first 500 chars of page text for debugging
                        print(f"  ℹ Page text sample: {page_text[:500]}")
                        
                else:
                    print("  ⚠ Interior section not found on page")
                    
            except Exception as e:
                print(f"  ⚠ Error accessing Interior section: {e}")
                import traceback
                traceback.print_exc()
            
            # Get listing agent info
            try:
                listing_agent = self.driver.find_element(By.XPATH, '//span[contains(text(), "Listing by")]/span').text
                property_data['listing_agent'] = listing_agent
            except:
                property_data['listing_agent'] = 'N/A'
            
            try:
                broker = self.driver.find_element(By.CSS_SELECTOR, 'span.agent-basic-details--broker').text
                property_data['broker'] = broker.replace('•', '').strip()
            except:
                property_data['broker'] = 'N/A'
            
            print(f"  ✓ Extracted: {property_data.get('street_address', 'Unknown')} - Oil: {property_data['has_oil_heating']}")
            
        except Exception as e:
            print(f"  ✗ Error extracting property details: {e}")
            property_data['error'] = str(e)
        
        finally:
            # Close tab and switch back
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])
            time.sleep(1)
        
        return property_data
    
    def scrape_current_page(self):
        """Scrape all properties on current page"""
        properties = []
        
        try:
            # Wait for property cards to load
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.bp-Homecard'))
            )
            
            # Get all property links
            property_links = self.driver.find_elements(By.CSS_SELECTOR, 'a.bp-Homecard__Address')
            property_urls = [link.get_attribute('href') for link in property_links if link.get_attribute('href')]
            
            print(f"\n📋 Found {len(property_urls)} properties on this page\n")
            
            # Process each property
            for i, url in enumerate(property_urls, 1):
                print(f"[{i}/{len(property_urls)}] Processing: {url}")
                property_data = self.extract_property_details(url)
                properties.append(property_data)
                time.sleep(1)  # Be nice to the server
            
        except Exception as e:
            print(f"✗ Error scraping page: {e}")
        
        return properties
    
    def has_next_page(self):
        """Check if next page button exists and is clickable"""
        try:
            next_button = self.driver.find_element(By.CSS_SELECTOR, 'button.PageArrow--next')
            classes = next_button.get_attribute('class')
            
            # Check if hidden
            if 'PageArrow--hidden' in classes:
                return False
            return True
        except:
            return False
    
    def go_to_next_page(self):
        """Navigate to next page"""
        try:
            next_button = self.driver.find_element(By.CSS_SELECTOR, 'button.PageArrow--next')
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
            time.sleep(1)
            next_button.click()
            time.sleep(3)
            
            # Wait for new page to load
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.bp-Homecard'))
            )
            return True
        except Exception as e:
            print(f"✗ Error navigating to next page: {e}")
            return False
    
    def save_to_excel(self, properties):
        """Save or append properties to Excel file"""
        if not properties:
            print("⚠ No properties to save")
            return
        
        df_new = pd.DataFrame(properties)
        
        # Check if file exists
        if os.path.exists(self.excel_file):
            try:
                # Read existing data
                df_existing = pd.read_excel(self.excel_file)
                # Append new data
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                # Remove duplicates based on URL
                df_combined = df_combined.drop_duplicates(subset=['url'], keep='first')
                df_combined.to_excel(self.excel_file, index=False)
                print(f"\n✓ Appended {len(df_new)} properties to existing file")
                print(f"✓ Total properties in file: {len(df_combined)}")
            except Exception as e:
                print(f"✗ Error appending to Excel: {e}")
                # Save as backup
                backup_file = f"redfin_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                df_new.to_excel(backup_file, index=False)
                print(f"✓ Saved to backup file: {backup_file}")
        else:
            # Create new file
            df_new.to_excel(self.excel_file, index=False)
            print(f"\n✓ Created new Excel file: {self.excel_file}")
            print(f"✓ Saved {len(df_new)} properties")
    
    def run(self):
        """Main run method"""
        try:
            # Kill any existing Chrome processes first
            self.kill_chrome_processes()
            
            # Setup browser
            self.setup_driver()
            
            # Let user apply filters
            self.start_and_wait_for_user()
            
            all_properties = []
            page_num = 1
            
            while True:
                print("\n" + "="*60)
                print(f"SCRAPING PAGE {page_num}")
                print("="*60)
                
                # Scrape current page
                properties = self.scrape_current_page()
                all_properties.extend(properties)
                
                # Save after each page
                self.save_to_excel(properties)
                
                # Check for next page
                if not self.has_next_page():
                    print("\n✓ No more pages available")
                    break
                
                # Ask user if they want to continue
                print("\n" + "-"*60)
                continue_scraping = input("Continue to next page? (y/n): ").strip().lower()
                
                if continue_scraping != 'y':
                    print("\n✓ Scraping stopped by user")
                    break
                
                # Go to next page
                print("\n→ Navigating to next page...")
                if not self.go_to_next_page():
                    print("✗ Failed to navigate to next page")
                    break
                
                page_num += 1
            
            # Final summary
            print("\n" + "="*60)
            print("SCRAPING COMPLETED!")
            print("="*60)
            print(f"✓ Total properties scraped: {len(all_properties)}")
            print(f"✓ Data saved to: {self.excel_file}")
            
            # Count oil heating properties
            oil_count = sum(1 for p in all_properties if p.get('has_oil_heating') == 'Yes')
            print(f"🔥 Properties with OIL heating: {oil_count}")
            print("="*60 + "\n")
            
        except KeyboardInterrupt:
            print("\n\n⚠ Scraping interrupted by user (Ctrl+C)")
        except Exception as e:
            print(f"\n✗ Error during scraping: {e}")
        finally:
            if self.driver:
                print("\n→ Closing browser...")
                self.driver.quit()
                print("✓ Browser closed")


def main():
    """Run the scraper"""
    print("\n")
    print("╔" + "="*58 + "╗")
    print("║" + " "*15 + "REDFIN WEB SCRAPER" + " "*25 + "║")
    print("║" + " "*12 + "Interactive User Control" + " "*22 + "║")
    print("╚" + "="*58 + "╝")
    
    excel_file = input("\nEnter Excel filename (default: redfin_properties.xlsx): ").strip()
    if not excel_file:
        excel_file = "redfin_properties.xlsx"
    
    if not excel_file.endswith('.xlsx'):
        excel_file += '.xlsx'
    
    print(f"\n✓ Will save to: {excel_file}")
    
    scraper = RedfinScraperInteractive(excel_file=excel_file)
    scraper.run()


if __name__ == "__main__":
    main()