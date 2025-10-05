"""
Redfin Property Scraper - Interactive User Control
Fixed version with auto-save and oil-only filtering
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
        self.properties_saved_count = 0
        self.start_element = 1
        self.current_page_num = 1
    
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
        
        # Ask about starting page and element
        print("\n" + "-"*60)
        print("STARTING POSITION:")
        print("-"*60)
        
        start_page_input = input("What page should we start from? (default: 1): ").strip()
        self.current_page_num = int(start_page_input) if start_page_input else 1
        
        # Navigate to starting page if needed
        if self.current_page_num > 1:
            print(f"\nNavigating to page {self.current_page_num}...")
            current_url = self.driver.current_url
            if '/page-' in current_url:
                base_url = current_url.split('/page-')[0]
                new_url = f"{base_url}/page-{self.current_page_num}"
            else:
                new_url = f"{current_url}/page-{self.current_page_num}"
            
            self.driver.get(new_url)
            time.sleep(3)
            print(f"✓ On page {self.current_page_num}")
        
        # Wait for page to fully load and count ACTUAL elements
        time.sleep(2)
        total_elements = 0
        try:
            # Wait for property cards to load
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.bp-Homecard'))
            )
            time.sleep(1)  # Extra wait for all elements to render
            
            # Count actual property links
            property_links = self.driver.find_elements(By.CSS_SELECTOR, 'a.bp-Homecard__Address')
            property_urls = [link.get_attribute('href') for link in property_links if link.get_attribute('href')]
            total_elements = len(property_urls)
            print(f"\n📋 Found {total_elements} properties on this page")
        except Exception as e:
            print(f"\n⚠ Could not count properties: {e}")
            total_elements = 40  # Default fallback
            print(f"📋 Using default count: {total_elements} properties")
        
        # Ask which element to start from
        if total_elements > 0:
            start_element_input = input(f"Which property should we start from? (1-{total_elements}, default: 1): ").strip()
        else:
            start_element_input = input(f"Which property should we start from? (default: 1): ").strip()
            
        self.start_element = int(start_element_input) if start_element_input else 1
        
        if self.start_element < 1:
            self.start_element = 1
        elif total_elements > 0 and self.start_element > total_elements:
            print(f"⚠ Element {self.start_element} is out of range, starting from element 1")
            self.start_element = 1
        
        if self.start_element > 1:
            print(f"✓ Will start from property #{self.start_element} (skipping first {self.start_element - 1})")
        else:
            print(f"✓ Will start from property #1")
        
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
        property_data = {
            'url': property_url,
            'listing_status': 'unknown',
            'scrape_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        try:
            # Open in new tab with the URL directly
            self.driver.execute_script(f"window.open('{property_url}', '_blank');")
            time.sleep(2)
            self.driver.switch_to.window(self.driver.window_handles[-1])
            
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
            
            # Detect listing status dynamically from the banner
            try:
                status_banner = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div.ListingStatusBannerSection'))
                )
                banner_text = status_banner.text.upper()
                
                if 'SOLD' in banner_text:
                    property_data['listing_status'] = 'sold'
                    # Extract sold date
                    if 'ON' in banner_text:
                        sold_date = banner_text.split('ON')[1].strip()
                        property_data['sold_date'] = sold_date
                    else:
                        property_data['sold_date'] = banner_text.replace('SOLD', '').strip()
                    print(f"  ℹ Status: SOLD on {property_data['sold_date']}")
                elif 'FOR SALE' in banner_text:
                    property_data['listing_status'] = 'for-sale'
                    property_data['sold_date'] = '-'
                    print(f"  ℹ Status: FOR SALE")
                else:
                    property_data['listing_status'] = 'unknown'
                    property_data['sold_date'] = '-'
                    print(f"  ⚠ Status: Unknown ({banner_text})")
                    
            except Exception as e:
                print(f"  ⚠ Could not detect listing status: {e}")
                property_data['listing_status'] = 'unknown'
                property_data['sold_date'] = '-'
            
            # Get address - try multiple selectors for both sold and for-sale
            try:
                # Method 1: Try the standard full address (works for SOLD)
                try:
                    address_elem = self.driver.find_element(By.CSS_SELECTOR, 'h1.full-address')
                    full_address_text = address_elem.text.strip()
                    
                    # Parse the address - SOLD format: "25 Schooner Ln, Port Washington, NY 11050"
                    if ',' in full_address_text:
                        parts = [p.strip() for p in full_address_text.split(',')]
                        
                        if len(parts) >= 3:
                            # Format: Street, City, State Zip
                            property_data['street_address'] = parts[0]
                            property_data['city'] = parts[1]
                            
                            # Parse "NY 11050"
                            state_zip = parts[2].split()
                            if len(state_zip) >= 2:
                                property_data['state'] = state_zip[0]
                                property_data['zip_code'] = state_zip[1]
                            else:
                                property_data['state'] = parts[2]
                                property_data['zip_code'] = '-'
                                
                        elif len(parts) == 2:
                            # Format: Street, City State Zip
                            property_data['street_address'] = parts[0]
                            
                            # Parse "Port Washington NY 11050"
                            remainder = parts[1].split()
                            if len(remainder) >= 3:
                                property_data['zip_code'] = remainder[-1]
                                property_data['state'] = remainder[-2]
                                property_data['city'] = ' '.join(remainder[:-2])
                            elif len(remainder) == 2:
                                property_data['state'] = remainder[0]
                                property_data['zip_code'] = remainder[1]
                                property_data['city'] = '-'
                            else:
                                property_data['city'] = parts[1]
                                property_data['state'] = '-'
                                property_data['zip_code'] = '-'
                        else:
                            property_data['street_address'] = parts[0]
                            property_data['city'] = '-'
                            property_data['state'] = '-'
                            property_data['zip_code'] = '-'
                    else:
                        # Single line address
                        property_data['street_address'] = full_address_text
                        property_data['city'] = '-'
                        property_data['state'] = '-'
                        property_data['zip_code'] = '-'
                    
                    property_data['full_address'] = full_address_text
                    print(f"  ✓ Address: {property_data['full_address']}")
                    
                except:
                    # Method 2: For FOR SALE properties with different structure
                    try:
                        # Try street-address class specifically
                        street_elem = self.driver.find_element(By.CSS_SELECTOR, 'h1.street-address')
                        full_address_text = street_elem.text.strip()
                        
                        # Parse: "166 N Oak St, Massapequa, NY 11758"
                        if ',' in full_address_text:
                            parts = [p.strip() for p in full_address_text.split(',')]
                            
                            if len(parts) >= 3:
                                # Street, City, State ZIP
                                property_data['street_address'] = parts[0]
                                property_data['city'] = parts[1]
                                
                                # Parse "NY 11758"
                                state_zip = parts[2].split()
                                if len(state_zip) >= 2:
                                    property_data['state'] = state_zip[0]
                                    property_data['zip_code'] = state_zip[1]
                                else:
                                    property_data['state'] = parts[2]
                                    property_data['zip_code'] = '-'
                            elif len(parts) == 2:
                                property_data['street_address'] = parts[0]
                                # Try to parse city, state, zip from second part
                                remainder = parts[1].split()
                                if len(remainder) >= 2:
                                    property_data['zip_code'] = remainder[-1]
                                    property_data['state'] = remainder[-2]
                                    property_data['city'] = ' '.join(remainder[:-2])
                                else:
                                    property_data['city'] = parts[1]
                                    property_data['state'] = '-'
                                    property_data['zip_code'] = '-'
                            else:
                                property_data['street_address'] = parts[0]
                                property_data['city'] = '-'
                                property_data['state'] = '-'
                                property_data['zip_code'] = '-'
                        else:
                            property_data['street_address'] = full_address_text
                            property_data['city'] = '-'
                            property_data['state'] = '-'
                            property_data['zip_code'] = '-'
                        
                        property_data['full_address'] = full_address_text
                        print(f"  ✓ Address: {property_data['full_address']}")
                        
                    except Exception as e2:
                        print(f"  ⚠ Both address methods failed: {e2}")
                        property_data['street_address'] = '-'
                        property_data['city'] = '-'
                        property_data['state'] = '-'
                        property_data['zip_code'] = '-'
                        property_data['full_address'] = '-'
                        
            except Exception as e:
                print(f"  ⚠ Error extracting address: {e}")
                property_data['street_address'] = '-'
                property_data['city'] = '-'
                property_data['state'] = '-'
                property_data['zip_code'] = '-'
                property_data['full_address'] = '-'
            
            # Get price
            try:
                price_elem = self.driver.find_element(By.CSS_SELECTOR, 'div.statsValue')
                property_data['price'] = price_elem.text.strip()
            except:
                try:
                    price_elem = self.driver.find_element(By.CSS_SELECTOR, 'div.price')
                    property_data['price'] = price_elem.text.strip()
                except:
                    property_data['price'] = '-'
            
            # Get beds
            try:
                beds_elem = self.driver.find_element(By.CSS_SELECTOR, 'div.beds-section .statsValue')
                property_data['beds'] = beds_elem.text.strip()
            except:
                property_data['beds'] = '-'
            
            # Get baths
            try:
                baths_elem = self.driver.find_element(By.CSS_SELECTOR, 'div.baths-section .statsValue')
                property_data['baths'] = baths_elem.text.strip()
            except:
                property_data['baths'] = '-'
            
            # Get sqft
            try:
                sqft_elem = self.driver.find_element(By.CSS_SELECTOR, 'div.sqft-section .statsValue')
                property_data['sqft'] = sqft_elem.text.strip()
            except:
                property_data['sqft'] = '-'
            
            # Get property type
            try:
                # Look for property type in key details
                prop_type_elem = self.driver.find_element(By.XPATH, '//span[text()="Property Type"]/preceding-sibling::span[@class="valueText"]')
                property_data['property_type'] = prop_type_elem.text.strip()
            except:
                property_data['property_type'] = '-'
            
            # Scroll down to find Interior section
            try:
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight/2);")
                time.sleep(1)
            except:
                pass
            
            # Find and click Interior section
            property_data['has_oil_heating'] = 'No'
            property_data['heating_type'] = '-'
            property_data['cooling_type'] = '-'
            
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
                property_data['listing_agent'] = '-'
            
            try:
                broker = self.driver.find_element(By.CSS_SELECTOR, 'span.agent-basic-details--broker').text
                property_data['broker'] = broker.replace('•', '').strip()
            except:
                property_data['broker'] = '-'
            
            print(f"  ✓ Extracted: {property_data.get('street_address', 'Unknown')} - Oil: {property_data['has_oil_heating']}")
            
        except Exception as e:
            print(f"  ✗ Error extracting property details: {e}")
            property_data['error'] = str(e)
        
        except Exception as outer_e:
            print(f"  ✗ Critical error (browser issue): {outer_e}")
            property_data['error'] = f"Browser error: {str(outer_e)}"
        
        finally:
            # Safely close tab and switch back
            try:
                # Check if we have multiple windows before closing
                if len(self.driver.window_handles) > 1:
                    self.driver.close()
                    self.driver.switch_to.window(self.driver.window_handles[0])
                    time.sleep(1)
            except Exception as close_error:
                print(f"  ⚠ Error closing tab: {close_error}")
                # Try to recover by switching to first window
                try:
                    if len(self.driver.window_handles) > 0:
                        self.driver.switch_to.window(self.driver.window_handles[0])
                except:
                    pass
        
        return property_data
    
    def save_property_immediately(self, property_data):
        """Save single property immediately if it has oil heating"""
        # Only save if has oil heating
        if property_data.get('has_oil_heating') != 'Yes':
            print(f"  ⊗ Skipped (No oil heating)")
            return False
        
        try:
            # Create DataFrame from single property
            df_new = pd.DataFrame([property_data])
            
            # Remove unnecessary columns before saving
            columns_to_remove = ['url', 'scrape_date', 'price', 'beds', 'baths', 
                                'sqft', 'has_oil_heating', 'listing_agent', 'broker']
            for col in columns_to_remove:
                if col in df_new.columns:
                    df_new = df_new.drop(col, axis=1)
            
            # Check if file exists
            if os.path.exists(self.excel_file):
                # Read existing data
                df_existing = pd.read_excel(self.excel_file)
                
                # Append new data
                df_combined = pd.concat([df_existing, df_new], ignore_index=True)
                
                # Remove duplicates based on full_address
                df_combined = df_combined.drop_duplicates(subset=['full_address'], keep='first')
                
                # Save back to file
                df_combined.to_excel(self.excel_file, index=False)
                
                self.properties_saved_count += 1
                print(f"  ✓ SAVED to Excel (Total oil properties: {self.properties_saved_count})")
            else:
                # Create new file
                df_new.to_excel(self.excel_file, index=False)
                self.properties_saved_count += 1
                print(f"  ✓ Created Excel file and saved property (Total: {self.properties_saved_count})")
            
            return True
            
        except Exception as e:
            print(f"  ✗ Error saving property: {e}")
            return False
    
    def scrape_current_page(self):
        """Scrape all properties on current page"""
        oil_properties_on_page = 0
        
        try:
            # Wait for property cards to load
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.bp-Homecard'))
            )
            time.sleep(1)  # Extra wait for all to render
            
            # Get all property links
            property_links = self.driver.find_elements(By.CSS_SELECTOR, 'a.bp-Homecard__Address')
            property_urls = [link.get_attribute('href') for link in property_links if link.get_attribute('href')]
            
            total_on_page = len(property_urls)
            print(f"\n📋 Found {total_on_page} properties on this page")
            
            # Apply starting element filter (only for first time)
            if self.start_element > 1:
                print(f"⚡ Starting from property #{self.start_element}")
                property_urls = property_urls[self.start_element - 1:]  # Python uses 0-based index
                properties_to_scrape = len(property_urls)
                print(f"📋 Will scrape {properties_to_scrape} properties (skipped first {self.start_element - 1})")
                # Reset to 1 for subsequent pages
                self.start_element = 1
            else:
                properties_to_scrape = total_on_page
            
            print()
            
            # Process each property
            for i, url in enumerate(property_urls, 1):
                print(f"[{i}/{properties_to_scrape}] Processing: {url}")
                
                try:
                    property_data = self.extract_property_details(url)
                    
                    # Save immediately if has oil heating
                    if self.save_property_immediately(property_data):
                        oil_properties_on_page += 1
                    
                    time.sleep(1)  # Be nice to the server
                    
                except Exception as e:
                    print(f"  ✗ Error processing property: {e}")
                    # Continue to next property instead of stopping
                    continue
            
            print(f"\n✓ Oil properties found on this page: {oil_properties_on_page}")
            
        except Exception as e:
            print(f"✗ Error scraping page: {e}")
        
        return oil_properties_on_page
    
    def has_next_page(self):
        """Check if next page button exists and is clickable"""
        try:
            # Look for the next button with the new structure
            next_button = self.driver.find_element(By.CSS_SELECTOR, 'button.PageArrow__direction--next')
            
            # Check if it has the hidden class
            button_classes = next_button.get_attribute('class')
            if 'PageArrow--hidden' in button_classes:
                print("  ℹ Next button is hidden (last page)")
                return False
            
            # Check if button is disabled
            if not next_button.is_enabled():
                print("  ℹ Next button is disabled (last page)")
                return False
                
            return True
            
        except NoSuchElementException:
            print("  ℹ Next button not found (last page)")
            return False
        except Exception as e:
            print(f"  ⚠ Error checking for next page: {e}")
            return False
    
    def go_to_next_page(self):
        """Navigate to next page"""
        try:
            # Find the next button
            next_button = self.driver.find_element(By.CSS_SELECTOR, 'button.PageArrow__direction--next')
            
            # Scroll to it
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
            time.sleep(1)
            
            # Click it
            try:
                next_button.click()
                print("  ✓ Clicked next button")
            except:
                # Try JavaScript click if normal click fails
                self.driver.execute_script("arguments[0].click();", next_button)
                print("  ✓ Clicked next button (JavaScript)")
            
            time.sleep(3)
            
            # Wait for new page to load
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.bp-Homecard'))
            )
            
            print("  ✓ New page loaded")
            return True
            
        except Exception as e:
            print(f"  ✗ Error navigating to next page: {e}")
            return False
    
    def run(self):
        """Main run method"""
        try:
            # Kill any existing Chrome processes first
            self.kill_chrome_processes()
            
            # Setup browser
            self.setup_driver()
            
            # Let user apply filters
            self.start_and_wait_for_user()
            
            # Continue scraping until no more pages or user stops
            while True:
                print("\n" + "="*60)
                print(f"SCRAPING PAGE {self.current_page_num}")
                print("="*60)
                
                # Scrape current page (saves automatically)
                oil_count = self.scrape_current_page()
                
                print(f"\n📊 Page {self.current_page_num} Summary:")
                print(f"   • Oil properties on this page: {oil_count}")
                print(f"   • Total oil properties saved: {self.properties_saved_count}")
                
                # Check for next page
                print("\n→ Checking for next page...")
                if not self.has_next_page():
                    print("\n" + "="*60)
                    print("✓ NO MORE PAGES - Reached the end")
                    print("="*60)
                    break
                
                # Ask user if they want to continue
                print("\n" + "-"*60)
                print(f"📄 More pages available...")
                continue_scraping = input("Continue to next page? (y/n, default: y): ").strip().lower()
                
                if continue_scraping == 'n':
                    print("\n✓ Scraping stopped by user")
                    break
                
                # Go to next page
                print("\n→ Navigating to next page...")
                if not self.go_to_next_page():
                    print("\n✗ Failed to navigate to next page - stopping")
                    break
                
                self.current_page_num += 1
                time.sleep(2)  # Extra wait between pages
            
            # Final summary
            print("\n" + "="*60)
            print("SCRAPING COMPLETED!")
            print("="*60)
            print(f"📄 Total pages scraped: {self.current_page_num}")
            print(f"🔥 Total OIL properties saved: {self.properties_saved_count}")
            print(f"✓ Data saved to: {self.excel_file}")
            print("="*60 + "\n")
            
        except KeyboardInterrupt:
            print("\n\n⚠ Scraping interrupted by user (Ctrl+C)")
            print(f"✓ Data saved before interruption: {self.properties_saved_count} oil properties")
            print(f"✓ Check file: {self.excel_file}")
        except Exception as e:
            print(f"\n✗ Error during scraping: {e}")
            print(f"✓ Data saved so far: {self.properties_saved_count} oil properties")
            import traceback
            traceback.print_exc()
        finally:
            if self.driver:
                print("\n→ Closing browser...")
                try:
                    self.driver.quit()
                    print("✓ Browser closed")
                except:
                    print("⚠ Browser may still be open")


def main():
    """Run the scraper"""
    print("\n")
    print("╔" + "="*58 + "╗")
    print("║" + " "*15 + "REDFIN WEB SCRAPER" + " "*25 + "║")
    print("║" + " "*10 + "OIL HEATING PROPERTIES ONLY" + " "*20 + "║")
    print("╚" + "="*58 + "╝")
    
    excel_file = input("\nEnter Excel filename (default: redfin_oil_properties.xlsx): ").strip()
    if not excel_file:
        excel_file = "redfin_oil_properties.xlsx"
    
    if not excel_file.endswith('.xlsx'):
        excel_file += '.xlsx'
    
    print(f"\n✓ Will save to: {excel_file}")
    print("✓ Only properties with OIL heating will be saved")
    print("✓ URL column will NOT be included")
    print("✓ Data is saved after EACH property (safe from interruptions)")
    
    if os.path.exists(excel_file):
        print(f"\n⚠ File already exists: {excel_file}")
        print("✓ New data will be APPENDED to existing file")
    
    scraper = RedfinScraperInteractive(excel_file=excel_file)
    scraper.run()


if __name__ == "__main__":
    main()