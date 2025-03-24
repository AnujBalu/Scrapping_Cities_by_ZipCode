#  pip install selenium pandas openpyxl


# Import required libraries
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Load the Excel file containing ZIP codes
input_file = "zip_codes.xlsx"  # Update with your file name
output_file = "usps_city_names.xlsx"

# Read the ZIP codes from the Excel file
df = pd.read_excel(input_file)
zip_codes = df['ZIP_CODE'].tolist()

# Set up the Selenium WebDriver
driver = webdriver.Chrome()  # Make sure chromedriver is in PATH

# Open the USPS Zip Code Lookup page
url = "https://tools.usps.com/zip-code-lookup.htm?citybyzipcode"
driver.get(url)
time.sleep(3)

# Create a list to store the extracted data
data_list = []

# Loop through each ZIP code
for zip_code in zip_codes:
    try:
        # Locate the input field and enter the ZIP code
        input_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "tZip"))
        )
        input_box.clear()
        input_box.send_keys(str(zip_code))
        time.sleep(1)

        # Click the correct "Find" button after entering ZIP code
        find_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "cities-by-zip-code"))
        )
        find_button.click()
        time.sleep(3)  # Wait for results to load

       # Extract "Recommended City Name"
        try:
            recommended_city_elements = driver.find_elements(By.XPATH, "//div[@class='row col-md-5 col-sm-12 col-xs-12 recommended-cities']/p[@class='row-detail-wrapper']")
            
            # Collect all recommended city names and join them as a string
            recommended_city = ", ".join([city.text.strip() for city in recommended_city_elements if city.text.strip()])
            if not recommended_city:
                recommended_city = "N/A"
        except:
            recommended_city = "N/A"

        # Extract "Other City Names"
        try:
            other_city_elements = driver.find_elements(By.XPATH, "//div[@class='row col-md-7 col-sm-12 col-xs-12 other-city-names']/p[@class='row-detail-wrapper']")
            
            # Collect all city names and join them as a string
            other_cities = ", ".join([city.text.strip() for city in other_city_elements if city.text.strip()])
            if not other_cities:
                other_cities = "N/A"
        except:
            other_cities = "N/A"

        # Add the data to the list
        data_list.append({
            "ZIP Code": zip_code,
            "Recommended City Name": recommended_city,
            "Other City Names": other_cities
        })

        print(f"✅ Scraped ZIP: {zip_code}")

        # Click "Look Up Another ZIP Code™" button
        try:
            lookup_another_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.ID, "look-up-another-zip-code-citybyzipcode")
                )
            )
            lookup_another_button.click()
            time.sleep(2)  # Give some time to go back to the main page
        except:
            print(f"⚠️ Failed to click 'Look Up Another ZIP Code™', reloading page...")
            driver.get(url)  # Reload the page if button fails
            time.sleep(3)

    except Exception as e:
        print(f"❌ Error with ZIP {zip_code}: {e}")
        data_list.append({
            "ZIP Code": zip_code,
            "Recommended City Name": "N/A",
            "Other City Names": "N/A"
        })
        driver.get(url)
        time.sleep(3)
        continue

# Close the browser
driver.quit()

# Save the data to a new Excel file
output_df = pd.DataFrame(data_list)
output_df.to_excel(output_file, index=False)
print(f"✅ Data saved to {output_file}")
