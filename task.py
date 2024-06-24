from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# Set up ChromeDriver options
chrome_options = webdriver.ChromeOptions()
# chrome_options.add_argument("--headless")  # Uncomment this line for headless mode

# Specify the path to the ChromeDriver executable
driver_path = r'C:\Users\Pappu kumar\Downloads\chromedriver-win64\chromedriver-win64\chromedriver.exe'  # Update this path to your chromedriver.exe location

# Initialize ChromeDriver
service = ChromeService(executable_path=driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

# Add logging
print("Starting the browser and navigating to the URL...")

# Navigate to the provided URL
url = 'https://www.mcxindia.com/market-data/spot-market-price'
driver.get(url)

# Wait for the page to load and the required elements to be available
wait = WebDriverWait(driver, 20)  # Increased the wait time to 20 seconds

# Add logging
print("Page loaded. Looking for the button...")

# Click the button to change it to another state (e.g., from "Show" to "Hide")
try:
    button = wait.until(EC.element_to_be_clickable((By.ID, 'myBtn')))
    button.click()
    print("Button clicked.")
except Exception as e:
    print(f"Error clicking button: {e}")
    print(driver.page_source)  # Print page source for debugging

# Fill in the required details
try:
    commodity_dropdown = wait.until(EC.presence_of_element_located((By.ID, 'ctl00_ContentPlaceHolder1_ddlCommodity')))
    Select(commodity_dropdown).select_by_visible_text('Gold')
    print("Commodity selected.")
    
    location_dropdown = wait.until(EC.presence_of_element_located((By.ID, 'ctl00_ContentPlaceHolder1_ddlLocation')))
    Select(location_dropdown).select_by_visible_text('Mumbai')
    print("Location selected.")
except Exception as e:
    print(f"Error selecting dropdowns: {e}")
    print(driver.page_source)  # Print page source for debugging

# Click on the "Show" button to get the table
try:
    show_button = wait.until(EC.element_to_be_clickable((By.ID, 'ctl00_ContentPlaceHolder1_btnShow')))
    show_button.click()
    print("Show button clicked.")
except Exception as e:
    print(f"Error clicking Show button: {e}")
    print(driver.page_source)  # Print page source for debugging

# Wait for the table to load
try:
    table = wait.until(EC.presence_of_element_located((By.ID, 'ctl00_ContentPlaceHolder1_gridRecords')))
    print("Table found.")
except Exception as e:
    print(f"Error finding table: {e}")
    print(driver.page_source)  # Print page source for debugging

# Extract the table data
try:
    rows = table.find_elements(By.TAG_NAME, 'tr')
    data = []
    for row in rows:
        cols = row.find_elements(By.TAG_NAME, 'td')
        cols = [col.text for col in cols]
        data.append(cols)

    # Convert to Pandas DataFrame
    df = pd.DataFrame(data[1:], columns=data[0])  # Assuming the first row is the header
    print("Data extracted.")
except Exception as e:
    print(f"Error extracting data: {e}")

# Data analysis tasks
try:
    total_rows = len(df)
    df["Spot Price (Rs.)"] = df["Spot Price (Rs.)"].str.replace(',', '').astype(float)
    highest_spot_price_row = df.loc[df["Spot Price (Rs.)"].idxmax()]
    highest_spot_price_date = highest_spot_price_row["Date"]

    # Save the data to an Excel file
    output_file = 'output.xlsx'
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Raw Data', index=False)
    print("Data analysis complete and saved to Excel.")
except Exception as e:
    print(f"Error in data analysis: {e}")

# Close the WebDriver
driver.quit()

# Print the results
print(f"Total number of rows: {total_rows}")
print(f"Date with highest Spot Price (Rs.): {highest_spot_price_date}")
