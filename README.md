Salesforce Field Reference Scraper
1. Overview
This Python script automates the process of auditing field usage within a Salesforce organization. It reads a list of fields from an Excel spreadsheet, navigates to the "Where is this used?" page for each field, scrapes the list of all references (e.g., Layouts, Flows, Validation Rules), and populates the findings back into the spreadsheet.

This tool is designed to save a significant amount of manual effort required for field cleanup, dependency analysis, and general org maintenance.

2. Features
Automated Browser Navigation: Uses Playwright to log in and navigate through the Salesforce Setup UI.

Dynamic Scraping: Scrapes data from within <iframe> elements on the field dependency page.

Data Filtering: Automatically excludes ReportType references and deduplicates Flow references to capture only the latest version.

Excel Integration: Reads from and writes to a local .xlsx file, inserting new rows with the scraped data.

Hyperlink Generation: Creates clickable hyperlinks to the field and each of its references for easy access from the spreadsheet.

Robust Error Handling: Includes mechanisms for handling special characters in field names and provides a manual override for custom object IDs.

3. Prerequisites
Before you begin, ensure you have the following installed on your system:

Python 3.8+: This script is written in Python. If you don't have it installed, you can download it from python.org. During installation on Windows, make sure to check the box that says "Add Python to PATH".

pip: Python's package installer. This is usually included with modern Python installations.

A command-line terminal:

Windows: PowerShell or Command Prompt.

macOS / Linux: Terminal.

4. Setup Instructions
Follow these steps to set up the project on your local machine.

Step 1: Clone or Download the Repository
Clone this repository to your local machine or download the scrape_sf_references.py script.

Step 2: Navigate to the Project Directory
Open your terminal and navigate to the folder where you saved the script.

cd path/to/your/project/folder

Step 3: Create and Activate a Virtual Environment (Recommended)
This creates an isolated environment for the project's dependencies.

# Create the virtual environment
python3 -m venv venv

# Activate it on macOS or Linux
source venv/bin/activate

# Or, activate it on Windows
.\venv\Scripts\activate

Step 4: Install Required Python Libraries
Run the following command to install all the necessary libraries from the Python Package Index (PyPI).

pip install "playwright==1.44.0" "openpyxl==3.1.2" "beautifulsoup4==4.12.3" "lxml==5.2.2" "tenacity==8.3.0"

Step 5: Install Playwright Browser Binaries
Playwright requires its own browser instances to work. This command will download them. This only needs to be done once.

playwright install

5. Preparing the Excel File
The script requires an .xlsx file with your field data.

The sheet you want to process must contain at least two columns in this exact order:

Column A: Field Label (e.g., Total Amount)

Column B: Field API Name (e.g., Total_Amount__c)

Place this Excel file in the same directory as the scrape_sf_references.py script.

6. Running the Script
You run the script from your terminal using a single command with several arguments to specify what it should do.

Example for a Standard Object (Case)
python scrape_sf_references.py --file "Your_Excel_File.xlsx" --sheet "Cases" --object-api-name "Case" --instance "your-instance.my.salesforce.com" --start-at-label "Status Detail"

Example for a Custom Object (CPQ Quote)
For custom objects, you must also provide the --object-id.

python scrape_sf_references.py --file "Your_Excel_File.xlsx" --sheet "CPQ Quotes" --object-api-name "SBQQ__Quote__c" --object-id "01If2000001ah0m" --instance "your-instance.my.salesforce.com" --start-at-label "*Total Hourly Remote Guarding Rate" --limit 50


Command-Line Arguments
--file: (Required) The name of your Excel file (e.g., "Elite_Test_Copy.xlsx").

--sheet: (Required) The name of the sheet to process (e.g., "CPQ Quotes"). Wrap in quotes if it contains spaces.

--object-api-name: (Required) The API name of the Salesforce object (e.g., "Case" or "SBQQ__Quote__c").

--object-id: (Optional, but required for custom objects) The 15 or 18-character ID of the custom object.

--instance: (Required) Your Salesforce instance domain (e.g., "your-company.my.salesforce.com").

--start-at-label: (Optional) The exact Field Label to begin processing from. If omitted, it starts from the top.

--limit: (Optional) The maximum number of fields to process in a single run. If omitted, it processes all fields. It is highly recommended to use this for batching.

--context: (Optional) The name of the directory to store browser session data. Defaults to sf_ctx.

First Run & Authentication
When you run the script for the first time, a Chromium browser window will open and navigate to Salesforce.

Log in using your credentials as you normally would.

Once you are successfully logged in and see the Salesforce home page, switch back to your terminal and press the Enter key to let the script know it can proceed.
