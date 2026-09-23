Activity: enterprise data connector

Project scenario
You are a Data Integration Specialist at a retail analytics consultancy. Your organization maintains sales data across multiple systems: transaction records in Excel files stored in OneDrive, and regional performance data in SharePoint lists. Leadership needs integrated reports combining these sources to analyze month-over-month trends, regional performance, and identify data quality issues.

Your task is to build a Python-based data integration pipeline that connects to both Excel and SharePoint data sources via Microsoft Graph API, validates and standardizes the data (handling missing values, inconsistent date formats, and type errors), merges datasets on common keys, and generates summary analytics with visualizations. This pipeline will automate manual reporting processes and provide reliable, timely insights to business stakeholders.

Objective
The goal of this activity is to design, implement, and assess an end-to-end testing and monitoring framework for the Azure AI Foundry-powered virtual assistant, ensuring the solution meets all business, security, and compliance requirements in the financial services domain. This entails establishing success metrics such as reliable intent recognition, secure processing and storage of sensitive financial information, seamless escalation protocols to human support staff, and comprehensive traceability for audit purposes. You will design and execute diverse test cases, including typical customer interactions and edge scenarios, verify that secure authentication and role-based access controls are enforced, and confirm that all data privacy protections remain active. Utilizing Azure AI Foundry's integrated monitoring and analytics, you will track virtual assistant performance, document incidents or regulatory nonconformities, and compile a detailed validation checklist to support a phased deployment and ongoing compliance. Successfully completing these activities will ensure a safe, reliable, and compliant conversational solution, fostering improved customer service and organizational trust.

Instructions
 1. Install Python libraries:  

 pip install requests pandas openpyxl msal  

2. Register an app in Microsoft Entra ID (formerly Azure AD) – document your Client ID, Tenant ID, and Client Secret.

Documentation: 
Microsoft Graph authentication walkthrough
  

Screenshot of the Microsoft Entra ID registration portal showing Application (client) ID and permissions panels.
3. Authenticate and obtain an access token using the MSAL library.

Example code provided below.

Common mistake: Incorrect client ID or missing permissions will block access.

4. Use the Graph API to list  Onedrive files with the right endpoint for your auth flow:

Sample code given; check for 200 OK HTTP status.

Delegated: GET https://graph.microsoft.com/v1.0/me/drive/root/children

App-only:  GET https://graph.microsoft.com/v1.0/users/{user-id-or-upn}/drive/root/children

Self-review:

What error message appears if the token is invalid?

5. Use Graph API to access a SharePoint site and enumerate available lists.

Save the list names for later.

Common sticking point: API endpoint may change depending on tenant name – double-check the URL.

Sample Output: List of Excel files and SharePoint lists as Python dictionaries.

Diagram of authentication and data flow: Python script to Microsoft Entra ID to Microsoft Graph API to OneDrive/SharePoint.

1.
Question 1
A Python script fails to retrieve a file from SharePoint even though the file exists. Which is the most likely cause? Select the best answer. 



The Python library pandas are missing.

The script lacks appropriate permissions.

The SharePoint list is too large.

The Excel file format is outdated.

1 point
For the chosen Excel and SharePoint data sources:

Use pandas.read_excel() for Excel and requests/Graph API for SharePoint lists.

Check for missing/null values in key columns.

Use pandas methods like df.isnull().sum() to profile missingness.

Standardize all date values to ISO format (YYYY-MM-DD) using pandas.

Convert number fields to float for consistent analysis.

Catch and report type conversion errors.

Troubleshooting:

Got an error loading a file? CheckGraph permissions and supported formats (.xlsx recommended for Excel APIs)..

Dates not converting? Check locale/format mismatches; standardize to YYYY-MM-DD.

Example table with messy vs. cleaned slash validated data (before-and-after), highlighting null values.

2.
Question 2
While processing a dataset, you notice several rows have missing date values. Which initial step best ensures your analysis is reliable? Select the best answer.



Ignore the missing data.

Delete all rows with missing data.

Validate and decide according to your analysis needs.

Replace missing values with zero.

1 point
Before merging:

1. Verify key columns exist in both datasets:

12
print("Excel columns:", excel_data.columns.tolist())
   print("SharePoint columns:", sharepoint_df.columns.tolist())
2. Check if key column names match:

12
 # Rename if needed
   sharepoint_df = sharepoint_df.rename(columns={'Customer_ID': 'CustomerID'})
3. Ensure key data types match:

1234
print("Excel key type:", excel_data['CustomerID'].dtype)
   print("SharePoint key type:", sharepoint_df['CustomerID'].dtype)
   # Convert if needed
   sharepoint_df['CustomerID'] = sharepoint_df['CustomerID'].astype(int)
4. Merge with explicit parameters:

12345678910111213
merged_data = pd.merge(
       excel_data,
       sharepoint_df,
       on='CustomerID',
       how='inner',  # Only keep matching records
       validate='1:1',  # Ensure no duplicates
       indicator=True  # Show merge source
   )
   
   print(f"\n✓ Merged {len(merged_data)} records")

  5. Verify merge results:

123
print("\nMerged data preview:")
   print(merged_data.head())
   print("\nColumns after merge:", merged_data.columns.tolist())
 Then proceed to:

1. Identify a common key column present in both datasets (e.g., “EmployeeID”).

2. Use pandas.merge() to combine Excel and SharePoint data.

Flowchart: Merge process of Excel and SharePoint datasets with key alignment visualized.
3. Calculate summary metrics:

Example: df['Sales'].mean() for average sales, or df.groupby('Month')['Sales'].sum() for month-over-month change.

4. Write results to a new Excel file, or display in the console.

 Sample bar graph with monthly sales totals

3.
Question 3
You want to combine data from an Excel file and a SharePoint list. What is essential to do before merging these datasets? Select the best answer.  



Sort the Excel file alphabetically.

Convert all data to CSV format.

Identify and align common keys (columns) for merging.

Change all text to uppercase.

1 point
As your scripts grow, small performance and reliability improvements make a big difference. In this step, you’ll add caching, error handling, and logging, and briefly document bottlenecks.  

1.  Add basic logging.  

12345
import logging
logging.basicConfig(
    level=logging.INFO,          # Use DEBUG for more detail
    format="%(asctime)s - %(levelname)s - %(message)s"
)
 Use logging instead of print() for important events:  

123
logging.info("Requesting OneDrive files...")
logging.warning("SharePoint list returned no rows.")
logging.error("Failed to merge datasets: %s", e)
2.  Wrap Graph calls with error handling.

Create a helper function to call Graph API with robust error handling:

123456789101112131415161718
import requests

def call_graph_api(url, headers, params=None):
    try:
        response = requests.get(url, headers=headers, params=params, timeout=30)
        logging.info("Graph call to %s returned status %s", url, response.status_code)

        if response.status_code == 200:
            return response.json()
        else:

Common sticking point: If you ignore non-200 status codes, you may accidentally continue with empty or partial data.

3. Add a simple in-memory cache for repeated calls.

If you call the same endpoint multiple times with the same parameters, the cache results are:

123456789
from functools import lru_cache

@lru_cache(maxsize=32)
def get_one_drive_children_cached(user_mode="me"):
    if user_mode == "me":
        url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
    else:
        url = f"https://graph.microsoft.com/v1.0/users/{user_mode}/drive/root/children"
    return call_graph_api(url, headers)
For the activity, compare runtime with and without caching if you make repeated calls.

Self-review:
Which endpoints in your script are called more than once with the same parameters?

Are those calls good candidates for caching?

4. Time key operations and document bottlenecks.

Use simple timing to identify slow sections:

123456
import time

start = time.perf_counter()
merged_data = pd.merge(excel_data, sharepoint_df, on="CustomerID", how="inner")
end = time.perf_counter()
logging.info("Merge completed in %.3f seconds", end - start)
Check:
    •   Which operations are slowest (API calls, merges, group-bys, writing Excel)?

    •   Are you making unnecessary repeated calls (e.g., same Graph request in a loop)?

Document at least two bottlenecks in comments or a short note, for example:

Troubleshooting: Do you see HTTP 429 or throttling errors? Add basic backoff (e.g., time.sleep() and retry) and reduce duplicate calls.

4.
Question 4
During testing, your script repeatedly calls the same Graph API endpoint with identical parameters, and most of the runtime is spent waiting for these responses. What is the best first step to improve performance? Select the best answer.  



Increase the logging level to DEBUG.

Cache the API response and reuse it for identical requests.

Remove error handling to make the script run faster.

Rewrite the script in another programming language.