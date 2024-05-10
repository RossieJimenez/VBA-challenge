# Stock Analysis with VBA
### Overview
This VBA script is designed to analyze stock data for multiple quarters within an Excel workbook. It calculates the quarterly change, percentage change, and total stock volume for each stock ticker. Additionally, it identifies the stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume.

### Instructions
1. Quarterly Change and Percentage Change:
* The script calculates the quarterly change and percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
* It outputs the ticker symbol, quarterly change, and percentage change in separate columns.
  
2. Total Stock Volume:
*The script calculates the total stock volume of each stock for the quarter.
*It outputs the total stock volume in a separate column.

3. Conditional Formatting:
* Positive changes are highlighted in green, while negative changes are highlighted in red using conditional formatting.
  
4. Greatest % Increase, Greatest % Decrease, and Greatest Total Volume:
* The script identifies the stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume for the quarter.
*It outputs the ticker symbols and corresponding values in a separate section.

### Usage
1. Open the Workbook:
* Open the Excel workbook containing the stock data.
* 
2. Enable Macros:
* If prompted, enable macros to allow the VBA script to run.
  
3. Run the Script:
* Navigate to the Developer tab and click on "Macros."
* Select the "stocks_hw" macro and click "Run" to execute the script.
  
4. View Results:
* The script will analyze the stock data and output the results in the specified columns.
* Check the output for the quarterly change, percentage change, total stock volume, and stocks with the greatest changes.

### Additional Notes
* Moderate Solution:
  - The script has been enhanced to identify stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume.
* Hard Solution:
- The script is designed to run on every worksheet (quarter) within the workbook simultaneously, allowing for comprehensive analysis across multiple quarters.
* Conditional Formatting:
** Ensure that conditional formatting is applied to highlight positive and negative changes for better visualization.
* Customization:
** Modify the script as needed to accommodate specific requirements or additional functionalities.
