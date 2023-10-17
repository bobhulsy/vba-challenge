Stock Analysis using VBA
This VBA script provides an analysis of stock data from multiple worksheets in an Excel workbook. The script extracts key metrics for each stock ticker and further identifies stocks with the "Greatest % Increase", "Greatest % Decrease", and "Greatest Total Volume".

Features:
Extract Key Metrics: For each stock, the script will calculate:

Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
Total stock volume for the year.
Highlight Key Performers: Based on the extracted metrics, the script identifies and highlights:

Stock with the "Greatest % Increase".
Stock with the "Greatest % Decrease".
Stock with the "Greatest Total Volume".
Multi-Sheet Analysis: The script is designed to analyze stock data across multiple sheets in a workbook, making it easy to compare metrics year over year.

Usage:
Open your Excel workbook that contains stock data in multiple sheets. Ensure each sheet's data follows the structure: <ticker>, <date>, <open>, <high>, <low>, <close>, and <vol>.

Press ALT + F11 to open the VBA editor in Excel.

Insert a new module (Right-click on any existing module or the workbook name on the left > Insert > Module).

Copy the provided VBA script into this new module.

Close the VBA editor.

Press ALT + F8, select MainAnalysis from the list, and click "Run".

After execution, each worksheet will have an additional table that provides a summary of the metrics for each stock and another table highlighting the stocks with the greatest percentage increase, greatest percentage decrease, and the greatest total volume.

Modules and Routines:
ExtractStockMetrics:

Extracts the primary metrics for each stock across all worksheets.
AnalyzeMetrics:

Based on the metrics extracted, this routine identifies stocks with the "Greatest % Increase", "Greatest % Decrease", and "Greatest Total Volume" across all worksheets.
MainAnalysis:

A controller routine that runs ExtractStockMetrics followed by AnalyzeMetrics for ease of use.
