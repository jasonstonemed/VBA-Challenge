# VBA-Challenge

This coding challenge offered three levels of difficulty: easy, medium, and hard. This code represents the highest level of difficulty.

# Stock Market Data Analysis - VBA Macro

This VBA macro is designed to analyze stock market data in an Excel worksheet. It calculates and summarizes information for each stock ticker symbol, including yearly change, percentage change, and total volume. Additionally, it identifies the stocks with the greatest percentage increase, percentage decrease, and total trading volume.

## How to Use

1. **Open the Excel Workbook**: Ensure that you have your stock market data in an Excel workbook, and that you have a VBA module to run the macro.

2. **Access the VBA Editor**: Press `Alt + F11` to open the Visual Basic for Applications (VBA) editor.

3. **Insert or Copy-Paste the Code**: Insert this code into a new or existing VBA module within your workbook.

4. **Run the Macro**: Execute the `allWorksheets` macro to analyze data on all worksheets within the workbook. You can also call the `YearlySummary` macro with a specific worksheet as an argument if you want to analyze a single worksheet.

5. **Review the Results**: The macro will populate the following columns in your worksheet:
   - Column K: Ticker symbol
   - Column L: Yearly change (with conditional formatting for positive and negative changes)
   - Column M: Percentage change
   - Column N: Total trading volume

   Additionally, the greatest percentage increase, greatest percentage decrease, and greatest total volume will be displayed in columns Q and R.

## Important Notes

- Ensure your data is organized with columns containing the following information:
   - Column A: Ticker symbols
   - Column C: Opening prices
   - Column F: Closing prices
   - Column G: Trading volumes

- The macro assumes that your data is sorted by ticker symbol and date in ascending order.

- Make sure to save your workbook with the `.xlsm` extension to preserve the VBA code.

## Customization

You can customize the VBA code to fit your specific needs. For example, you can adjust the column references, formatting rules, or add error handling for different scenarios.

Remember to always make a backup of your data before running any macros, especially if you are not familiar with VBA, to prevent accidental data loss.

Credit to ChatGPT for assisting in creating this readme file
