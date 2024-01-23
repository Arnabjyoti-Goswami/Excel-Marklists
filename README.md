# Excel Marklists

Script to generate merged, styled, and sorted excel files, with the average of the numerical columns displayed at the last row. Uses pandas for data manipulation and openpyxl for Excel formatting.

# Usage

Before running the script, ensure that:

- The excel file(s) contain(s) different column headings as the first row (delete the first row if its something like the title of the marklist)

- If there are multiple excel files (for merging all of them into a single output file), each file must have the column headings and data in the same format

- The excel file(s) are present in the same directory as this notebook.

Running the script below will save the required output excel file in the same directory as this notebook.
