# JPG_to_EXCEL
Transform *.JPG file to *.XLSX file by color-formatting the cells in the Excel sheet.

This is a python shell script to transform a *.JPG file into a *.XLSX file. I saw this "artist" on youtube who (supposedly) paints in Excel by color-formatting each cell. I thought this could be automated and wrote this script (maybe he did the same?).
Just enter the path of your *.JPG and the path were the *.XLSX should be saved. Theoretically, the Excel "picture" should be limited to 1,048,576 rows by 16,384 columns (i.e. pixels), but Excel sometimes gives me errors when opening the file even if the translated *.JPG is much smaller than those limits. I am not sure why this is yet but I am not going to investigate further as this was just a proof-of-concept script.
I attached a sample *.JPG that should work and the *.XLSX output that should be generated.
