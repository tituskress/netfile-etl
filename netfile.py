import xlrd
import os
import glob
import datetime

# Step 1
current_year = datetime.datetime.now().year
years_with_data = range(2011, current_year + 1)
remote_path = "https://ssl.netfile.com/pub2/excel/COAKBrowsable/"

for year in years_with_data:
  print "Downloading " + str(year) + " data..."
  # Construct filename
  filename_for_year = "efile_newest_COAK_" + str(year) + ".zip"  
  # Download file from http
  os.system("curl -f -L -O " + remote_path + filename_for_year)
  # Unzip archives, overwrite existing files quietly
  os.system("unzip -oq " + filename_for_year)
  # Delete downloaded zip file
  os.system("del /q " + filename_for_year)
# Step 2
# Populate variable with list of files
xlsx_files_in_current_dir = glob.glob("*.xlsx")
# Loop through file list
for excel_filename in xlsx_files_in_current_dir:
  print "Opening " + excel_filename + "..."
  workbook = xlrd.open_workbook("./" + excel_filename)
  filename_without_extension = excel_filename.split(".")[0]
  # Populate variable with list of worksheets
  sheet_names = workbook.sheet_names()
  # Loop through worksheet list
  for sheet_name in sheet_names:
    # Construct output filename
    csv_filename = filename_without_extension + "_" +sheet_name + ".csv"
    print "Writing " + csv_filename
    # Export worksheet to flat file csv
    os.system("in2csv " + excel_filename + " --sheet " + "\"" + sheet_name + "\"" + " > " + csv_filename)
# Step 3
# Populate variable with file list
csv_files_for_2011 = glob.glob("*2011*.csv")
# Loop through list
for filename in csv_files_for_2011:
  # Build awk expression
  filename_wildcard_expression = filename.replace("2011","*")
  # Set output filename
  output_filename = filename.replace("_2011","")
  # build awk command
  av = 'awk "FNR==1 && NR!=1"{next;}{print} '
  cv = filename_wildcard_expression + " > " + output_filename
  # Output command
  print av + cv
  # Run command
  os.system(av + cv)
