# Python_for_Excel
This repository has the Python code to copy the contents of .csv files from a folder to a destination .xlsx file matching with some given string in seconds. Using notebook (file.ipynb) for the code execution.

Here is a pseudo code for the same:
Write a python code to read .csv files present in a folder and write to a .xlsx file according to the following:
  Read .csv file present in the folder named 'folder'
  Check all the .csv files present in the folder
	
    If the file has String present in its first 'A' column matching with "Kew_" in the iterating A'th column cell, then:
      Copy all the contents of the row with the String "Kew_" and paste it in the .xlsx file
      Continue this iteration for all the .csv files and print "Successful" when done.
