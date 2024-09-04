countLineXLSB
This Python script allows users to analyze .xlsb files to count the occurrences of the term "Pending" in the Status column of two selected sheets and calculates the total count by considering the row differences between the two files.

Features
Process .xlsb Files: The script processes .xlsb files in the current directory.
User Selection: Users can select which .xlsb files and sheets to analyze.
Counts "Pending" Status: Counts the occurrences of the term "Pending" in the Status column for both selected sheets.
Row Difference Calculation: Calculates the difference in the total number of rows between the two selected sheets and includes this difference in the final result.
Requirements
Python 3.x
pandas library
pyxlsb library
You can install the required libraries using the following command:

sh
Copy code
pip install pandas pyxlsb
Usage
Clone the repository:

sh
Copy code
git clone https://github.com/abamakbar07/countLineXLSB.git
cd countLineXLSB
Place your .xlsb files in the cloned repository's directory.

Run the script:

sh
Copy code
python script.py
Follow the prompts to select the files and sheets you wish to analyze.

View the results:

The script will output the number of rows with "Pending" status in each file.
It will also calculate the difference in row count between the two selected sheets and display the final result.
Example
After running the script and selecting the files and sheets, the output might look like this:

sql
Copy code
Available .xlsb files:
1. file1.xlsb
2. file2.xlsb
Enter the number corresponding to the first file: 1
Enter the number corresponding to the second file: 2

Available sheets in 'file1.xlsb':
1. Sheet1
2. Sheet2
Enter the number corresponding to the sheet you want to analyze: 1

Available sheets in 'file2.xlsb':
1. SheetA
2. SheetB
Enter the number corresponding to the sheet you want to analyze: 2

Pending in first file: 15
Pending in second file: 20
Difference in row count: 10
Final result (Pending in both files + row difference): 45
License
This project is open-source and available under the MIT License.