# countLineXLSB

This Python script allows users to analyze `.xlsb` files to count the occurrences of the term "Pending" in the `Status` column of two selected sheets and calculates the total count by considering the row differences between the two files.

## Features

- **Process `.xlsb` Files**: The script processes `.xlsb` files in the current directory.
- **User Selection**: Users can select which `.xlsb` files and sheets to analyze.
- **Counts "Pending" Status**: Counts the occurrences of the term "Pending" in the `Status` column for both selected sheets.
- **Row Difference Calculation**: Calculates the difference in the total number of rows between the two selected sheets and includes this difference in the final result.

## Requirements

- Python 3.x
- `pandas` library
- `pyxlsb` library

You can install the required libraries using the following command:

```sh
pip install pandas pyxlsb
