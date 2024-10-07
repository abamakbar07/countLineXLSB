# countLineXLSB

This Python script allows users to analyze `.xlsb` files to count the occurrences of the term "Pending" in the `Status` column of two selected sheets and calculates the total count by considering the row differences between the two files.

## Features

- **Process `.xlsb` Files**: The script processes `.xlsb` files in the current directory.
- **Counts "Pending" Status**: Counts the occurrences of the term "Pending" in the `Status` column for both selected sheets.
- **Row Difference Calculation**: Calculates the difference in the total number of rows between the two selected sheets and includes this difference in the final result.

## Requirements

- Python 3.x
- `pandas` library
- `pyxlsb` library

You can install the required libraries using the following command:

```sh
pip install pandas pyxlsb

```

## Usage

1. Clone the repository:

```sh
git clone https://github.com/abamakbar07/countLineXLSB.git
cd countLineXLSB
```

2. Place your `.xlsb` files in the cloned repository's directory.

3. Run the script:

```sh
python script.py
```

4. Follow the prompts to select the files and sheets you wish to analyze.

5. View the results:

- The script will output the number of rows with "Pending" status in each file.
- It will also calculate the difference in row count between the two selected sheets and display the final result.

