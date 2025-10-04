# CSV to Excel

This program reads a csv file and creates a file in Excel format

Command format to run:
```
      xlsx_main.py -i <inputfile> -o <outputfile> -d <delimiter> -t <title 1=y, 0=n> -f <format 1=y, 0=n>
```
Example  
```      
      python3 xlsx_main.py -i 'test_format_and_header.txt' -o 'python_test.xlsx' -d ';'  -t 1 -f 1
```

## Input parameters
* inputfile   Name of the input file. Encoding is ISO-8859-1
* outputfile  Name of the output file    
* delimiter   Field delimiter in csv file. 
              Default is comma
* format      Indicates whether or not there is a format line in the input file 0=False, 1= True

              Data type per column. Ex generic;int;float;generic
  
              If there is a format line in the input file, it must be the first line.
  
              The valid types are generic, int and float, and 0.00, 0.000, 0.0000, 0.00000, 0.000000, 0.0000000, 0.00000000
  
              Formats are validated. Any invalid format is printed to standard output, and the 
              program exits.
  
              Default is 0 (false)
* title       Indicates whether or not there is a header line in the input file 0=False, 1= True
              If there is a header line in the input file it must be the line after any format line.
              Default is 1 (false)
## Formats
The format denotes the data type for the column values (excluding the header line)

Valid formats
* generic       Writes text as is
* float         Writes the data as float. The input data is expected to use a full stop as the decimal separator. E.g. 250.50
* int           Writes the data as int.
* 0.00          Writes the data as float with two decimals.
* 0.000         Writes the data as float with three decimals.
* 0.0000        Writes the data as float with four decimals.
* 0.00000       Writes the data as float with five decimals.
* 0.000000      Writes the data as float with six decimals.
* 0.0000000     Writes the data as float with seven decimals.
* 0.00000000    Writes the data as float with eight decimals.

## Validation
### Format 
If a format line is provided, the formats are validated before the rest of the file is treated. Leading and trailing spaces are stripped from the format.
Invalid formats are printed to the standard output. 

### Int and float
Cells denoted as int or float are validated before attempting to write them to the Excel file.

### Error logging


## Example files
      
### Header line, no format line
```
Id;No of items;Price per item
AABZ;100;36.50
AACR;3500;22.00

```
      
### Header line and format line
```
generic;int;float;0.000000
Id;No of items;Price per item;Interest
AABZ;100;36.50;13.42145
AACR;3500;22.00;9.223344
```
      
### Format line, no header line
      
      
### No format line, no header line
      
      
