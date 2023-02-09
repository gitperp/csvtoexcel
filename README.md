# CSV to Excel

This program reads a csv file and creates a file in Excel format

Command format to run:
      xlsx_main.py -i <inputfile> -o <outputfile> -d <delimiter> -t <title 1=y, 0=n> -f <format 1=y, 0=n>
Example 
    OpenVMS
```
      python xlsx_main.py -i "transakt.csv" -o "transakt.xlsx" -d ";" -t 1 -f 1
```      
    Windows
```      
      *python3 xlsx_main.py -i 'python_csv_test-ansi_format_and_header.txt' -o 'python_test.xlsx' -d ';'  -t 1 -f 1*
```

## Input parameters
* inputfile   Name of the input file
* outputfile  Name of the output file    
* delimiter   Field delimiter in csv file. 
              Default is comma
* format      Indicates whether or not there is a format line in the input file 0=False, 1= True
              Data type per column. Ex generic;int;float;generic
              If there is a format line in the input file, it must be the first line
              The valid types are generic, int and float
              Formats are validated. Any invalid format is printed to standard output, and the 
              program exits.
              Default is 0 (false)
* title       Indicates whether or not there is a header line in the input file 0=False, 1= True
              If there is a header line in the input file it must be the line after any format line.
              Default is 0 (false)

## Example files
      
### Header line, no format line
      
      
### Header line and format lne
      
### Format line, no header line
      
      
### No format line, no header line
      
      
