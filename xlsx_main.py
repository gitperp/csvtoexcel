import sys, getopt
import xlsxwriter
import csv

# This program reads a csv file and creates a file in Excel format
# ================================================================
# Command format to run:
#       xlsx_main.py -i <inputfile> -o <outputfile> -d <delimiter> -t <title 1=y, 0=n> -f <format 1=y, 0=n>
# Example 
#     OpenVMS::
#       python xlsx_main.py -i "transakt.csv" -o "transakt.xlsx" -d ";" -t 1 -f 1
#     Windows::
#       python3 xlsx_main.py -i 'python_csv_test-ansi_format_and_header.txt' -o 'python_test.xlsx' -d ';'  -t 1 -f 1


# Input parameters
# ----------------
# - inputfile   Name of the input file
# - outputfile  Name of the output file    
# - delimiter   Field delimiter in csv file. 
#               Default is comma
# - format      Indicates whether or not there is a format line in the input file 0=False, 1= True
#               Data type per column. Ex generic;int;float;generic
#               If there is a format line in the input file, it must be the first line
#               The valid types are generic, int and float
#               Formats are validated. Any invalid format is printed to standard output, and the 
#               program exits.
#               Default is 0 (false)
# - title       Indicates whether or not there is a header line in the input file 0=False, 1= True
#               If there is a header line in the input file it must be the line after any format line.
#               Default is 0 (false)
# 


# VMS specific modules
# --------------------
if sys.platform == 'OpenVMS':
    import vms.decc
    import vms.sys
    import vms.ile3


input_file_encoding = "ISO8859-1"
#input_file_encoding = "ANSI"

# Check if value is a float
def isfloat(value):
    try:
        a = float(value)
    except (TypeError, ValueError):
        return False
    else:
        return True

def isValidFormat(validFormats, format):
    for frm in validFormats:
        if (frm == format):
            return True
    return False
     
def validateFormats(validFormats, formatRow, formatDict):
    isValid = True
    for column_count, col in enumerate(formatRow):
        if not (isValidFormat(validFormats, col.strip())):
            isValid = False
            print('#### ' + col + ' is not a valid format')
        formatDict[column_count] = col

    # Validate all columns in the format line before returning
    if isValid:
        return True
    else:
        print('### Valid formats are:')
        for frm in validFormats:
            print('### - ' + frm)
        sys.exit(1)
        
def main_prog(infile, outfile, delimiter, hasTitle, hasFormatLine, validFormats):

    if sys.platform == 'OpenVMS':
        infile  = vms.decc.from_vms(infile)
        outfile = vms.decc.from_vms(outfile)
        infile  = infile[0]
        outfile = outfile[0]
        
    workbook = xlsxwriter.Workbook(outfile)
    worksheet = workbook.add_worksheet()
    
    formatDict = {}
    
    
    # If a format line is provided, is must be the first line
    if (hasFormatLine > 0):
        rowNumFormat = 0
    else:
        rowNumFormat = -1

    # If a header line is provided, it is the first line after any format line, so first or second line    
    if (hasTitle > 0):
        rowNumTitle = rowNumFormat + 1
    else:
        rowNumTitle = -1
        

    with open(infile, encoding=input_file_encoding) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=delimiter)
        row_count_infile = 0
        row_count_outfile = 0
        format = 'generic'
        for row in csv_reader:
            
            # Validate format line if any
            if (row_count_infile == rowNumFormat):
                validateFormats(validFormats, row, formatDict)
            else:
                for column_count, col in enumerate(row):
                    # Print header line if any
                    if (row_count_infile == rowNumTitle):
                        worksheet.write(row_count_outfile, column_count, col)
                    elif hasFormatLine > 0:
                       format = formatDict[column_count].strip()
                       if (format == 'int'):
                            worksheet.write(row_count_outfile, column_count, int(col))
                       elif (format == 'float'):
                            worksheet.write(row_count_outfile, column_count, float(col))
                       else:
                            worksheet.write(row_count_outfile, column_count, col)
                    else:
                        worksheet.write(row_count_outfile, column_count, col)
                row_count_outfile += 1
            row_count_infile += 1


    workbook.close()


def main(argv):
    
    inputfile = ''
    outputfile = ''
    delimiter = ''
    hasTitle = 1
    hasFormatLine = 0
    
    argv = sys.argv[1:]
    try:
    #   Colon (:) after option (hio (help/input file/output file) means mandatory)
    #    opts, args = getopt.getopt(argv,"hi:o:",["ifile=","ofile="])
        opts, args = getopt.getopt(argv,"hi:o:d:t:f:",["ifile=","ofile=","delim=","title=","format="])
    except getopt.GetoptError:
        print('')
        print('xlsx_main.py -i <inputfile> -o <outputfile> -d <delimiter> -t <title 1=y, 0=n> -f <format 1=y, 0=n>')
        sys.exit(2)

    for opt, arg in opts:
        print('opt: ' + opt + ' arg: ' + arg)
        if opt == '-h':
            print('xlsx_main.py -i <inputfile> -o <outputfile> -d <delimiter> -t <title 1=y, 0=n> -f <format 1=y, 0=n>')
            sys.exit()
        elif opt in ("-i", "--ifile"):
            inputfile = arg
            print(inputfile + ' '  + arg)
        elif opt in ("-o", "--ofile"):
            outputfile = arg
        elif opt in ("-d", "--delim"):
            delimiter = arg
        elif opt in ("-t", "--title"):
            print('hasTitle: ' + arg + '')
            hasTitle = arg
        elif opt in ("-f", "--format"):
            print('hasFormatLine: ' + arg + '')
            hasFormatLine = arg
        else:
            print('### Unknown option: ' + opt)


# Valid formats
    validFormats = ('generic', 'float', 'int')


# Default values     
    if delimiter == '':
        delimiter = ';'
    if inputfile == '':
        # inputfile = 'python_csv_test-ansi_header.txt'
        # inputfile = 'python_csv_test-ansi_format.txt'
        inputfile = 'python_csv_test-ansi_format_error.txt'
        # inputfile = 'python_csv_test-ansi_format_and_header.txt'
    if outputfile == '':
        outputfile = 'python_csv_test.xlsx'
    # hasTitle = 1
    # hasFormatLine = 1
    print()
    print('Parameters:')
    print('Input file is ', inputfile)
    print('Output file is ', outputfile)
    print('Delimiter: ', delimiter )
    print('hasTitle: ', hasTitle)
    print('hasFormatLine: ', hasFormatLine)
    
    
    main_prog(inputfile, outputfile, delimiter, int(hasTitle), int(hasFormatLine), validFormats)

if __name__ == "__main__":
   main(sys.argv[1:])
