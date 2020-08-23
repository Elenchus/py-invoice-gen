# Generate invoices from timesheets
A Python invoice generator for contractors, from an Excel timesheet and a Word template

## Software version compatibility
Tested with Python 3.7, package versions as in requirements.txt

## Usage
The meat of the code is within a single function, create_invoice.
The script can also be run from the command line, with arguments input with '--[argname] [value]', e.g.
```
python invoice_generator.py --invoice_no 2
```

The function takes the following arguments:

### timesheet
The xlsx timesheet to use. The first three columns should be: Date; Hours; Description.
Defaults to assets/SampleTimesheet if run from the command line; required argument if the function is called directly.
Note that no file extension should be included.
### invoice_no
Invoice number. Defaults to 1
### rate
Hourly rate to charge. Defaults to 20
### template
Word document to use as an invoice template.
Defaults to ./assets/InvoiceTemplate
Note that no file extension should be included.


## Example
An example of the output is available in assets/SampleTimesheet.pdf