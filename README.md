# PyExcell
This is a python script which transform callersdata to CallPro ready data.
Trying to connect this to CDDN and EDM API's.

17-08-2020
Build the start with OpenPyXL.
Importing file, making a columnindex dictionary and reading row by row.
It checks phonenumbers, if there is non, fills the cell with 'checked'

18-08-2020
Build a ZIP cleaner and checker with RE.
Checks the ZIP combination and puts it as DDDD LL.
If there are no 4 digits and 2 letters it puts 'FOUT' (dutch for false) in the cell.
Also updated the phonenumber check: if there is no phonenumber and a false zipcode it returns a different value to the cell