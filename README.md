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

19-08-2020
Created a second version which doesnt keep working with openpyxl, but instead transfers the data of the xlsx to a list with dictionaries like an JSON file. This way using an API is much easier cause the data is already structured like JSON.
Also easier to program data-changes.
Added functionality: 
- Email transferred to lowercase
- Phonenumber extracted (making sure no -, +, () is in it) and deleting the first zero(s).
- Create a seperate list with all the adresses with no phonenumber after the APIs, for the 'geenTelnr' file.