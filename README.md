# Code breakdown

## Libraries used.

Openpyxl is used to import the Excel spreadsheet data. I can parse data from each row for the merges.

Mailmerge is used to populate data into the work doc using mergefields.<br>
If you get an error when you import this package, try uninstalling mailmerge then reinstalling docx-mailmerge. <br>
https://pypi.org/project/docx-mailmerge/

datetime is used to format the date objects.


```python
from openpyxl import load_workbook #used to import the Excel Data
from mailmerge import MailMerge #used for merge tags. If getting error, uninstall and install docx-mailmerge
from datetime import datetime #used to work with date times
```

## Excel Setup

`wb` is loading the Excel workbook. From there, `sheet` selects the tab and `max_row` represent <br>
how many rows of data is in the sheet.

`max_row` is needed when looping through the rows of data.


```python
# Setting up Excel sheet variables
wb = load_workbook('SampleData.xlsx') #open excel workbook
sheet = wb['SalesOrders'] #Tab to get information
max_row = sheet.max_row #count of all of the rows
```

## Unique reps
Looping through all the possible sales reps, then making a distinct list using `list(set())`


```python
# Getting Unique reps. Need to make each of their reports
rep_list = []
for cell_row in range(2 , max_row+1):
    rep = sheet.cell(row = cell_row, column = 3).value
    rep_list.append(rep)
unique_rep_list = list(set(rep_list)) #getting unique list of reps
```

## Loop: interating sales reps
To create a document for each sales rep. We need to use the word document template for each rep. <br>
The `sales_history_list` and `raw_subtotal_list` list will store information for the merge:

- the `sales_history_list` will store nested dictionaries to merge data to the word doc
- the `raw_subtotal_list` will store the individual subtotals, which is needed to calculate the `total`


```python
for rep in unique_rep_list:
    sales_history_list = [] #needed to create the docs dynamically
    raw_subtotal_list = [] #needed to calculate the total
    
    # Setting up Word document variables. Need to reuse template for each rep
    template_doc = "OrderTemplate.docx"
    word_doc = MailMerge(template_doc)    
```

## Nested Loop: Formating dates and numbers for reps

For every row of data that pretains to a rep (`if current_rep == rep:`), we format clean the dates and subtotal

The dates are formatted to show as 03/17/2020, using datetime's `strftime` function. <br>
This converts the datetime into a formatted string for the merge.

the subtotal is formatted to use commas and cents (example 1,278.25) using python built-in `format` method.


```python
for cell_row in range(2 , max_row+1):
        #looping to check the current rep in spreadsheet
        current_rep = sheet.cell(row = cell_row, column = 3).value

        #Checking to see if line item is for rep
        if current_rep == rep:
            #formating date
            raw_date_time = sheet.cell(row = cell_row, column = 1).value #unformatted datetime as a string
            clean_date_time = raw_date_time.strftime("%m/%d/%Y") #converts datetime back into formatted string
            
            #Formatting subtotals
            raw_subtotal = sheet.cell(row = cell_row, column = 7).value
            raw_subtotal_list.append(raw_subtotal) #appending raw number for the total calculation
            
            #convert the number into a string and format (example 1,278.25)
            clean_subtotal = "{:,.2f}".format(raw_subtotal) 
            
```

## Nested Loop (cont.) : interating the sales line items created in table
`product_dict` is storing each row of datas from the spreadsheet for each rep. <br>
each dictionary of data is appended to the `sales_history_list`


```python
           #Appending product as a dict into a list, which will be merged as a table
            product_dict = {
            'Date' : clean_date_time,
            'Item' : str(sheet.cell(row = cell_row, column = 4).value),
            'Quantity' : str(sheet.cell(row = cell_row, column = 5).value),
            'Cost' : str(sheet.cell(row = cell_row, column = 6).value),
            'Subtotal' : clean_subtotal
            }

            #Appending dicts to merge as a table
            sales_history_list.append(product_dict)
```

## Dynamically creating order summaries using merged fields
Calculate the total of the subtotal

afterwards, merge the name and total calculated for each rep. then the `sales_history_list` that contains the dicts <br>
of data are merged into the table.

Last, the document is created and named using rep name.


```python
    # summing raw numbers into a total
    total = sum(raw_subtotal_list)

    # Merging the name and formatting totals  
    word_doc.merge(
    Name = rep,
    Total = "{:,.2f}".format(total)
    )

    word_doc.merge_rows('Date', sales_history_list) #merge which creates table
    word_doc.write(f'Invoice for {rep}.docx') #merge which creates Word doc and name it
```

## Full Code


```python
from openpyxl import load_workbook #used to import the Excel Data
from mailmerge import MailMerge #used for merge tags. If getting error, uninstall and install mailmerge-docx
from datetime import datetime #used to work with date times

# Setting up Excel sheet variables
wb = load_workbook('SampleData.xlsx') #open excel workbook
sheet = wb['SalesOrders'] #Tab to get information
max_row = sheet.max_row #count of all of the rows

# Getting Unique reps. Need to make each of their reports
rep_list = []
for cell_row in range(2 , max_row+1):
    rep = sheet.cell(row = cell_row, column = 3).value
    rep_list.append(rep)
unique_rep_list = list(set(rep_list)) #getting unique list of reps

# For each rep, create their order reports
for rep in unique_rep_list:
    sales_history_list = [] #needed to create the docs dynamically
    raw_subtotal_list = [] #needed to calculate the total
    
    # Setting up Word document variables. Need to reuse template for each rep
    template_doc = "OrderTemplate.docx"
    word_doc = MailMerge(template_doc)    

    for cell_row in range(2 , max_row+1):
        #looping to check the current rep in spreadsheet
        current_rep = sheet.cell(row = cell_row, column = 3).value

        #Checking to see if line item is for rep
        if current_rep == rep:
            #formating date
            raw_date_time = sheet.cell(row = cell_row, column = 1).value #unformatted datetime as a string
            clean_date_time = raw_date_time.strftime("%m/%d/%Y") #converts datetime back into formatted string
            
            #Formatting subtotals
            raw_subtotal = sheet.cell(row = cell_row, column = 7).value
            raw_subtotal_list.append(raw_subtotal) #appending raw number for the total calculation
            
            #convert the number into a string and format (example 1,278.25)
            clean_subtotal = "{:,.2f}".format(raw_subtotal) 
            
            #Appending product as a dict into a list, which will be merged as a table
            product_dict = {
            'Date' : clean_date_time,
            'Item' : str(sheet.cell(row = cell_row, column = 4).value),
            'Quantity' : str(sheet.cell(row = cell_row, column = 5).value),
            'Cost' : str(sheet.cell(row = cell_row, column = 6).value),
            'Subtotal' : clean_subtotal
            }

            #Appending dicts to merge as a table
            sales_history_list.append(product_dict)

    # summing raw numbers into a total
    total = sum(raw_subtotal_list)

    # Merging the name and formatting totals  
    word_doc.merge(
    Name = rep,
    Total = "{:,.2f}".format(total)
    )

    word_doc.merge_rows('Date', sales_history_list) #merge which creates table
    word_doc.write(f'Invoice for {rep}.docx') #merge which creates Word doc and name it
            
```

## References for word doc

Reference: <br>
https://pbpython.com/python-word-template.html


MailMerge Example: <br>
https://answers.microsoft.com/en-us/msoffice/forum/all/mail-merge-is-an-if-then-symbol-possible/0d19a5da-856f-4466-90a6-f7cf6a668339

Marking Merge Fields:
- Key command
> On windows: Ctrl-F9 <br>
>Mac: Cmd-F9 on a Mac <br>
>Once this is done, input MERGEFIELD, as such:
>`{MERGEFIELD Name}`

- Using Word.
>Under <b>Insert</b>, go to <b>Field</b> under the <b>Quick Parts</b> button
![image.png](attachment:image.png)

> <br>
>In the <b>Field window</b>, select <b>Merge Field</b> in the field names and input the name of the field in the <b>Field Name</b> text box

![Merge%20field%20image.png](attachment:Merge%20field%20image.png)

>Once down, it will look like this:
![NewMergeField.png](attachment:NewMergeField.png)
