# README

## Table of contents
* [PrismBook.py](#prismbook.py)
* [PrismBook](#prismbook)
	* [Motivation](#motivation)
	* [Attributes](#attributes)
	* [Methods](#methods)

# <span>PrismBook.py</span>
*PrismBook.py* is a runnable Python script that exists as a  [Python tool](https://help.alteryx.com/2018.3/Python.htm) in the PRISM Alteryx workflow. The script creates buyer-channel PRISM workbooks, automating the workbook creation process. The `PrismBook` class contained in this script wraps [Pandas DataFrame](https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.html) and [XlsxWriter](https://xlsxwriter.readthedocs.io/) operations, providing an interface for easy workbook modification. Details can be found in the [PrismBook](#PrismBook) section. 

*PrismBook.py* contains:
* The definition of the ```PrismBook``` class.
* Auxiliary functions for date formatting and data preprocessing.
* A main function to run the script.

Docstrings for each function and class in *PrismBook.py*  are available, and can be called with  ```help(name)``` as usual. 

When *PrismBook.py* is run, a file structure is created to contain the buyer-channel workbooks generated. A sample output folder directory might look like this:

    PRISM_12_10_2019/
        CS/
            CS_Blaine_Gobler.xlsm
            CS_Brad_LeVan.xlsm
            CS_Brian_Greene.xlsm
        EC/
            EC_Brie_Connolly.xlsm
            EC_Hayli_Taylor.xlsm
            EC_Jaime_Tackett.xlsm
        FO/
            FO_Amy_Roos.xlsm
            FO_Anna_Parker.xlsm
            FO_Katie_Rodgers.xlsm

# PrismBook
In a situation where column addition/removal/reshuffling requests are made with short turnaround, it behooves us to develop a framework for developing robust column-specific formula and formatting.

* [Motivation](#Motivation)
* [Attributes](#Attributes)
* [Methods](#Methods)

## Motivation
PrismBook was developed to implement and maintain the PRISM workbook, with the ability to implement changes in a short time. XlsxWriter is not robust to column addition or reshuffling. 

An example of formula writing in XlsxWriter:
```python
for row in range(start_idx, end_idx):
    worksheet.write_formula(row, 12, "=VLOOKUP(A%d, ’Sheet1'!A4:X1000, 24,FALSE)" %(row+1))
```
The following aspects are problematic:
* Explicit indices for any cell reference
* Formulas written one cell at a time in for-loop
* Formula must be a string, formatting required



## Attributes

### `self.sheet_dict`: 
Dictionary of format `dict(sheet_name, sheet_df)`. 

`self.sheet_props`: Dictionary of format `dict(sheet_name, dict(col_name, col_properties))`.

```python
{'sheet_one': 
    {'class_no': {'locked': None,
                  'display_header': None,
                  'formatting': {'border': 1,
                                'bg_color': '#F2F2F2',
                                'num_format': '0',
                                'display_header': 'Class no.',
                                'locked': True,
                                'hidden': False,
                                'level': 1
                                },
                  'formula': None
                },
    ...
    },
'sheet_two':  ...
}
```

### `self.offset_dict`
Dictionary of format `dict(sheet_name, offset)`. Offset is an integer specifying the number of rows from the top. Offset is the height of the maximum header present + 1. This number is used to specify the start and end row of headers, and makes it so the dataframe in `self.sheet_dict` begins at the correct row and that formulas and formatting are likewise applied correctly. 

### `self.header_dict`
Keeps track of any existing headers on the sheets. A dictionary of format `dict(sheet_name, header_name, header_formatting_list)`, where `header_formatting_list` is a list of format `[[start_col_name, end_col_name], height, format_dict]`.
```python
{'sheet_three': 
    {'Days of lateness': [['OTA Strict & Moderate', 'DOL 30+'],5,{'bold': 1, 'align': 'center', 'border': 1, 'bg_color': '#92D050'}],
    ...
    },
 
 'sheet_four': ...
}
```

### `self.excel_dict`
A dictionary of format `dict(numeric_idx, alphabetical_idx)`. Provides easy conversion between a numeric column index to the alphabetical index as it would appear in Excel. Zero-indexed.

```python
{
    0: 'A',
    1: 'B',
    2: 'C',
    ...

    26: 'AA',
    27: 'AB',
    ...
    
    52: 'BA'
}
```

### `self.valid_dict`
Contains data validation information. Dictionary of format `dict(sheet_name, dict(col_name, valid_dict)) `. If no validation present on sheet, `{}` is value for sheet name. 

```python
{'sheet_four': 
    {'submit': {'validate': 'list', 'source': ['Yes','No']}},
 'sheet_one': {},
 'sheet_two': {},
 'sheet_three': {}
 }
```

## Methods
* [`insert()`](#insert())
* [`insert_range()`](#insert_range())
*  [`update_prop()`](#`update_prop()`)
* [`update_specific_prop()`](#`update_specific_prop()`)
* [`update_specific_format()`](#`update_specific_format()`)
* [`delete()`](#`delete()`)
* [`rearrange()`](#`rearrange()`)
* [`move()`](#`move()`)
* [`evaluate_formulas()`](#`evaluate_formulas()`)
* [`add_validation()`](#`add_validation()`)
* [`add_header()`](#`add_header()`)
* [`to_excel()`](#`to_excel()`)
* [`close_workbook()`](#`close_workbook()`)
* [`change_display_headers()`](#`change_display_headers()`)

### `insert()`
Inserts a list of columns starting at a specified index. Optionally specify column properties for these columns.
#### Parameters
  * **sheet_name** *(str)* : Sheet name as it appears in `self.sheet_dict`.
  * **col_names** *(list)* :  Column names to insert.
  * **index** *(int, default=None)* : First index to insert column.
  * **col_props** *(dict, default=None)* : Property dictionary.
  * **default_val** *(str, default='')* : Default value to fill in Pandas DataFrame. 
  * **to_end** *(bool, default=False)* : If this is True, adds columns to end of dataframe. If this is True, index will be ignored if specified.

#### Examples
```python
#Example 1
 test_book.insert('sheet_one', ['p_value', 'ros_ratio_his'], col_props={'p_value': {'formula': '=_xlfn.NORM.S.DIST(-ABS({log_ros_ratio}-{mu})/{sigma}, TRUE)*2'},'ros_ratio_his': {'formula': 'EXP({mu})'}}, to_end=True)

#Example 2
test_book.insert('sheet_one', ['Use Adjusted FC'], col_props = {'Use Adjusted FC': {'formula': None, 'formatting': {'bg_color': 'white', 'locked': False}}}, to_end=True, default_val = 'X')

#Example 3
test_book.insert('sheet_three', ['month_number', 'concat'], to_end=True)
```
### `insert_range()`
#### Parameters
* **sheet_name** *(str)* : Sheet name as it appears in `self.sheet_dict`.
* **column_name** *(list)* : List of names for columns to be inserted.
* **props** *(dict)* : Skeleton property dictionary for these columns. Insert string formatting to be provided by `ref_dict`. 
* **ref_dict** *(dict)* : Dictionary of (ref_str, list), where for each `ref_str`, the first element of its list is the reference for the first column, and so on. 
* **start_idx** *(int, default=None)* : First index to insert column.
* **to_end** *(bool, default=False)* : Inserts columns to end. `start_idx` is ignored if provided.
#### Examples
```python
ref_col = ['mol1','mol2','mol3','mol4','mol5','mol6']
mol_cols = ['user_mol1', 'user_mol2','user_mol3','user_mol4','user_mol5','user_mol6']
test_book.insert_range('sheet_three', mol_cols, {'formula': '={X}', 'formatting': {'locked': False, 'bg_color': 'white'}}, {'X': ref_col}, start_idx=17)
```

### `update_prop()`
Replaces entire property dictionary with new one by updating the entry in `self.sheet_props`. 
#### Parameters
* **sheet_name** *(str)* : Sheet name as it appears in `self.sheet_dict`.
* **col_name** *(str)* : Name of column whose properties to update; the key to be looked up in `self.sheet_props.get(sheet_name)`.
* **new_prop** *(dict)* : Property dictionary of column. 

#### Examples
```python
test_book.update_prop('sheet_four', 'x_check',{'formula': '={remaining_fraction}*{sheet_one|init_ros, on=article}'})
```

### `update_specific_prop()`
Updates a single property in the property dictionary. 
#### Parameters
* **sheet_name** *(str)* : Sheet name as it appears in `self.sheet_dict`.
* **col_name** *(str)* : Name of column whose property to update; the key to be looked up in `self.sheet_props.get(sheet_name)`.
* **prop** *(str)* : Name of property to update. The key to be looked up in `self.sheet_props.get(sheet_name).get(col_name)`.
* **newval** *(variable type)* : New property value; depending on the property the type varies. 

#### Examples
```python
test_book.update_specific_prop('sheet_three', 'Buyer', 'formula', '={sheet_one|buyer, on=article}')
```
### `update_specific_format()`
Updates a dictionary of specified formats within the `formatting` property in the property dictionary to be the same value for a list of columns. 
#### Parameters
def update_specific_format(self, sheet_name, cols, new_formats):
* **sheet_name** *(str)* : Sheet name as it appears in `self.sheet_dict`.
* **cols** *(list)* : Name of columns whose formatting property to update; the key to be looked up in `self.sheet_props.get(sheet_name)`.
* **new_formats** *(dict)* : Formats to be modified, `dict(format_name, format_val)`.  

#### Examples
```python
test_book.update_specific_format('sheet_four', ['buyer', 'class_no', 'class_desc','subclass_desc', 'gender', 'product_division'], {'hidden': True,'level': 1})
```

### `delete()`
Deletes a column in a pandas dataframe in `self.sheet_dict` and removes corresponding column from sheet entry in `self.sheet_props`.

#### Parameters
* **sheet_name** *(str)* : Sheet name as it appears in `self.sheet_dict`.
* **col_name** *(str)* : Column to be deleted.
#### Examples
```python
test_book.delete('sheet_one', 'useless_col')
```
### `rearrange()`
Redefines the column order in a sheet; all column names should be included. 
#### Parameters
    def rearrange(self, sheet_name, new_order):
* **sheet_name** *(str)* : Sheet name as it appears in `self.sheet_dict`.
* **new_order** *(list)* : List containing new column order.

#### Examples
```python
test_book.rearrange(sheet_one, ['article', 'article_description', 'buyer', 'class_no', 'class_desc',
       'subclass_desc', 'gender', 'product_division', 'product_type',
       'wks_remaining', 'total_sales_td', 'remaining_seasonal_fc',
       'price_group', 'fcst_qty', 'period_flag', 'init_ros', 'log_ros_ratio',
       'mu', 'sigma', 'p_value',
       'ros_ratio_his', 'sell_out_week_ending_onhand', 'Adjusted FC', 'Use Adjusted FC'])
```
### `move()`
Moves a single column in a dataframe. 
def move(self, sheet_name, col_names, index = -1, to_end = False):

#### Parameters
* **sheet_name** *(str)* : Sheet name as it appears in `self.sheet_dict`.
* **col_names** *(list)* : Columns to move, do not have to be adjacent originally. They will be placed adjacently starting at the specified index. 
* **index** *(int, default = -1)* : Index to insert columns at. Columns are placed adjacent to one another in the order specified by `col_names`.
* **to_end** *(bool, default=False)* : Inserts columns to end. `index` is ignored if provided.

#### Examples
```python
test_book.move('sheet_four', ['x_check'], to_end=True)
```
### `evaluate_formulas()`
Evaluates the `formula` property in `self.sheet_props`, converting column names in curly braces to alphabetical indices as they would appear in Excel. Appends `{row}` to the indices, unless string in curly braces is a VLOOKUP.

Users enter formulas where cell references are curly braces. Formulas are intended to be **column-specific**. References in each formula are also column-specific. In `to_excel()`, we iterate through each row of a column, writing the formula in each cell, appending the current row number to each column reference index. This means that only same-row column references are supported.

The resulting formula string is written into the specific cell within the `worksheet`. The process is illustrated as follows:
<img src="https://imgur.com/whkm4Gc.jpg" width="500" height="170">

The third step illustrates a single row in the for-loop; note how all the column references are for the same row. 

VLOOKUP is a common function, so we simplify the syntax to be `lookup_sheet_name|lookup_col, on=index_col`. Here is an example:

<img src="https://imgur.com/qJSIh2z.jpg" width="470"  height="180">

#### Examples
```python
test_book.evaluate_formulas()
```

### `add_validation()`
Adds data validation to a single column.
def add_validation(self, sheet_name, col, col_valid_dict):

#### Parameters
* **sheet_name** *(str)* : Sheet name as it appears in `self.sheet_dict`.
* **col** *(str)* : Column to insert validation in.
* **valid_dict** *(dict)* : Validation dictionary, the `options` parameter of the `data_validation()` function in [XlsxWriter](https://xlsxwriter.readthedocs.io/worksheet.html#data_validation).

#### Examples
```python
#Example 1
user_adj_con_rdp = ['user_con_rdp1','user_con_rdp2','user_con_rdp3','user_con_rdp4','user_con_rdp5','user_con_rdp6']
valid_str = {'validate': 'decimal',
'criteria': '>',
'value': '=IF(COLUMN()+1-COLUMN({con_adj_rdp1}) <> 1+DATEDIF(DATEVALUE({season_start_date}),{last_run_date},"m"),0,{sales_td_thismonth})',
'input_title': 'Input note:',
'input_message': 'Adjustments must be greater than zero. Adjustments for the current month must exceed Sales MTD.',
'ignore_blank': True}
for col in user_adj_con_rdp:
    test_book.add_validation('sheet_four', col, valid_str)

#Example 2
test_book.add_validation('sheet_four', 'submit', {'validate':'list', 'source':['Yes','No']})
```

### `add_header()`
Adds a merged header across multiple columns.  

    def add_header(self, sheet_name, name, col_range, height, format_dict):

#### Parameters
* **sheet_name** *(str)* : Sheet name as it appears in `self.sheet_dict`.
* **name** *(str)* : Name of header. 
* **col_range** *(list)* : Two element list specifying first and last column covered by header.
* **height** *(int)* : Height of header.
* **format_dict** *(dict)*: Format dictionary of header. Of the same syntax as the `formatting` entry in `self.sheet_props`. 

#### Examples
```python
test_book.add_header('sheet_three', 'Days of lateness', ['OTA Strict & Moderate', 'DOL 30+'], 5,  {'bold': 1,'align': 'center','border': 1,'bg_color': '#92D050'})
```
<img src="https://imgur.com/jxeko3I.jpg">

### `to_excel()`
Writes each pandas dataframe in `self.sheet_dict` to its own XlsxWriter `worksheet` object in the same `workbook` object, formatting columns with properties in `self.sheet_props`. Adds headers in `self.header_dict` and data validation for columns in `self.valid_dict`. Finally, protects each sheet in the `workbook`. 

#### Parameters
* **filename** *(str)* : Name of the excel workbook. The file extension MUST be `.xlsm`.

#### Examples
```python
test_book.to_excel('PRISM_test.xlsm')
```
### `close_workbook()`
Closes and writes workbook by running the `workbook.close()` function in [XlsxWriter](https://xlsxwriter.readthedocs.io/workbook.html). Run once when everything is finalized. Also writes `.json` files for logging current state of workbook, to be used in reingestion.

#### Examples
```python
test_book.close_workbook()
```

### `change_display_headers()`
Changes the display header for specified columns by modifying the `display_header` entry in `self.sheet_props`. The display header is the column header as shown in Excel. 

#### Parameters
* **sheet_name** *(str)* : Sheet name as it appears in `self.sheet_dict`.
* **display_dict** *(dict)* : Dictionary of type `dict(col_name, new_display_header)`. 

#### Examples
```python
#Example 1
test_book.change_display_headers('sheet_two', {'article': 'Article'})

#Example 2
season_start_date = '7/1/2019' #only edit this. 
month_int = dt.strptime(season_start_date, '%m/%d/%Y') 
month_names = calendar.month_name[month_int.month:month_int.month+6] 

test_book.change_display_headers('sheet_four', {x[i]:month_names[i] for i in range(0,6) for x in [rdp_cols, adj_rdp_cols, con_rdp_cols, user_adj_con_rdp, totals, monthlysupp, monthlyfa, og_rdp]})
```





