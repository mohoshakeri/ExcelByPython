# Excel by python

### A class that helps you by its functions to work easier with Excel data

#

## Install package

```bash
pip install openpyxl
```

#

## Create object

```python
from Excel import Excel


work_book =  Excel("file.xlsx") #or an excel file path
```

- if file not exist its will be create.

#

# Let's go!

## Write on excel

### 1. Write on a column

```python
sheet = "Numbers"
data = [12,15,18,21]
work_book.write_on_column(sheet, "B", data)

```

### Result :

|       | A   | B   | C   |
| ----- | --- | --- | --- |
| **1** |     | 12  |     |
| **2** |     | 15  |     |
| **3** |     | 18  |     |
| **4** |     | 21  |     |

* **row_start** is optional, its an Excel row index that the function will start from there. Ex 12 **between 1 and 1048576** (default = 1)

* **center_style** is optional, if equal True, styles of the cells will be middle (default = False)
---


### 2. Write on a row

```python
sheet = "Numbers"
data = [12,15,18]
work_book.write_on_row(sheet, 3, data)
```

### Result :

|       | A   | B   | C   |
| ----- | --- | --- | --- |
| **1** |     |     |     |
| **2** |     |     |     |
| **3** | 12  | 15  | 18  |
| **4** |     |     |     |

* **col_start** is optional, its an Excel column that the function will start from there. Ex AB **between A and XFD** (default = 1)

* **center_style** is optional, if equal True, styles of the cells will be middle (default = False)

---

### 3. Write on a cell

```python
sheet = "Numbers"
data = 18
column = "B"
row = 3
work_book.write_on_row(sheet, column, row data)
```

### Result :

|       | A   | B   | C   |
| ----- | --- | --- | --- |
| **1** |     |     |     |
| **2** |     |     |     |
| **3** |     | 18  |     |
| **4** |     |     |     |

* **center_style** is optional, if equal True, styles of the cells will be middle (default = False)

#

## Read data

### Example sheet

|       | A   | B    | C   |
| ----- | --- | ---- | --- |
| **1** | Id  | Name | Age |
| **2** | 1   | Ali  | 35  |
| **3** | 2   | Amir | 12  |
| **4** | 3   | Reza | 27  |

---

### 1. Read a column

```python
sheet = "Users"
data = work_book.read_column(sheet, "B")
for item in data:
    print(item)

>>> "Age"
>>> "35"
>>> "12"
>>> "27"
```
* **row_start** is optional, its an Excel row index that the generator will start from there. Ex 12 **between 1 and 1048576** (default = 1)

* **row_end** is optional, its an Excel row index that the generator will break in there. Ex 12 **between 1 and 1048576** (default = 1048576)
---

### 2. Read a row

```python
sheet = "Users"
data = work_book.read_row(sheet, 3)
for item in data:
    print(item)

>>> "2"
>>> "Amir"
>>> "12"
```
* **col_start** is optional, its an Excel column that the generator will start from there. Ex AB **between A and XFD** (default = A)

* **col_end** is optional, its an Excel column that the generator will break in there. Ex AB **between A and XFD** (default = XFD)
---

### 3. Read a cell

```python
sheet = "Users"
column = "B"
row = 3
data = work_book.read_cell(sheet, column, row)
print(data)

>>> "Amir"
```
