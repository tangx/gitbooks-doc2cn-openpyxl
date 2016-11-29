# 在内存中编辑workbook

[英文源地址](https://openpyxl.readthedocs.io/en/default/tutorial.html)

----------

## 创建workbook

使用openpyxl并不需要在系统上新建一个文件，只需引用Workbook类即可。


``` python
>>> from openpyxl import Workbook
>>> wb = Workbook()
```

一个Excel文档在创建是至少会有一个标签(worksheet)。因此可以通过使用`openpyxl.workbook.Workbook.active()`方法激活该标签。

```python
>>> ws = wb.active
```

> ### Note
> 该功能使用了`_active_sheet_index`方法，默认值为0。如果你没有改变其值，那么通过此方法始终会获取第一个worksheet。
> 同时，你也可以通过说使用方法`openpyxl.workbook.Workbook.create_sheet()`创建一个新的worksheet。

```python
>>> ws1 = wb.create_sheet() # 在末尾创建 (默认)
# or
>>> ws2 = wb.create_sheet(0) # 在第一个位置插入
```

创建Sheets的时候会按照数字队列自动命名（Sheet,Sheet1,Sheet2, ...）。你可以通过`title`属性随时改变其名称。

```python
ws.title = "New Title"
```

页面的背景色默认为白色.你可以为`ws.sheet_properties.tabColor`属性赋值`RRGGBB颜色编码`进行更改:

```python
ws.sheet_properties.tabColor = "1072BA"
```

一旦你给一个worksheet命名之后，你可以通过将`worksheet的名称`作为`workbook的key`或者使用`openpyxl.workbook.Workbook.get_sheet_by_name()`方法得到它。

```python
>>> ws3 = wb["New Title"]
>>> ws4 = wb.get_sheet_by_name("New Title")
>>> ws is ws3 is ws4
True
```

你可以通过方法`openpyxl.workbook.Workbook.get_sheet_names()`获取说为sheet名称：

```python
>>> print(wb.get_sheet_names())
['Sheet2', 'New Title', 'Sheet1']
```

你可以遍历所有worksheets
```python
>>> for sheet in wb:
...     print(sheet.title)
```

## 使用数据
### 访问单个单元格

现在，我们知道怎样方位一个worksheet了，现在我们来学习如何改变单元格的内容。

我们可以直接通过worksheet的键值（key）方位单元格。

```python
>>> c = ws['A4']
```

这样就可以直接获取A4单元格了。如果该单元格不存在，则会被创建。 也可以直接为单元格赋值。

```python
>>> ws['A4'] = 4
```

也可以使用方法`openpyxl.workbook.Workbook.cell()`进行操作

```
>>> c = ws.cell('A4')
```

同样，你也可以通过使用单元格的行和列标记访问：

```
>>> d = ws.cell(row = 4, column = 2)
```

> ### Note
>> 在worksheet在被创建时存在于内存中的。此时worksheet不包含任何单元格。 单元格将在第一次被访问时被创建。 此功能将不会创建不会访问的对象，从而达到可以节约内存的目的。

<br>

> ### Warning
>> 正式因为此功能，当遍历所有单元格而非直接访问时，将会在内存中创建它们，即使你不为他们赋值。
>>
>> 比如说
>>
>> ```python
>> >>> for i in range(1,101):
>> ...        for j in range(1,101):
>> ...            ws.cell(row = i, column = j)
>> ```
>> 将会在内存中创建100x100个空白单元格。
>>
>> 然后，openpyxl提供了一种方法清空这些不需要的单元格，这个我们将在之后介绍。



### 访问多个单元格 

单元格区块(Ranges of cells)可以通过切片方式访问。

```ptyhon
>>> cell_range = ws['A1':'C2']
```

同样，你也可以使用方法`openpyxl.worksheet.Workbook.iter_rows()`:

```python
>>> tuple(ws.iter_rows('A1:C2'))
((<Cell Sheet1.A1>, <Cell Sheet1.B1>, <Cell Sheet1.C1>),
 (<Cell Sheet1.A2>, <Cell Sheet1.B2>, <Cell Sheet1.C2>))

>>> for row in ws.iter_rows('A1:C2'):
...        for cell in row:
...            print cell
<Cell Sheet1.A1>
<Cell Sheet1.B1>
<Cell Sheet1.C1>
<Cell Sheet1.A2>
<Cell Sheet1.B2>
<Cell Sheet1.C2>
```

如果你需要循环访问一个文件中的所有行或者列，你可以使用`openpyxl.worksheet.Workbook.rows()`属性。

```python
>>> ws = wb.active
>>> ws['C9'] = 'hello world'
>>> ws.rows
((<Cell Sheet.A1>, <Cell Sheet.B1>, <Cell Sheet.C1>),
(<Cell Sheet.A2>, <Cell Sheet.B2>, <Cell Sheet.C2>),
(<Cell Sheet.A3>, <Cell Sheet.B3>, <Cell Sheet.C3>),
(<Cell Sheet.A4>, <Cell Sheet.B4>, <Cell Sheet.C4>),
(<Cell Sheet.A5>, <Cell Sheet.B5>, <Cell Sheet.C5>),
(<Cell Sheet.A6>, <Cell Sheet.B6>, <Cell Sheet.C6>),
(<Cell Sheet.A7>, <Cell Sheet.B7>, <Cell Sheet.C7>),
(<Cell Sheet.A8>, <Cell Sheet.B8>, <Cell Sheet.C8>),
(<Cell Sheet.A9>, <Cell Sheet.B9>, <Cell Sheet.C9>))
```

或者`openpyxl.workbook.Workbook.columns()`属性：

```python
>>> ws.columns
((<Cell Sheet.A1>,
<Cell Sheet.A2>,
<Cell Sheet.A3>,
<Cell Sheet.A4>,
<Cell Sheet.A5>,
<Cell Sheet.A6>,
...
<Cell Sheet.B7>,
<Cell Sheet.B8>,
<Cell Sheet.B9>),
(<Cell Sheet.C1>,
<Cell Sheet.C2>,
<Cell Sheet.C3>,
<Cell Sheet.C4>,
<Cell Sheet.C5>,
<Cell Sheet.C6>,
<Cell Sheet.C7>,
<Cell Sheet.C8>,
<Cell Sheet.C9>))
```

### 数据储存 

Once we have a openpyxl.cell.Cell, we can assign it a value:
当我们获取了一个单元格`openpyxl.cell.Cell`之后，就可以给他赋值了：

```python
>>> c.value = 'hello, world'
>>> print(c.value)
'hello, world'

>>> d.value = 3.14
>>> print(d.value)
3.14
```

你也可以使用type和format接口：

```python
>>> wb = Workbook(guess_types=True)
>>> c.value = '12%'
>>> print(c.value)
0.12

>>> import datetime
>>> d.value = datetime.datetime.now()
>>> print d.value
datetime.datetime(2010, 9, 10, 22, 25, 18)

>>> c.value = '31.50'
>>> print(c.value)
31.5
```

## 保存文件
使用`openpyxl.workbook.Workbook.save`方法是保存`openpyxl.workbook.Workbook`对象的最简单、最安全的方式。

```python
>>> wb = Workbook()
>>> wb.save('balances.xlsx')
```

> ### Warning
> 该操作会覆盖已存在的文件，并且没有warning提示。

<br>


> ### Note
>>Extension is not forced to be xlsx or xlsm, although you might have some trouble opening it directly with another application if you don’t use an official extension.
>>
>>As OOXML files are basically ZIP files, you can also end the filename with .zip and open it with your favourite ZIP archive manager.


你可以使用属性`as_template=True`将文件保存为一个模板(template)。

```python
>>> wb = load_workbook('document.xlsx')
>>> wb.save('document_template.xltx', as_template=True)
```

或者使用属性`as_template=False`(默认)将一个模板/文件保存为一个文件。

```python
>>> wb = load_workbook('document_template.xltx')
>>> wb.save('document.xlsx', as_template=False)
>>> wb = load_workbook('document.xlsx')
>>> wb.save('new_document.xlsx', as_template=False)
```

> ### Warning 
>> 在保存/打开文件时你需要监视(monitor)文档模板内的数据属性和文档扩展，否则在保存后的文档可能无法打开。
>> You should monitor the data attributes and document extensions for saving documents in the document templates and vice versa, otherwise the result table engine can not open the document.


<br>


> ### Note 
>> 以下操作会失败：
>>```python
>> >>> wb = load_workbook('document.xlsx')
>> >>> # Need to save with the extension *.xlsx
>> >>> wb.save('new_document.xlsm')
>> >>> # MS Excel can't open the document
>> >>>
>> >>> # or
>> >>>
>> >>> # Need specify attribute keep_vba=True
>> >>> wb = load_workbook('document.xlsm')
>> >>> wb.save('new_document.xlsm')
>> >>> # MS Excel can't open the document
>> >>>
>> >>> # or
>> >>>
>> >>> wb = load_workbook('document.xltm', keep_vba=True)
>> >>> # If us need template document, then we need specify extension as *.xltm.
>> >>> # If us need document, then we need specify attribute as_template=False.
>> >>> wb.save('new_document.xlsm', as_template=True)
>> >>> # MS Excel can't open the document
>>
>>```


## 加载一个文件
与保存文件一样，你可以通过导入`openpyxl.load_workbook()`开打一个已存在的workbook:

```python
>>> from openpyxl import load_workbook
>>> wb2 = load_workbook('test.xlsx')
>>> print wb2.get_sheet_names()
['Sheet2', 'New Title', 'Sheet1']
```

openpyxl 概览到此结束，你现在可开始 简单用法章节([Simple usage section](./usage.md) )了。
