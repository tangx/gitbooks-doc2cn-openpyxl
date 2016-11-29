# 简单用例 
[ Simple Usage ]( https://openpyxl.readthedocs.io/en/default/usage.html )

----

## 编辑工作簿

```python
>>> from openpyxl import Workbook
>>> from openpyxl.compat import range     # import后的range 将覆盖标准库中的range函数
>>> from openpyxl.cell import get_column_letter
>>>
>>> wb = Workbook()     # 创建一个workbook
>>>
>>> dest_filename = 'empty_book.xlsx'       # 保存使用的位置和文件名
>>>
>>> ws1 = wb.active     # 获取工作表 (创建新工作簿时始终会有一个工作表)
>>> ws1.title = "range names"   # 为工作命名
>>>
>>> for row in range(1, 40):    # 向工作表写入40行
...     ws1.append(range(600))  # 向工作表写入600列
>>>
>>> ws2 = wb.create_sheet(title="Pi")   # 创建第二个工作表，并命名为Pi
>>>
>>> ws2['F5'] = 3.14        # 给单元格'F5'赋值
>>>
>>> ws3 = wb.create_sheet(title="Data") # 创建第三个工作表，并命名为Data
>>> for row in range(10, 20):
...     for col in range(27, 54):
...         _ = ws3.cell(column=col, row=row, value="%s" % get_column_letter(col)) # 将单元格的列名作为单元格的值
>>> print(ws3['AA10'].value)
AA
>>>
>>> ws4=wb.create_sheet('NEW_DATA')
>>> for row in range(1,10):
        for col in range(1,10):
            # 如果没有使用赋值语句 "_ = .... ",结果如下
            ws4.cell(column=col,row=row, value="%s" % get_column_letter(col))
            
<Cell NEW_DATA.A1>
<Cell NEW_DATA.B1>
<Cell NEW_DATA.C1>
<Cell NEW_DATA.D1>
<Cell NEW_DATA.E1>
...
>>> 
>>> wb.save(filename = dest_filename)       # 保存文件
```


### 读取xltx模板并另存为xlsx文件

```python
>>> from openpyxl import load_workbook
>>>
>>>
>>> wb = load_workbook('sample_book.xltx')      # 读取xltx文件
>>> ws = wb.active      # 获取工作表
>>> ws['D2'] = 42       # 为单元格赋值
>>>
>>> wb.save('sample_book.xlsx') # 保存文件
>>>
>>> # 你也可以覆盖当前的文件模板
>>> # wb.save('sample_book.xltx')
```


## 读取xltm模板并另存为xlsm文件

```python
>>> from openpyxl import load_workbook
>>>
>>>
>>> wb = load_workbook('sample_book.xltm', keep_vba=True) # keep_vba=True 保留vba代码
>>> ws = wb.active 
>>> ws['D2'] = 42 
>>>
>>> wb.save('sample_book.xlsm') 
>>>
>>> # 你也可以覆盖当前的文件模板
>>> # wb.save('sample_book.xltm')
```

## 读取一个已存在的工作簿

```python
>>> from openpyxl import load_workbook
>>> wb = load_workbook(filename = 'empty_book.xlsx')    # 加载工作簿
>>> sheet_ranges = wb['range names']        # 使用名为"range names"的工作表
>>> print(sheet_ranges['D18'].value)        # 打印单元格的值
3
```


>###**Note**
>>
>> 使用load_workbook时，有几个标记(flag)可以用。
>>
>> 1. guess_types : 当读取单元格信息时，是否启用推测单元格格式，默认为`不启用 False`。
>>
>> 2. data_only ：控制读取单元格时是否保留公式(formulae)，默认为`保留 False`。或者直接读取单元格所存的最终值。
>>
>> 3. keep_vba : 是否保留Visual Basic元素功能，默认为`不保留 False`。即使VB元素被保留，也不能被编辑。

<br>

>###**Warning**
>>
>> openpyxl 目前不能完全读取Excel文件中的所有项目， 因此在打开和保存过程中使用相同的文件名，图片和图表会被丢失。
>> 


<br>


## 使用数字格式

```python
>>> import datetime
>>> from openpyxl import Workbook
>>> wb = Workbook(guess_types=True)         #创建一个工作簿，并且使用单元格格式
>>> ws = wb.active
>>> # set date using a Python datetime
>>> # 使用日期格式
>>> ws['A1'] = datetime.datetime(2010, 7, 21)
>>>
>>> ws['A1'].number_format
'yyyy-mm-dd h:mm:ss'
>>>
>>> # set percentage using a string followed by the percent sign
>>> # 使用百分数格式，需要在数字后面添加一个百分号。
>>> ws['B1'] = '3.14%'
>>>
>>> ws['B1'].value
0.031400000000000004
>>>
>>> ws['B1'].number_format      # 显示单元格B1的格式
'0%'
```

## 使用公式

```python
>>> from openpyxl import Workbook
>>> wb = Workbook()
>>> ws = wb.active
>>> # add a simple formula
>>> # 是单元格的值为一个公式。实际上就是一个可以解析的字符串
>>> ws["A1"] = "=SUM(1, 1)"
>>> wb.save("formula.xlsx")
```

>**Warning**
>>
>> 函数名称必须使用英文字母；函数的参数必须使用`逗号( , )`分隔，不能使用其他符号，比如半角冒号。

<br>

openpyxl 不会执行公式，但可能会检查公式名称是否正确：

```python
>>> from openpyxl.utils import FORMULAE
>>> "HEX2DEC" in FORMULAE       # 查看公式名称是否在FORMULAE中
True
>>> print type(openpyxl.utils.FORMULAE)
<type 'frozenset'>
```

If you’re trying to use a formula that isn’t known this could be because you’re using a formula that was not included in the initial specification. Such formulae must be prefixed with `xlfn.` to work.

如果你尝试使用一个并不常用的公式，可能这个公式没有包含在初始规范中。这种公式必须使用前缀`xlfn.`才能正常工作。


## 合并/拆分单元格

```python
>>> from openpyxl.workbook import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> ws.merge_cells('A1:B1')         # 合并单元格
>>> ws.unmerge_cells('A1:B1')       # 拆分单元格
>>>
>>> # 或者使用单元格坐标进行操作
>>> ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=4)
>>> ws.unmerge_cells(start_row=2,start_column=1,end_row=2,end_column=4)
```


## 插入图片

```python
>>> from openpyxl import Workbook
>>> from openpyxl.drawing.image import Image
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>> ws['A1'] = 'You should see three logos below'
>>> # create an image
>>> img = Image('logo.png')
>>> # add to worksheet and anchor next to cells
>>> # 在单元格旁边插入图片
>>> ws.add_image(img, 'A1')
>>> wb.save('logo.xlsx')
```


## Fold columns (outline)
## 折叠列 (outline)

```python
>>> import openpyxl
>>> wb = openpyxl.Workbook(True)
>>> ws = wb.create_sheet()
>>> ws.column_dimensions.group('A','D', hidden=True)
>>> wb.save('group.xlsx')
```

----

^[ <-Previous ]( tutorial.md )  |  [ Next-> ]( ./charts/introduction.md )|