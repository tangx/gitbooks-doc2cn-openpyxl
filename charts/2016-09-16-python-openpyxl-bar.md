# Bar and Column Charts
[Bar and Column Charts](https://openpyxl.readthedocs.io/en/default/charts/bar.html)

----


柱状图会绘制水平或者垂直的数据列。

## Vertical, Horizontal and Stacked Bar Charts
## 垂直、水平、堆栈柱状图

>###**Note**
>
> 以下设置将影响不同类型的柱状图。
> 通过分别设置type的值为`bar或col`从而使用垂直`type=col`或水平`type=bar`柱状图。
>
> 当使用堆栈柱状图时，overlap的值需要设置为100。`overlap=100`
>
> 如果是水平柱状图，x和y轴将颠倒。

<br>


!["Sample bar charts"](http://docs.uyinn.com/openpyxl.readthedocs.io/en/default/_images/bar.png)


```python
from openpyxl import Workbook
from openpyxl.chart import BarChart, Series, Reference

wb = Workbook(write_only=True)
ws = wb.create_sheet()

rows = [
    ('Number', 'Batch 1', 'Batch 2'),
    (2, 10, 30),
    (3, 40, 60),
    (4, 50, 70),
    (5, 20, 10),
    (6, 10, 40),
    (7, 50, 30),
]


for row in rows:            # 添加单元格绘图数据
    ws.append(row)


chart1 = BarChart()         # 创建  2D柱状图 对象
chart1.type = "col"         # 柱状图type=col，垂直图。
chart1.style = 10           # style显示不同的风格
chart1.title = "Bar Chart"                      # 图形名称
chart1.y_axis.title = 'Test number'             # y轴名称
chart1.x_axis.title = 'Sample length (mm)'      # x名称

data = Reference(ws, min_col=2, min_row=1, max_row=7, max_col=3)        # 关联图形数据
cats = Reference(ws, min_col=1, min_row=2, max_row=7)                   # 关联数据分类
chart1.add_data(data, titles_from_data=True)            # 数据标题从titles_from_data=True获取，即数据单元格组的第一行
chart1.set_categories(cats)
chart1.shape = 4
ws.add_chart(chart1, "A10")         # 生成柱状图及定位

from copy import deepcopy           

chart2 = deepcopy(chart1)           # deepcopy ，之后直接修改部分属性
chart2.style = 11
chart2.type = "bar"                 # 水平柱状图
chart2.title = "Horizontal Bar Chart"   

ws.add_chart(chart2, "G10")          # 生成柱状图及定位


chart3 = deepcopy(chart1)           # deepcopy ，之后直接修改部分属性
chart3.type = "col"                 # 水平柱状图
chart3.style = 12
chart3.grouping = "stacked"         # 类型为堆栈
chart3.overlap = 100                # 必须设置内容
chart3.title = 'Stacked Chart'

ws.add_chart(chart3, "A27")         # 生成柱状图及定位


chart4 = deepcopy(chart1)
chart4.type = "bar"                 # 水平柱状图
chart4.style = 13
chart4.grouping = "percentStacked"  # 类型 百分比堆栈
chart4.overlap = 100
chart4.title = 'Percent Stacked Chart'

ws.add_chart(chart4, "G27")

wb.save("bar.xlsx")
```

这样便生成了4种不同类型的柱状图。


## 3D Bar Charts
## 3D柱状图


可以生成3D柱状图。

```python
from openpyxl import Workbook
from openpyxl.chart import (
    Reference,
    Series,
    BarChart3D,
)

wb = Workbook()
ws = wb.active

rows = [
    (None, 2013, 2014),
    ("Apples", 5, 4),
    ("Oranges", 6, 2),
    ("Pears", 8, 3)
]

for row in rows:
    ws.append(row)

data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=4)
titles = Reference(ws, min_col=1, min_row=2, max_row=4)
chart = BarChart3D()                    # 创建 3D柱状图 对象
chart.title = "3D Bar Chart"
chart.add_data(data=data, titles_from_data=True)
chart.set_categories(titles)

ws.add_chart(chart, "E5")
wb.save("bar3d.xlsx")
```

如此便生成了一个简单的 3D柱状图


!["Sample 3D bar chart"](http://docs.uyinn.com/openpyxl.readthedocs.io/en/default/_images/bar3D.png)

----

^[ <-Previous ]( ./bar.md )  |  [ Next-> ]( ./bubble.md )|