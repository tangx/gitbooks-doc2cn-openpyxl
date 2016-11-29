#面积图

[ Area Charts ](http://docs.uyinn.com/openpyxl.readthedocs.io/en/default/charts/area.html)

----

##2D面积图

Area charts are similar to line charts with the addition that the area underneath the plotted line is filled. Different variants are available by setting the grouping to “standard”, “stacked” or “percentStacked”; “standard” is the default.
 
面积图与线性图类似；描绘的线下面进行填充，用面来表示。通过设置`grouping`不同的值(`chart.grouping="standard"`)绘制不同的面积图，包括"standard","stacked","percentStacked"，默认为 standard。

```python
from openpyxl import Workbook
from openpyxl.chart import (
    AreaChart,
    Reference,
    Series,
)

wb = Workbook()
ws = wb.active

rows = [
    ['Number', 'Batch 1', 'Batch 2'],
    [2, 40, 30],
    [3, 40, 25],
    [4, 50, 30],
    [5, 30, 10],
    [6, 25, 5],
    [7, 50, 10],
]

for row in rows:            # 为worksheet写入多行数据
    ws.append(row)          # 将list作为一行数据写入worksheet中

chart = AreaChart()         # 创建 2D图表 对象
chart.title = "Area Chart"  # 设置图表标题
chart.style = 13            # 设置图标风格
# chart.grouping="standard"   # 设置图表 grouping 类型
chart.x_axis.title = 'Test'     # 设置x轴名称
chart.y_axis.title = 'Percentage'   # 设置y轴名称

cats = Reference(ws, min_col=1, min_row=1, max_row=7)       # 关联分类数据的值
data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=7)    # 关联构图数据的值
chart.add_data(data, titles_from_data=True)     # 为图表添加数据        
chart.set_categories(cats)      # 为图表添加分类

ws.add_chart(chart, "A10")      # 设置图表位置

wb.save("area.xlsx")
```

!["Sample area charts"](http://docs.uyinn.com/openpyxl.readthedocs.io/en/default/_images/area.png)


##3D面积图

创建3D面积图

```python
from openpyxl import Workbook
from openpyxl.chart import (
    AreaChart3D,
    Reference,
    Series,
)

wb = Workbook()
ws = wb.active

rows = [
    ['Number', 'Batch 1', 'Batch 2'],
    [2, 30, 40],
    [3, 25, 40],
    [4 ,30, 50],
    [5 ,10, 30],
    [6,  5, 25],
    [7 ,10, 50],
]

for row in rows:
    ws.append(row)      # 绘制单元格数据

chart = AreaChart3D()           # 创建 3D面积图 对象           
chart.title = "Area Chart"      
chart.style = 13
chart.x_axis.title = 'Test'
chart.y_axis.title = 'Percentage'
chart.legend = None

cats = Reference(ws, min_col=1, min_row=1, max_row=7)       # 关联分类数据的值
data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=7)    #关联构图数据的值
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

ws.add_chart(chart, "A10")

wb.save("area3D.xlsx")
```


这样就创建了一个简单的 3D面积图。 面积图的 Z轴 可以用来作为数据说明。

!["Sample 3D area chart with a series axis"](http://docs.uyinn.com/openpyxl.readthedocs.io/en/default/_images/area3D.png)

----

^[ <-Previous ]( ./introduction.md )  |  [ Next-> ]( ./bar.md )|
