# 气泡图
[ Bubble Charts ](https://openpyxl.readthedocs.io/en/default/charts/bubble.html)

----

气泡图与散列图类似，气泡图还使用气泡的大小作为第三维来描述数据的值。图标可以包含多个数据组。

```python

"""

简单气泡图
"""

from openpyxl import Workbook
from openpyxl.chart import Series, Reference, BubbleChart

wb = Workbook()
ws = wb.active

rows = [
    ("Number of Products", "Sales in USD", "Market share"),
    (14, 12200, 15),
    (20, 60000, 33),
    (18, 24400, 10),
    (22, 32000, 42),
    (),
    (12, 8200, 18),
    (15, 50000, 30),
    (19, 22400, 15),
    (25, 25000, 50),
]

for row in rows:
    ws.append(row)

chart = BubbleChart()       # 创建一个 气泡图
chart.style = 18 # use a preset style

# add the first series of data
# 添加第一个数据组
xvalues = Reference(ws, min_col=1, min_row=2, max_row=5)
yvalues = Reference(ws, min_col=2, min_row=2, max_row=5)
size = Reference(ws, min_col=3, min_row=2, max_row=5)
series = Series(values=yvalues, xvalues=xvalues, zvalues=size, title="2013")
chart.series.append(series)     ###  图形数据(chart.series)是一个list。将多个数据组追加进去。

# add the second
#  添加第二个数据组
xvalues = Reference(ws, min_col=1, min_row=7, max_row=10)
yvalues = Reference(ws, min_col=2, min_row=7, max_row=10)
size = Reference(ws, min_col=3, min_row=7, max_row=10)
series = Series(values=yvalues, xvalues=xvalues, zvalues=size, title="2014")
chart.series.append(series)     ###  图形数据(chart.series)是一个list。追加第二个

# place the chart starting in cell E1
# 气泡图生成位置
ws.add_chart(chart, "E1")
wb.save("bubble.xlsx")
```

> **注意**
>
> 本例中，并没有使用第三列数据`Market share`。

以上操作将生成一个两组数据的气泡图，效果如如下：

!["Sample bubble chart"]( https://openpyxl.readthedocs.io/en/default/_images/bubble.png )


----

^[ <-Previous ]( ./bar.md )  |  [ Next-> ]( ./line.md )|