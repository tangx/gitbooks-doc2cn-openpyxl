# 图表
[Charts](https://openpyxl.readthedocs.io/en/default/charts/introduction.html)

----


>**###Warning**
>>
>>Openpyxl currently supports chart creation within a worksheet only. Charts in existing workbooks will be lost.
>>
>>openpyxl 目前只支持在内存中创建图表。操作一个已存在的工作簿时，图表会丢失。
>>

## Chart types
## 图表类型

The following charts are available:
以下图表类型可用：

* [ 面积图  ](./area.md)
    * [ 2D Area Charts ](./area.md#2D面积图)
    * [ 3D Area Charts ](./area.md#3D面积图)
* [Bar and Column Charts](./bar.md)
    * [Vertical, Horizontal and Stacked Bar Charts](bar.md#vertical-horizontal-and-stacked-bar-charts)
    * [3D Bar Charts](./bar.md#3d-bar-charts)
* [Bubble Charts]( ./bubble.md )
* [Line Charts]( ./line.md )
    * [Line Charts]( ./line.md )
    * [3D Line Charts]( ./line.md )
* Scatter Charts
* Pie Charts
    * Pie Charts
    * Projected Pie Charts
    * 3D Pie Charts
* Doughnut Charts
* Radar Charts
* Stock Charts
* Surface charts


## Creating a chart
Charts are composed of at least one series of one or more data points. Series themselves are comprised of references to cell ranges.

```python
>>> from openpyxl import Workbook
>>> wb = Workbook()
>>> ws = wb.active
>>> for i in range(10):
...     ws.append([i])
>>>
>>> from openpyxl.chart import BarChart, Reference, Series
>>> values = Reference(ws, min_col=1, min_row=1, max_col=1, max_row=10)
>>> chart = BarChart()
>>> chart.add_data(values)
>>> ws.add_chart(chart)
>>> wb.save("SampleChart.xlsx")
```

## Working with axes
* Axis Limits and Scale
    * Minima and Maxima
    * Logarithmic Scaling
    * Axis Orientation
* Adding a second axis


## Change the chart layout
    * Changing the layout of plot area and legend
        * Chart layout
            * Size and position
            * Mode
            * Target
        * Legend layout


## Styling charts
Adding Patterns


## Advanced charts
Charts can be combined to create new charts:
* Gauge Charts


## Using chartsheets
Charts can be added to special worksheets called chartsheets:
* Chartsheets