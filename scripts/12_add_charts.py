from openpyxl import load_workbook
from openpyxl.chart import LineChart, BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList

OUT = "/home/claude/fpa-project/model/FPA_Controllership_Suite.xlsx"
wb = load_workbook(OUT)
ws = wb["Executive_Dashboard"]

# Revenue + EBITDA margin trend (combo-ish: two separate charts stacked for clarity)
r_rev = 11  # month labels row
chart1 = LineChart()
chart1.title = "Monthly Revenue Trend (Consolidated, USD)"
chart1.y_axis.title = "Revenue ($)"
chart1.x_axis.title = "Month"
chart1.height = 8
chart1.width = 18
data = Reference(ws, min_col=2, max_col=13, min_row=r_rev + 1, max_row=r_rev + 1)
cats = Reference(ws, min_col=2, max_col=13, min_row=r_rev, max_row=r_rev)
chart1.add_data(data, titles_from_data=False, from_rows=True)
chart1.set_categories(cats)
chart1.series[0].tx = None
ws.add_chart(chart1, "B22")

chart2 = LineChart()
chart2.title = "EBITDA Margin Trend (Consolidated, %)"
chart2.y_axis.title = "EBITDA Margin"
chart2.y_axis.numFmt = '0%'
chart2.height = 8
chart2.width = 18
data2 = Reference(ws, min_col=2, max_col=13, min_row=r_rev + 2, max_row=r_rev + 2)
chart2.add_data(data2, titles_from_data=False, from_rows=True)
chart2.set_categories(cats)
ws.add_chart(chart2, "K22")

# Budget vs Actual bridge (bar)
r_bva = 15
chart3 = BarChart()
chart3.type = "col"
chart3.title = "FY Revenue: Actual vs Budget (Consolidated)"
chart3.y_axis.title = "USD"
chart3.height = 8
chart3.width = 12
data3 = Reference(ws, min_col=2, max_col=2, min_row=r_bva + 1, max_row=r_bva + 2)
cats3 = Reference(ws, min_col=1, max_col=1, min_row=r_bva + 1, max_row=r_bva + 2)
chart3.add_data(data3, titles_from_data=False)
chart3.set_categories(cats3)
chart3.legend = None
ws.add_chart(chart3, "B38")

# Cost center pie
r_cc = 19
chart4 = PieChart()
chart4.title = "FY OpEx Spend by Cost Center (Consolidated)"
chart4.height = 8
chart4.width = 12
data4 = Reference(ws, min_col=2, max_col=2, min_row=r_cc + 1, max_row=r_cc + 5)
cats4 = Reference(ws, min_col=1, max_col=1, min_row=r_cc + 1, max_row=r_cc + 5)
chart4.add_data(data4, titles_from_data=False)
chart4.set_categories(cats4)
chart4.dataLabels = DataLabelList()
chart4.dataLabels.showPercent = True
ws.add_chart(chart4, "K38")

wb.save(OUT)
print("Charts added to Executive Dashboard.")
