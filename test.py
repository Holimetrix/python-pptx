from pptx import Presentation
from pptx.chart.chart import Plot, Chart
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

# create presentation with 1 slide ------
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])

# define chart data ---------------------
chart_data = CategoryChartData()
chart_data.categories = ['East', 'West', 'Midwest']
chart_data.add_series('Series 1', (19.2, 21.4, 16.7))

plot = Plot(XL_CHART_TYPE.COLUMN_CLUSTERED, chart_data)

chart_data = CategoryChartData()
chart_data.categories = ['East', 'West', 'Midwest']
chart_data.add_series('Series 1', (20, 4, 5))

line = Plot(XL_CHART_TYPE.LINE, chart_data)

chart = Chart()
chart.add_plot(plot)
# chart.add_plot(line)

# add chart to slide --------------------
x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4.5)
slide.shapes.add_chart(
    x, y, cx, cy, chart
)

prs.save('chart-01.pptx')
