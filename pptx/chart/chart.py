# encoding: utf-8

"""Chart-related objects such as Chart and ChartTitle."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from collections import Sequence

import random

from pptx.chart.axis import CategoryAxis, DateAxis, ValueAxis
from pptx.chart.legend import Legend
from pptx.chart.plot import PlotFactory, PlotTypeInspector
from pptx.chart.series import SeriesCollection
from pptx.chart.xmlwriter import SeriesXmlRewriterFactory
from pptx.dml.chtfmt import ChartFormat
from pptx.enum.chart import XL_CHART_TYPE
from pptx.shared import ElementProxy, PartElementProxy
from pptx.text.text import Font, TextFrame
from pptx.util import lazyproperty

from .xmlwriter import ChartXmlWriter


class _Chart(PartElementProxy):
    """A chart object."""

    def __init__(self, chartSpace, chart_part):
        super(_Chart, self).__init__(chartSpace, chart_part)
        self._chartSpace = chartSpace

    @property
    def category_axis(self):
        """
        The category axis of this chart. In the case of an XY or Bubble
        chart, this is the X axis. Raises |ValueError| if no category
        axis is defined (as is the case for a pie chart, for example).
        """

        def get_first_non_deleted(seq):
            for el in seq:
                if el.delete_.val == 1:
                    continue

                return el

            return None

        catAx_lst = self._chartSpace.catAx_lst
        if catAx_lst:
            return CategoryAxis(get_first_non_deleted(catAx_lst))

        dateAx_lst = self._chartSpace.dateAx_lst
        if dateAx_lst:
            return DateAxis(get_first_non_deleted(dateAx_lst))

        valAx_lst = self._chartSpace.valAx_lst
        if valAx_lst:
            return ValueAxis(get_first_non_deleted(valAx_lst))

        raise ValueError('chart has no category axis')

    @property
    def chart_style(self):
        """
        Read/write integer index of chart style used to format this chart.
        Range is from 1 to 48. Value is |None| if no explicit style has been
        assigned, in which case the default chart style is used. Assigning
        |None| causes any explicit setting to be removed. The integer index
        corresponds to the style's position in the chart style gallery in the
        PowerPoint UI.
        """
        style = self._chartSpace.style
        if style is None:
            return None
        return style.val

    @chart_style.setter
    def chart_style(self, value):
        self._chartSpace._remove_style()
        if value is None:
            return
        self._chartSpace._add_style(val=value)

    @property
    def chart_title(self):
        """A |ChartTitle| object providing access to title properties.

        Calling this property is destructive in the sense it adds a chart
        title element (`c:title`) to the chart XML if one is not already
        present. Use :attr:`has_title` to test for presence of a chart title
        non-destructively.
        """
        return ChartTitle(self._element.get_or_add_title())

    @property
    def chart_type(self):
        """
        Read-only :ref:`XlChartType` enumeration value specifying the type of
        this chart. If the chart has two plots, for example, a line plot
        overlayed on a bar plot, the type reported is for the first
        (back-most) plot.
        """
        first_plot = self.plots[0]
        return PlotTypeInspector.chart_type(first_plot)

    @lazyproperty
    def font(self):
        """Font object controlling text format defaults for this chart."""
        defRPr = (
            self._chartSpace
                .get_or_add_txPr()
                .p_lst[0]
                .get_or_add_pPr()
                .get_or_add_defRPr()
        )
        return Font(defRPr)

    @property
    def has_legend(self):
        """
        Read/write boolean, |True| if the chart has a legend. Assigning
        |True| causes a legend to be added to the chart if it doesn't already
        have one. Assigning False removes any existing legend definition
        along with any existing legend settings.
        """
        return self._chartSpace.chart.has_legend

    @has_legend.setter
    def has_legend(self, value):
        self._chartSpace.chart.has_legend = bool(value)

    @property
    def has_title(self):
        """Read/write boolean, specifying whether this chart has a title.

        Assigning |True| causes a title to be added if not already present.
        Assigning |False| removes any existing title along with its text and
        settings.
        """
        title = self._chartSpace.chart.title
        if title is None:
            return False
        return True

    @has_title.setter
    def has_title(self, value):
        chart = self._chartSpace.chart
        if bool(value) is False:
            chart._remove_title()
            autoTitleDeleted = chart.get_or_add_autoTitleDeleted()
            autoTitleDeleted.val = True
            return
        chart.get_or_add_title()

    @property
    def legend(self):
        """
        A |Legend| object providing access to the properties of the legend
        for this chart.
        """
        legend_elm = self._chartSpace.chart.legend
        if legend_elm is None:
            return None
        return Legend(legend_elm)

    @lazyproperty
    def plots(self):
        """
        The sequence of plots in this chart. A plot, called a *chart group*
        in the Microsoft API, is a distinct sequence of one or more series
        depicted in a particular charting type. For example, a chart having
        a series plotted as a line overlaid on three series plotted as
        columns would have two plots; the first corresponding to the three
        column series and the second to the line series. Plots are sequenced
        in the order drawn, i.e. back-most to front-most. Supports *len()*,
        membership (e.g. ``p in plots``), iteration, slicing, and indexed
        access (e.g. ``plot = plots[i]``).
        """
        plotArea = self._chartSpace.chart.plotArea
        return _Plots(plotArea, self)

    def replace_data(self, chart_data):
        """
        Use the categories and series values in the |ChartData| object
        *chart_data* to replace those in the XML and Excel worksheet for this
        chart.
        """
        rewriter = SeriesXmlRewriterFactory(self.chart_type, chart_data)
        rewriter.replace_series_data(self._chartSpace)
        self._workbook.update_from_xlsx_blob(chart_data.xlsx_blob)

    def value_axis(self, idx=0):
        """
        The |ValueAxis| object providing access to properties of the value
        axis of this chart. Raises |ValueError| if the chart has no value
        axis.
        """
        valAx_lst = self._chartSpace.valAx_lst
        if not valAx_lst:
            raise ValueError('chart has no value axis')

        return ValueAxis(valAx_lst[idx])

    @property
    def _workbook(self):
        """
        The |ChartWorkbook| object providing access to the Excel source data
        for this chart.
        """
        return self.part.chart_workbook


class ChartTitle(ElementProxy):
    """Provides properties for manipulating a chart title."""

    # This shares functionality with AxisTitle, which could be factored out
    # into a base class, perhaps pptx.chart.shared.BaseTitle. I suspect they
    # actually differ in certain fuller behaviors, but at present they're
    # essentially identical.

    __slots__ = ('_title', '_format')

    def __init__(self, title):
        super(ChartTitle, self).__init__(title)
        self._title = title

    @lazyproperty
    def format(self):
        """|ChartFormat| object providing access to line and fill formatting.

        Return the |ChartFormat| object providing shape formatting properties
        for this chart title, such as its line color and fill.
        """
        return ChartFormat(self._title)

    @property
    def has_text_frame(self):
        """Read/write Boolean specifying whether this title has a text frame.

        Return |True| if this chart title has a text frame, and |False|
        otherwise. Assigning |True| causes a text frame to be added if not
        already present. Assigning |False| causes any existing text frame to
        be removed along with its text and formatting.
        """
        if self._title.tx_rich is None:
            return False
        return True

    @has_text_frame.setter
    def has_text_frame(self, value):
        if bool(value) is False:
            self._title._remove_tx()
            return
        self._title.get_or_add_tx_rich()

    @property
    def text_frame(self):
        """|TextFrame| instance for this chart title.

        Return a |TextFrame| instance allowing read/write access to the text
        of this chart title and its text formatting properties. Accessing this
        property is destructive in the sense it adds a text frame if one is
        not present. Use :attr:`has_text_frame` to test for the presence of
        a text frame non-destructively.
        """
        rich = self._title.get_or_add_tx_rich()
        return TextFrame(rich, self)


class _Plots(Sequence):
    """
    The sequence of plots in a chart, such as a bar plot or a line plot. Most
    charts have only a single plot. The concept is necessary when two chart
    types are displayed in a single set of axes, like a bar plot with
    a superimposed line plot.
    """
    def __init__(self, plotArea, chart):
        super(_Plots, self).__init__()
        self._plotArea = plotArea
        self._chart = chart

    def __getitem__(self, index):
        xCharts = self._plotArea.xCharts
        if isinstance(index, slice):
            plots = [PlotFactory(xChart, self._chart) for xChart in xCharts]
            return plots[index]
        else:
            xChart = xCharts[index]
            return PlotFactory(xChart, self._chart)

    def __len__(self):
        return len(self._plotArea.xCharts)


class Chart(object):
    def __init__(self, data=None):
        self._plots = []  # type: List[Plot]
        self._data = data
        self._x_axis_id = random.getrandbits(24)
        self._y_axis_id = random.getrandbits(24)
        self._secondary_x_axis_id = None
        self._secondary_y_axis_id = None
        self._has_axes = True

    def add_plot(self, plot):
        if len(self._plots) > 0 and self._has_axes != plot.has_axes:
            raise ValueError("can't mix a plot with and without axes")

        if plot.has_axes:
            # attach axes id to the plot
            if plot.secondary_axis and self._secondary_x_axis_id is None:
                self._secondary_x_axis_id = random.getrandbits(24)
                self._secondary_y_axis_id = random.getrandbits(24)

            plot.y_axis_id = self._secondary_y_axis_id if plot.secondary_axis else self._y_axis_id
            plot.x_axis_id = self._secondary_x_axis_id if plot.secondary_axis else self._x_axis_id
        else:
            self._has_axes = False

        self._plots.append(plot)

    @property
    def plots(self):
        return self._plots

    @property
    def data(self):
        return self._data

    @data.setter
    def data(self, values):
        self._data = values

    @property
    def y_axis_id(self):
        return self._y_axis_id

    @property
    def secondary_x_axis_id(self):
        return self._secondary_x_axis_id

    @property
    def secondary_y_axis_id(self):
        return self._secondary_y_axis_id

    @property
    def x_axis_id(self):
        return self._x_axis_id

    @property
    def has_axes(self):
        return self._has_axes

    def xml_bytes(self):
        """
        Return a blob containing the XML for a chart of *chart_type*
        containing the series in this chart data object, as bytes suitable
        for writing directly to a file.
        """
        return self._xml().encode('utf-8')

    def _xml(self):
        """
        Return (as unicode text) the XML for a chart of *chart_type*
        populated with the values in this chart data object. The XML is
        a complete XML document, including an XML declaration specifying
        UTF-8 encoding.
        """
        return ChartXmlWriter(self).xml

    @property
    def xlsx_blob(self):
        return self.data.xlsx_blob


class Plot(object):
    def __init__(self, chart_type, series_seq, secondary_axis=False):
        self._chart_type = chart_type
        self._series_seq = series_seq
        self._secondary_axis = secondary_axis
        self._x_axis_id = None
        self._y_axis_id = None

    @property
    def chart_type(self):
        return self._chart_type

    @property
    def series_seq(self):
        return self._series_seq

    @property
    def has_axes(self):
        XL_CT = XL_CHART_TYPE
        return {
            XL_CT.AREA: True,
            XL_CT.AREA_STACKED: True,
            XL_CT.AREA_STACKED_100: True,
            XL_CT.BAR_CLUSTERED: True,
            XL_CT.BAR_STACKED: True,
            XL_CT.BAR_STACKED_100: True,
            XL_CT.BUBBLE: True,
            XL_CT.BUBBLE_THREE_D_EFFECT: True,
            XL_CT.COLUMN_CLUSTERED: True,
            XL_CT.COLUMN_STACKED: True,
            XL_CT.COLUMN_STACKED_100: True,
            XL_CT.DOUGHNUT: False,
            XL_CT.DOUGHNUT_EXPLODED: False,
            XL_CT.LINE: True,
            XL_CT.LINE_MARKERS: True,
            XL_CT.LINE_MARKERS_STACKED: True,
            XL_CT.LINE_MARKERS_STACKED_100: True,
            XL_CT.LINE_STACKED: True,
            XL_CT.LINE_STACKED_100: True,
            XL_CT.PIE: False,
            XL_CT.PIE_EXPLODED: False,
            XL_CT.RADAR: False,
            XL_CT.RADAR_FILLED: False,
            XL_CT.RADAR_MARKERS: False,
            XL_CT.XY_SCATTER: True,
            XL_CT.XY_SCATTER_LINES: True,
            XL_CT.XY_SCATTER_LINES_NO_MARKERS: True,
            XL_CT.XY_SCATTER_SMOOTH: True,
            XL_CT.XY_SCATTER_SMOOTH_NO_MARKERS: True,
        }[self._chart_type]

    @property
    def secondary_axis(self):
        return self._secondary_axis

    @property
    def x_axis_id(self):
        return self._x_axis_id

    @x_axis_id.setter
    def x_axis_id(self, value):
        self._x_axis_id = value

    @property
    def y_axis_id(self):
        return self._y_axis_id

    @y_axis_id.setter
    def y_axis_id(self, value):
        self._y_axis_id = value
