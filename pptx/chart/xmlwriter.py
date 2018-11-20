# encoding: utf-8

"""
Composers for default chart XML for various chart types.
"""

from __future__ import absolute_import, print_function, unicode_literals

import collections

from copy import deepcopy
from xml.sax.saxutils import escape

from ..compat import to_unicode
from ..enum.chart import XL_CHART_TYPE
from ..oxml import parse_xml
from ..oxml.ns import nsdecls


class AxesXmlWriter(object):
    def __init__(self, chart):
        self._chart = chart

    @property
    def _x_axis_pos(self):
        return 'b'

    def _y_axis_pos(self, secondary=False):
        return 'l' if not secondary else 'r'

    def _x_axis_xml(self, axis_id, cross_ax_id, secondary=False):

        from pptx.chart.data import CategoryChartData

        hidden = '1' if secondary else '0'
        major_tick_mark = 'out' if secondary else 'out'

        if isinstance(self._chart.data, CategoryChartData):
            categories = self._chart.data.categories

            if categories.are_dates:
                xml = (
                    '      <c:dateAx>\n'
                    '        <c:axId val="{axis_id}"/>\n'
                    '        <c:scaling>\n'
                    '          <c:orientation val="minMax"/>\n'
                    '        </c:scaling>\n'
                    '        <c:delete val="{hidden}"/>\n'
                    '        <c:axPos val="{cat_ax_pos}"/>\n'
                    '        <c:numFmt formatCode="{nf}" sourceLinked="1"/>\n'
                    '        <c:majorTickMark val="{major_tick_mark}"/>\n'
                    '        <c:minorTickMark val="none"/>\n'
                    '        <c:tickLblPos val="nextTo"/>\n'
                    '        <c:crossAx val="{cross_ax_id}"/>\n'
                    #'        <c:crosses val="autoZero"/>\n'
                    '        <c:auto val="1"/>\n'
                    '        <c:lblOffset val="100"/>\n'
                    '        <c:baseTimeUnit val="days"/>\n'
                    #'        <c:majorUnit val="1"/>\n'
                    #'        <c:minorUnit val="1"/>\n'
                    '      </c:dateAx>\n'
                ).format(**{
                    'axis_id': axis_id,
                    'hidden': hidden,
                    'major_tick_mark': major_tick_mark,
                    'cross_ax_id': cross_ax_id,
                    'cat_ax_pos': self._x_axis_pos,
                    'nf': categories.number_format,
                })
            else:
                xml = (
                    '      <c:catAx>\n'
                    '        <c:axId val="{axis_id}"/>\n'
                    '        <c:scaling>\n'
                    '          <c:orientation val="minMax"/>\n'
                    '        </c:scaling>\n'
                    '        <c:delete val="0"/>\n'
                    '        <c:axPos val="{cat_ax_pos}"/>\n'
                    '        <c:majorTickMark val="{major_tick_mark}"/>\n'
                    '        <c:minorTickMark val="none"/>\n'
                    '        <c:tickLblPos val="nextTo"/>\n'
                    '        <c:crossAx val="{cross_ax_id}"/>\n'
                    #'        <c:crosses val="autoZero"/>\n'
                    '        <c:auto val="1"/>\n'
                    '        <c:lblAlgn val="ctr"/>\n'
                    '        <c:lblOffset val="100"/>\n'
                    '        <c:noMultiLvlLbl val="0"/>\n'
                    '      </c:catAx>\n'
                ).format(**{
                    'axis_id': axis_id,
                    'major_tick_mark': major_tick_mark,
                    'cross_ax_id': cross_ax_id,
                    'cat_ax_pos': self._x_axis_pos,
                })
        else:
            raise NotImplementedError("unknown chart_data type")

        return xml

    def _y_axis_xml(self, axis_id, cross_ax_id, secondary=False):

        crosses = '        <c:crosses val="autoZero"/>\n' if not secondary else '        <c:crosses val="max"/>\n'
        major_grid_lines = '' if secondary else '        <c:majorGridlines/>\n'
        major_tick_mark = 'none' if secondary else 'none'

        # Y axis
        xml = (
            '      <c:valAx>\n'
            '        <c:axId val="{axis_id}"/>\n'
            '        <c:scaling>\n'
            '          <c:orientation val="minMax"/>\n'
            '        </c:scaling>\n'
            '        <c:delete val="0"/>\n'
            '        <c:axPos val="{val_ax_pos}"/>\n'
            '{major_grid_lines}'
            '        <c:numFmt formatCode="General" sourceLinked="1"/>'
            '        <c:majorTickMark val="{major_tick_mark}"/>\n'
            '        <c:minorTickMark val="none"/>\n'
            '        <c:tickLblPos val="nextTo"/>\n'
            '        <c:crossAx val="{cross_ax_id}"/>\n'
            '{crosses}'
            '        <c:crossBetween val="between"/>\n'
            '      </c:valAx>\n'
        ).format(**{
            'val_ax_pos': self._y_axis_pos(secondary),
            'major_grid_lines': major_grid_lines,
            'axis_id': axis_id,
            'cross_ax_id': cross_ax_id,
            'crosses': crosses,
            'major_tick_mark': major_tick_mark
        })

        return xml

    @property
    def xml(self):
        axes_xml = ''

        axes_xml += self._y_axis_xml(self._chart.y_axis_id, self._chart.x_axis_id)
        axes_xml += self._x_axis_xml(self._chart.x_axis_id, self._chart.y_axis_id)

        if self._chart.secondary_y_axis_id is not None:
            axes_xml += self._y_axis_xml(self._chart.secondary_y_axis_id, self._chart.secondary_x_axis_id,
                                         secondary=True)
            axes_xml += self._x_axis_xml(self._chart.secondary_x_axis_id, self._chart.secondary_y_axis_id,
                                         secondary=True)

        return axes_xml


class ChartXmlWriter(object):
    """
    Generates XML text (unicode) for a default chart, like the one added by
    PowerPoint when you click the *Add Column Chart* button on the ribbon.
    Differentiated XML for different chart types is provided by subclasses.
    """

    def __init__(self, chart):
        self._chart = chart
        self._date_1904 = False
        self._rounded_corners = False

    @property
    def xml(self):
        """

            '  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.'
            'org/markup-compatibility/2006">\n'
            '    <mc:Choice xmlns:c14="http://schemas.microsoft.com/office/d'
            'rawing/2007/8/2/chart" Requires="c14">\n'
            '      <c14:style val="118"/>\n'
            '    </mc:Choice>\n'
            '    <mc:Fallback>\n'
            '      <c:style val="18"/>\n'
            '    </mc:Fallback>\n'
            '  </mc:AlternateContent>\n'

        :return:
        :rtype:
        """
        return (
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n'
            '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawin'
            'gml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/draw'
            'ingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/off'
            'iceDocument/2006/relationships">\n'
            '  <c:date1904 val="{date_1904_xml}"/>\n'
            '  <c:roundedCorners val="{rounded_corners_xml}"/>\n'
            '  <c:chart>\n'
            '    <c:autoTitleDeleted val="1"/>\n'
            '    <c:plotArea>\n'
            '      <c:layout/>\n'
            '{plots_xml}'
            '{axes_xml}'
            '    </c:plotArea>\n'
            '    <c:legend>\n'
            '      <c:legendPos val="r"/>\n'
            '      <c:layout/>\n'
            '      <c:overlay val="0"/>\n'
            '    </c:legend>\n'
            '    <c:plotVisOnly val="1"/>\n'
            '    <c:dispBlanksAs val="gap"/>\n'
            '    <c:showDLblsOverMax val="0"/>\n'
            '  </c:chart>\n'
            '  <c:spPr>\n'
            '    <a:noFill/>\n'
            '    <a:ln>\n'
            '      <a:noFill/>\n'
            '    </a:ln>\n'
            '    <a:effectLst/>\n'
            '  </c:spPr>\n'
            '  <c:txPr>\n'
            '    <a:bodyPr/>\n'
            '    <a:lstStyle/>\n'
            '    <a:p>\n'
            '      <a:pPr>\n'
            '        <a:defRPr/>\n'
            '      </a:pPr>\n'
            '      <a:endParaRPr lang="en-US"/>\n'
            '    </a:p>\n'
            '  </c:txPr>\n'
            '</c:chartSpace>\n'
        ).format(**{
            'date_1904_xml':       self.date_1904_xml,
            'rounded_corners_xml': self.rounded_corners_xml,
            'plots_xml':           self.plots_xml,
            'axes_xml':            self.axes_xml
        })

    @property
    def date_1904_xml(self):
        return self._bool_to_str(self._date_1904)

    @property
    def rounded_corners_xml(self):
        return self._bool_to_str(self._rounded_corners)

    @property
    def plots_xml(self):
        xml = ''

        for plot in self._chart.plots:
            writer = PlotXmlWriter(plot.chart_type, self._chart.data, plot)
            xml += writer.xml

        return xml

    @property
    def axes_xml(self):
        if self._chart.has_axes:
            return AxesXmlWriter(self._chart).xml
        return ''

    @staticmethod
    def _bool_to_str(value):
        return "1" if value else "0"


def PlotXmlWriter(chart_type, chart_data, plot):
    """
    Factory function returning appropriate XML writer object for
    *chart_type*, loaded with *chart_type* and *chart_data*.
    """
    XL_CT = XL_CHART_TYPE
    try:
        BuilderCls = {
            XL_CT.AREA:                         _AreaChartXmlWriter,
            XL_CT.AREA_STACKED:                 _AreaChartXmlWriter,
            XL_CT.AREA_STACKED_100:             _AreaChartXmlWriter,
            XL_CT.BAR_CLUSTERED:                _BarChartXmlWriter,
            XL_CT.BAR_STACKED:                  _BarChartXmlWriter,
            XL_CT.BAR_STACKED_100:              _BarChartXmlWriter,
            XL_CT.BUBBLE:                       _BubbleChartXmlWriter,
            XL_CT.BUBBLE_THREE_D_EFFECT:        _BubbleChartXmlWriter,
            XL_CT.COLUMN_CLUSTERED:             _BarChartXmlWriter,
            XL_CT.COLUMN_STACKED:               _BarChartXmlWriter,
            XL_CT.COLUMN_STACKED_100:           _BarChartXmlWriter,
            XL_CT.DOUGHNUT:                     _DoughnutChartXmlWriter,
            XL_CT.DOUGHNUT_EXPLODED:            _DoughnutChartXmlWriter,
            XL_CT.LINE:                         _LineChartXmlWriter,
            XL_CT.LINE_MARKERS:                 _LineChartXmlWriter,
            XL_CT.LINE_MARKERS_STACKED:         _LineChartXmlWriter,
            XL_CT.LINE_MARKERS_STACKED_100:     _LineChartXmlWriter,
            XL_CT.LINE_STACKED:                 _LineChartXmlWriter,
            XL_CT.LINE_STACKED_100:             _LineChartXmlWriter,
            XL_CT.PIE:                          _PieChartXmlWriter,
            XL_CT.PIE_EXPLODED:                 _PieChartXmlWriter,
            XL_CT.RADAR:                        _RadarChartXmlWriter,
            XL_CT.RADAR_FILLED:                 _RadarChartXmlWriter,
            XL_CT.RADAR_MARKERS:                _RadarChartXmlWriter,
            XL_CT.XY_SCATTER:                   _XyChartXmlWriter,
            XL_CT.XY_SCATTER_LINES:             _XyChartXmlWriter,
            XL_CT.XY_SCATTER_LINES_NO_MARKERS:  _XyChartXmlWriter,
            XL_CT.XY_SCATTER_SMOOTH:            _XyChartXmlWriter,
            XL_CT.XY_SCATTER_SMOOTH_NO_MARKERS: _XyChartXmlWriter,
        }[chart_type]
    except KeyError:
        raise NotImplementedError(
            'XML writer for chart type %s not yet implemented' % chart_type
        )
    return BuilderCls(chart_type, chart_data, plot)


def SeriesXmlRewriterFactory(chart_type, chart_data):
    """
    Return a |_BaseSeriesXmlRewriter| subclass appropriate to *chart_type*.
    """
    XL_CT = XL_CHART_TYPE

    RewriterCls = {
        # There are 73 distinct chart types, only specify non-category
        # types, others default to _CategorySeriesXmlRewriter. Stock-type
        # charts are multi-plot charts, so no guaratees on how they turn
        # out.
        XL_CT.BUBBLE:                       _BubbleSeriesXmlRewriter,
        XL_CT.BUBBLE_THREE_D_EFFECT:        _BubbleSeriesXmlRewriter,
        XL_CT.XY_SCATTER:                   _XySeriesXmlRewriter,
        XL_CT.XY_SCATTER_LINES:             _XySeriesXmlRewriter,
        XL_CT.XY_SCATTER_LINES_NO_MARKERS:  _XySeriesXmlRewriter,
        XL_CT.XY_SCATTER_SMOOTH:            _XySeriesXmlRewriter,
        XL_CT.XY_SCATTER_SMOOTH_NO_MARKERS: _XySeriesXmlRewriter,
    }.get(chart_type, _CategorySeriesXmlRewriter)

    return RewriterCls(chart_data)


class _BaseChartXmlWriter(object):
    """
    Generates XML text (unicode) for a default chart, like the one added by
    PowerPoint when you click the *Add Column Chart* button on the ribbon.
    Differentiated XML for different chart types is provided by subclasses.
    """
    def __init__(self, chart_type, chart_data, plot):
        super(_BaseChartXmlWriter, self).__init__()
        self._chart_type = chart_type
        self._chart_data = chart_data
        self._plot = plot

    @property
    def xml(self):
        """
        The full XML stream for the chart specified by this chart builder, as
        unicode text. This method must be overridden by each subclass.
        """
        raise NotImplementedError('must be implemented by all subclasses')

    @property
    def _axis_ids(self):
        return (
                '        <c:axId val="{x_axis_id}"/>\n'
                '        <c:axId val="{y_axis_id}"/>\n'
        ).format(**{
            "y_axis_id": self._plot.y_axis_id,
            "x_axis_id": self._plot.x_axis_id
        })

    @staticmethod
    def _bool_to_str(value):
        return "1" if value else "0"


class _BaseSeriesXmlWriter(object):
    """
    Provides shared members for series XML writers.
    """
    def __init__(self, series, date_1904=False):
        super(_BaseSeriesXmlWriter, self).__init__()
        self._series = series
        self._date_1904 = date_1904

    @property
    def name(self):
        """
        The XML-escaped name for this series.
        """
        return escape(self._series.name)

    def numRef_xml(self, wksht_ref, number_format, values):
        """
        Return the ``<c:numRef>`` element specified by the parameters as
        unicode text.
        """
        pt_xml = self.pt_xml(values)
        return (
            '            <c:numRef>\n'
            '              <c:f>{wksht_ref}</c:f>\n'
            '              <c:numCache>\n'
            '                <c:formatCode>{number_format}</c:formatCode>\n'
            '{pt_xml}'
            '              </c:numCache>\n'
            '            </c:numRef>\n'
        ).format(**{
            'wksht_ref':     wksht_ref,
            'number_format': number_format,
            'pt_xml':        pt_xml,
        })

    def pt_xml(self, values):
        """
        Return the ``<c:ptCount>`` and sequence of ``<c:pt>`` elements
        corresponding to *values* as a single unicode text string.
        `c:ptCount` refers to the number of `c:pt` elements in this sequence.
        The `idx` attribute value for `c:pt` elements locates the data point
        in the overall data point sequence of the chart and is started at
        *offset*.
        """
        xml = (
            '                <c:ptCount val="{pt_count}"/>\n'
        ).format(
            pt_count=len(values)
        )

        pt_tmpl = (
            '                <c:pt idx="{idx}">\n'
            '                  <c:v>{value}</c:v>\n'
            '                </c:pt>\n'
        )
        for idx, value in enumerate(values):
            if value is None:
                continue
            xml += pt_tmpl.format(idx=idx, value=value)

        return xml

    @property
    def tx(self):
        """
        Return a ``<c:tx>`` oxml element for this series, containing the
        series name.
        """
        xml = self._tx_tmpl.format(**{
            'wksht_ref':   self._series.name_ref,
            'series_name': self.name,
            'nsdecls':     ' %s' % nsdecls('c'),
        })
        return parse_xml(xml)

    @property
    def tx_xml(self):
        """
        Return the ``<c:tx>`` (tx is short for 'text') element for this
        series as unicode text. This element contains the series name.
        """
        return self._tx_tmpl.format(**{
            'wksht_ref':   self._series.name_ref,
            'series_name': self.name,
            'nsdecls':     '',
        })

    @property
    def _tx_tmpl(self):
        """
        The string formatting template for the ``<c:tx>`` element for this
        series, containing the series title and spreadsheet range reference.
        """
        return (
            '          <c:tx{nsdecls}>\n'
            '            <c:strRef>\n'
            '              <c:f>{wksht_ref}</c:f>\n'
            '              <c:strCache>\n'
            '                <c:ptCount val="1"/>\n'
            '                <c:pt idx="0">\n'
            '                  <c:v>{series_name}</c:v>\n'
            '                </c:pt>\n'
            '              </c:strCache>\n'
            '            </c:strRef>\n'
            '          </c:tx>\n'
        )


class _BaseSeriesXmlRewriter(object):
    """
    Base class for series XML rewriters.
    """
    def __init__(self, chart_data):
        super(_BaseSeriesXmlRewriter, self).__init__()
        self._chart_data = chart_data

    def replace_series_data(self, chartSpace):
        """
        Rewrite the series data under *chartSpace* using the chart data
        contents. All series-level formatting is left undisturbed. If
        the chart data contains fewer series than *chartSpace*, the extra
        series in *chartSpace* are deleted. If *chart_data* contains more
        series than the *chartSpace* element, new series are added to the
        last plot in the chart and series formatting is "cloned" from the
        last series in that plot.
        """
        plotArea, date_1904 = chartSpace.plotArea, chartSpace.date_1904
        chart_data = self._chart_data
        self._adjust_ser_count(plotArea, len(chart_data))
        for ser, series_data in zip(plotArea.sers, chart_data):
            self._rewrite_ser_data(ser, series_data, date_1904)

    def _add_cloned_sers(self, plotArea, count):
        """
        Add `c:ser` elements to the last xChart element in *plotArea*, cloned
        from the last `c:ser` child of that last xChart.
        """
        def clone_ser(ser):
            new_ser = deepcopy(ser)
            new_ser.idx.val = plotArea.next_idx
            new_ser.order.val = plotArea.next_order
            ser.addnext(new_ser)
            return new_ser

        last_ser = plotArea.last_ser
        for _ in range(count):
            last_ser = clone_ser(last_ser)

    def _adjust_ser_count(self, plotArea, new_ser_count):
        """
        Adjust the number of c:ser elements in *plotArea* to *new_ser_count*.
        Excess c:ser elements are deleted from the end, along with any xChart
        elements that are left empty as a result. Series elements are
        considered in xChart + series order. Any new c:ser elements required
        are added to the last xChart element and cloned from the last c:ser
        element in that xChart.
        """
        ser_count_diff = new_ser_count - len(plotArea.sers)
        if ser_count_diff > 0:
            self._add_cloned_sers(plotArea, ser_count_diff)
        elif ser_count_diff < 0:
            self._trim_ser_count_by(plotArea, abs(ser_count_diff))

    def _rewrite_ser_data(self, ser, series_data, date_1904):
        """
        Rewrite selected child elements of *ser* based on the values in
        *series_data*.
        """
        raise NotImplementedError('must be implemented by each subclass')

    def _trim_ser_count_by(self, plotArea, count):
        """
        Remove the last *count* ser elements from *plotArea*. Any xChart
        elements having no ser child elements after trimming are also
        removed.
        """
        extra_sers = plotArea.sers[-count:]
        for ser in extra_sers:
            parent = ser.getparent()
            parent.remove(ser)
        extra_xCharts = [
            xChart for xChart in plotArea.iter_xCharts()
            if len(xChart.sers) == 0
        ]
        for xChart in extra_xCharts:
            parent = xChart.getparent()
            parent.remove(xChart)


class _AreaChartXmlWriter(_BaseChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:areaChart>`` element.
    """
    @property
    def xml(self):
        return (
            '      <c:areaChart>\n'
            '{grouping_xml}'
            '        <c:varyColors val="0"/>\n'
            '{ser_xml}'
            '        <c:dLbls>\n'
            '          <c:showLegendKey val="0"/>\n'
            '          <c:showVal val="0"/>\n'
            '          <c:showCatName val="0"/>\n'
            '          <c:showSerName val="0"/>\n'
            '          <c:showPercent val="0"/>\n'
            '          <c:showBubbleSize val="0"/>\n'
            '        </c:dLbls>\n'
            '{axis_ids}'
            '      </c:areaChart>\n'
        ).format(**{
            'grouping_xml': self._grouping_xml,
            'ser_xml':      self._ser_xml,
            'axis_ids':     self._axis_ids,
        })

    @property
    def _grouping_xml(self):
        val = {
            XL_CHART_TYPE.AREA:             'standard',
            XL_CHART_TYPE.AREA_STACKED:     'stacked',
            XL_CHART_TYPE.AREA_STACKED_100: 'percentStacked',
        }[self._chart_type]
        return '        <c:grouping val="%s"/>\n' % val

    @property
    def _ser_xml(self):
        xml = ''
        for series in self._chart_data:
            xml_writer = _CategorySeriesXmlWriter(series)
            xml += (
                '        <c:ser>\n'
                '          <c:idx val="{ser_idx}"/>\n'
                '          <c:order val="{ser_order}"/>\n'
                '{tx_xml}'
                '{cat_xml}'
                '{val_xml}'
                '        </c:ser>\n'
            ).format(**{
                'ser_idx':    series.index,
                'ser_order':  series.index,
                'tx_xml':     xml_writer.tx_xml,
                'cat_xml':    xml_writer.cat_xml,
                'val_xml':    xml_writer.val_xml,
            })
        return xml


class _BarChartXmlWriter(_BaseChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:barChart>`` element.
    """
    @property
    def xml(self):
        return (
            '      <c:barChart>\n'
            '{barDir_xml}'
            '{grouping_xml}'
            '{ser_xml}'
            '        <c:dLbls>\n'
            '          <c:showLegendKey val="0"/>\n'
            '          <c:showVal val="0"/>\n'
            '          <c:showCatName val="0"/>\n'
            '          <c:showSerName val="0"/>\n'
            '          <c:showPercent val="0"/>\n'
            '          <c:showBubbleSize val="0"/>\n'
            '        </c:dLbls>\n'
            '{overlap_xml}'
            '{axis_ids}'
            '      </c:barChart>\n'
        ).format(**{
            'barDir_xml':   self._barDir_xml,
            'grouping_xml': self._grouping_xml,
            'ser_xml':      self._ser_xml,
            'overlap_xml':  self._overlap_xml,
            'axis_ids':     self._axis_ids
        })

    @property
    def _barDir_xml(self):
        XL = XL_CHART_TYPE
        bar_types = (
            XL.BAR_CLUSTERED, XL.BAR_STACKED, XL.BAR_STACKED_100
        )
        col_types = (
            XL.COLUMN_CLUSTERED, XL.COLUMN_STACKED, XL.COLUMN_STACKED_100
        )
        if self._chart_type in bar_types:
            return '        <c:barDir val="bar"/>\n'
        elif self._chart_type in col_types:
            return '        <c:barDir val="col"/>\n'
        raise NotImplementedError(
            'no _barDir_xml() for chart type %s' % self._chart_type
        )

    @property
    def _x_axis_pos(self):
        return {
            XL_CHART_TYPE.BAR_CLUSTERED:      'l',
            XL_CHART_TYPE.BAR_STACKED:        'l',
            XL_CHART_TYPE.BAR_STACKED_100:    'l',
            XL_CHART_TYPE.COLUMN_CLUSTERED:   'b',
            XL_CHART_TYPE.COLUMN_STACKED:     'b',
            XL_CHART_TYPE.COLUMN_STACKED_100: 'b',
        }[self._chart_type]

    @property
    def _grouping_xml(self):
        XL = XL_CHART_TYPE
        clustered_types = (XL.BAR_CLUSTERED, XL.COLUMN_CLUSTERED)
        stacked_types = (XL.BAR_STACKED, XL.COLUMN_STACKED)
        percentStacked_types = (XL.BAR_STACKED_100, XL.COLUMN_STACKED_100)
        if self._chart_type in clustered_types:
            return '        <c:grouping val="clustered"/>\n'
        elif self._chart_type in stacked_types:
            return '        <c:grouping val="stacked"/>\n'
        elif self._chart_type in percentStacked_types:
            return '        <c:grouping val="percentStacked"/>\n'
        raise NotImplementedError(
            'no _grouping_xml() for chart type %s' % self._chart_type
        )

    @property
    def _overlap_xml(self):
        XL = XL_CHART_TYPE
        percentStacked_types = (
            XL.BAR_STACKED, XL.BAR_STACKED_100, XL.COLUMN_STACKED,
            XL.COLUMN_STACKED_100,
        )
        if self._chart_type in percentStacked_types:
            return '        <c:overlap val="100"/>\n'
        return ''

    @property
    def _ser_xml(self):
        xml = ''
        for series in self._plot.series_seq:
            xml_writer = _CategorySeriesXmlWriter(series)
            xml += (
                '        <c:ser>\n'
                '          <c:idx val="{ser_idx}"/>\n'
                '          <c:order val="{ser_order}"/>\n'
                '{tx_xml}'
                '{cat_xml}'
                '{val_xml}'
                '        </c:ser>\n'
            ).format(**{
                'ser_idx':    series.index,
                'ser_order':  series.index,
                'tx_xml':     xml_writer.tx_xml,
                'cat_xml':    xml_writer.cat_xml,
                'val_xml':    xml_writer.val_xml,
            })
        return xml

    @property
    def _y_ax_pos(self):
        return {
            XL_CHART_TYPE.BAR_CLUSTERED:      'b',
            XL_CHART_TYPE.BAR_STACKED:        'b',
            XL_CHART_TYPE.BAR_STACKED_100:    'b',
            XL_CHART_TYPE.COLUMN_CLUSTERED:   'l',
            XL_CHART_TYPE.COLUMN_STACKED:     'l',
            XL_CHART_TYPE.COLUMN_STACKED_100: 'l',
        }[self._chart_type]

class _DoughnutChartXmlWriter(_BaseChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:doughnutChart>``
    element.
    """
    @property
    def xml(self):
        return (
            '      <c:doughnutChart>\n'
            '        <c:varyColors val="1"/>\n'
            '{ser_xml}'
            '        <c:dLbls>\n'
            '          <c:showLegendKey val="0"/>\n'
            '          <c:showVal val="0"/>\n'
            '          <c:showCatName val="0"/>\n'
            '          <c:showSerName val="0"/>\n'
            '          <c:showPercent val="0"/>\n'
            '          <c:showBubbleSize val="0"/>\n'
            '          <c:showLeaderLines val="1"/>\n'
            '        </c:dLbls>\n'
            '        <c:firstSliceAng val="0"/>\n'
            '        <c:holeSize val="50"/>\n'
            '      </c:doughnutChart>\n'
        ).format(**{
            'ser_xml':      self._ser_xml,
        })

    @property
    def _explosion_xml(self):
        if self._chart_type == XL_CHART_TYPE.DOUGHNUT_EXPLODED:
            return '          <c:explosion val="25"/>\n'
        return ''

    @property
    def _ser_xml(self):
        xml = ''
        for series in self._chart_data:
            xml_writer = _CategorySeriesXmlWriter(series)
            xml += (
                '        <c:ser>\n'
                '          <c:idx val="{ser_idx}"/>\n'
                '          <c:order val="{ser_order}"/>\n'
                '{tx_xml}'
                '{explosion_xml}'
                '{cat_xml}'
                '{val_xml}'
                '        </c:ser>\n'
            ).format(**{
                'ser_idx':       series.index,
                'ser_order':     series.index,
                'tx_xml':        xml_writer.tx_xml,
                'explosion_xml': self._explosion_xml,
                'cat_xml':       xml_writer.cat_xml,
                'val_xml':       xml_writer.val_xml,
            })
        return xml


class _LineChartXmlWriter(_BaseChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:lineChart>`` element.
    """
    @property
    def xml(self):
        return (
            '      <c:lineChart>\n'
            '{grouping_xml}'
            '        <c:varyColors val="0"/>\n'
            '{ser_xml}'
            '        <c:dLbls>\n'
            '          <c:showLegendKey val="0"/>\n'
            '          <c:showVal val="0"/>\n'
            '          <c:showCatName val="0"/>\n'
            '          <c:showSerName val="0"/>\n'
            '          <c:showPercent val="0"/>\n'
            '          <c:showBubbleSize val="0"/>\n'
            '        </c:dLbls>\n'
            '        <c:marker val="1"/>\n'
            '        <c:smooth val="0"/>\n'
            '{axis_ids}'
            '      </c:lineChart>\n'
        ).format(**{
            'grouping_xml': self._grouping_xml,
            'ser_xml':      self._ser_xml,
            'axis_ids':     self._axis_ids
        })

    @property
    def _grouping_xml(self):
        XL = XL_CHART_TYPE
        standard_types = (XL.LINE, XL.LINE_MARKERS)
        stacked_types = (XL.LINE_STACKED, XL.LINE_MARKERS_STACKED)
        percentStacked_types = (
            XL.LINE_STACKED_100, XL.LINE_MARKERS_STACKED_100
        )
        if self._chart_type in standard_types:
            return '        <c:grouping val="standard"/>\n'
        elif self._chart_type in stacked_types:
            return '        <c:grouping val="stacked"/>\n'
        elif self._chart_type in percentStacked_types:
            return '        <c:grouping val="percentStacked"/>\n'
        raise NotImplementedError(
            'no _grouping_xml() for chart type %s' % self._chart_type
        )

    @property
    def _marker_xml(self):
        XL = XL_CHART_TYPE
        no_marker_types = (XL.LINE, XL.LINE_STACKED, XL.LINE_STACKED_100)
        if self._chart_type in no_marker_types:
            return (
                '          <c:marker>\n'
                '            <c:symbol val="none"/>\n'
                '          </c:marker>\n'
            )
        return ''

    @property
    def _ser_xml(self):
        xml = ''
        for series in self._plot.series_seq:
            xml_writer = _CategorySeriesXmlWriter(series)
            xml += (
                '        <c:ser>\n'
                '          <c:idx val="{ser_idx}"/>\n'
                '          <c:order val="{ser_order}"/>\n'
                '{tx_xml}'
                '{marker_xml}'
                '{cat_xml}'
                '{val_xml}'
                '          <c:smooth val="0"/>\n'
                '        </c:ser>\n'
            ).format(**{
                'ser_idx':    series.index,
                'ser_order':  series.index,
                'tx_xml':     xml_writer.tx_xml,
                'marker_xml': self._marker_xml,
                'cat_xml':    xml_writer.cat_xml,
                'val_xml':    xml_writer.val_xml,
            })
        return xml


class _PieChartXmlWriter(_BaseChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:pieChart>`` element.
    """
    @property
    def xml(self):
        return (
            '      <c:pieChart>\n'
            '        <c:varyColors val="1"/>\n'
            '{ser_xml}'
            '      </c:pieChart>\n'
        ).format(**{
            'ser_xml': self._ser_xml,
        })

    @property
    def _explosion_xml(self):
        if self._chart_type == XL_CHART_TYPE.PIE_EXPLODED:
            return '          <c:explosion val="25"/>\n'
        return ''

    @property
    def _ser_xml(self):
        xml_writer = _CategorySeriesXmlWriter(self._chart_data.series[0])
        xml = (
            '        <c:ser>\n'
            '          <c:idx val="0"/>\n'
            '          <c:order val="0"/>\n'
            '{tx_xml}'
            '{explosion_xml}'
            '{cat_xml}'
            '{val_xml}'
            '        </c:ser>\n'
        ).format(**{
            'tx_xml':        xml_writer.tx_xml,
            'explosion_xml': self._explosion_xml,
            'cat_xml':       xml_writer.cat_xml,
            'val_xml':       xml_writer.val_xml,
        })
        return xml


class _RadarChartXmlWriter(_BaseChartXmlWriter):
    """
    Generates XML for the ``<c:radarChart>`` element.
    """
    @property
    def xml(self):
        return (
            '      <c:radarChart>\n'
            '        <c:radarStyle val="{radar_style}"/>\n'
            '        <c:varyColors val="0"/>\n'
            '{ser_xml}'
            '{axis_ids}'
            '      </c:radarChart>\n'
        ).format(**{
            'radar_style': self._radar_style,
            'ser_xml':     self._ser_xml,
            'axis_ids':    self._axis_ids
        })

    @property
    def _marker_xml(self):
        if self._chart_type == XL_CHART_TYPE.RADAR:
            return (
                '          <c:marker>\n'
                '            <c:symbol val="none"/>\n'
                '          </c:marker>\n'
            )
        return ''

    @property
    def _radar_style(self):
        if self._chart_type == XL_CHART_TYPE.RADAR_FILLED:
            return 'filled'
        return 'marker'

    @property
    def _ser_xml(self):
        xml = ''
        for series in self._chart_data:
            xml_writer = _CategorySeriesXmlWriter(series)
            xml += (
                '        <c:ser>\n'
                '          <c:idx val="{ser_idx}"/>\n'
                '          <c:order val="{ser_order}"/>\n'
                '{tx_xml}'
                '{marker_xml}'
                '{cat_xml}'
                '{val_xml}'
                '          <c:smooth val="0"/>\n'
                '        </c:ser>\n'
            ).format(**{
                'ser_idx':    series.index,
                'ser_order':  series.index,
                'tx_xml':     xml_writer.tx_xml,
                'marker_xml': self._marker_xml,
                'cat_xml':    xml_writer.cat_xml,
                'val_xml':    xml_writer.val_xml,
            })
        return xml


class _XyChartXmlWriter(_BaseChartXmlWriter):
    """
    Generates XML for the ``<c:scatterChart>`` element.
    """
    @property
    def xml(self):
        xml = (
            '      <c:scatterChart>\n'
            '        <c:scatterStyle val="%s"/>\n'
            '        <c:varyColors val="0"/>\n'
            '%s'
            '%s'
            '      </c:scatterChart>\n'
        ) % (self._scatterStyle_val, self._ser_xml, self._axis_ids)
        return xml

    @property
    def _marker_xml(self):
        no_marker_types = (
            XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS,
            XL_CHART_TYPE.XY_SCATTER_SMOOTH_NO_MARKERS,
        )
        if self._chart_type in no_marker_types:
            return (
                '          <c:marker>\n'
                '            <c:symbol val="none"/>\n'
                '          </c:marker>\n'
            )
        return ''

    @property
    def _scatterStyle_val(self):
        smooth_types = (
            XL_CHART_TYPE.XY_SCATTER_SMOOTH,
            XL_CHART_TYPE.XY_SCATTER_SMOOTH_NO_MARKERS,
        )
        if self._chart_type in smooth_types:
            return 'smoothMarker'
        return 'lineMarker'

    @property
    def _ser_xml(self):
        xml = ''
        for series in self._chart_data:
            xml_writer = _XySeriesXmlWriter(series)
            xml += (
                '        <c:ser>\n'
                '          <c:idx val="{ser_idx}"/>\n'
                '          <c:order val="{ser_order}"/>\n'
                '{tx_xml}'
                '{spPr_xml}'
                '{marker_xml}'
                '{xVal_xml}'
                '{yVal_xml}'
                '          <c:smooth val="0"/>\n'
                '        </c:ser>\n'
            ).format(**{
                'ser_idx':    series.index,
                'ser_order':  series.index,
                'tx_xml':     xml_writer.tx_xml,
                'spPr_xml':   self._spPr_xml,
                'marker_xml': self._marker_xml,
                'xVal_xml':   xml_writer.xVal_xml,
                'yVal_xml':   xml_writer.yVal_xml,
            })
        return xml

    @property
    def _spPr_xml(self):
        if self._chart_type == XL_CHART_TYPE.XY_SCATTER:
            return (
                '          <c:spPr>\n'
                '            <a:ln w="47625">\n'
                '              <a:noFill/>\n'
                '            </a:ln>\n'
                '          </c:spPr>\n'
            )
        return ''


class _BubbleChartXmlWriter(_XyChartXmlWriter):
    """
    Provides specialized methods particular to the ``<c:bubbleChart>``
    element.
    """
    @property
    def xml(self):
        xml = (
            '      <c:bubbleChart>\n'
            '        <c:varyColors val="0"/>\n'
            '%s'
            '        <c:dLbls>\n'
            '          <c:showLegendKey val="0"/>\n'
            '          <c:showVal val="0"/>\n'
            '          <c:showCatName val="0"/>\n'
            '          <c:showSerName val="0"/>\n'
            '          <c:showPercent val="0"/>\n'
            '          <c:showBubbleSize val="0"/>\n'
            '        </c:dLbls>\n'
            '        <c:bubbleScale val="100"/>\n'
            '        <c:showNegBubbles val="0"/>\n'
            '%s'
            '      </c:bubbleChart>\n'
        ) % (self._ser_xml, self._axis_ids)
        return xml

    @property
    def _bubble3D_val(self):
        if self._chart_type == XL_CHART_TYPE.BUBBLE_THREE_D_EFFECT:
            return '1'
        return '0'

    @property
    def _ser_xml(self):
        xml = ''
        for series in self._chart_data:
            xml_writer = _BubbleSeriesXmlWriter(series)
            xml += (
                '        <c:ser>\n'
                '          <c:idx val="{ser_idx}"/>\n'
                '          <c:order val="{ser_order}"/>\n'
                '{tx_xml}'
                '          <c:invertIfNegative val="0"/>\n'
                '{xVal_xml}'
                '{yVal_xml}'
                '{bubbleSize_xml}'
                '          <c:bubble3D val="{bubble3D_val}"/>\n'
                '        </c:ser>\n'
            ).format(**{
                'ser_idx':        series.index,
                'ser_order':      series.index,
                'tx_xml':         xml_writer.tx_xml,
                'xVal_xml':       xml_writer.xVal_xml,
                'yVal_xml':       xml_writer.yVal_xml,
                'bubbleSize_xml': xml_writer.bubbleSize_xml,
                'bubble3D_val':   self._bubble3D_val,
            })
        return xml


class _CategorySeriesXmlWriter(_BaseSeriesXmlWriter):
    """
    Generates XML snippets particular to a category chart series.
    """
    @property
    def cat(self):
        """
        Return the ``<c:cat>`` element XML for this series, as an oxml
        element.
        """
        categories = self._series.categories

        if categories.are_numeric:
            return parse_xml(
                self._numRef_cat_tmpl.format(**{
                    'wksht_ref':     self._series.categories_ref,
                    'number_format': categories.number_format,
                    'cat_count':     categories.leaf_count,
                    'cat_pt_xml':    self._cat_num_pt_xml,
                    'nsdecls':       ' %s' % nsdecls('c'),
                })
            )

        if categories.depth == 1:
            return parse_xml(
                self._cat_tmpl.format(**{
                    'wksht_ref':  self._series.categories_ref,
                    'cat_count':  categories.leaf_count,
                    'cat_pt_xml': self._cat_pt_xml,
                    'nsdecls':    ' %s' % nsdecls('c'),
                })
            )

        return parse_xml(
            self._multiLvl_cat_tmpl.format(**{
                'wksht_ref': self._series.categories_ref,
                'cat_count': categories.leaf_count,
                'lvl_xml':   self._lvl_xml(categories),
                'nsdecls':   ' %s' % nsdecls('c'),
            })
        )

    @property
    def cat_xml(self):
        """
        The unicode XML snippet for the ``<c:cat>`` element for this series,
        containing the category labels and spreadsheet reference.
        """
        categories = self._series.categories

        if categories.are_numeric:
            return self._numRef_cat_tmpl.format(**{
                'wksht_ref':     self._series.categories_ref,
                'number_format': categories.number_format,
                'cat_count':     categories.leaf_count,
                'cat_pt_xml':    self._cat_num_pt_xml,
                'nsdecls':       '',
            })

        if categories.depth == 1:
            return self._cat_tmpl.format(**{
                'wksht_ref':  self._series.categories_ref,
                'cat_count':  categories.leaf_count,
                'cat_pt_xml': self._cat_pt_xml,
                'nsdecls':    '',
            })

        return self._multiLvl_cat_tmpl.format(**{
            'wksht_ref': self._series.categories_ref,
            'cat_count': categories.leaf_count,
            'lvl_xml':   self._lvl_xml(categories),
            'nsdecls':   '',
        })

    @property
    def val(self):
        """
        The ``<c:val>`` XML for this series, as an oxml element.
        """
        xml = self._val_tmpl.format(**{
            'nsdecls':       ' %s' % nsdecls('c'),
            'values_ref':    self._series.values_ref,
            'number_format': self._series.number_format,
            'val_count':     len(self._series),
            'val_pt_xml':    self._val_pt_xml,
        })
        return parse_xml(xml)

    @property
    def val_xml(self):
        """
        Return the unicode XML snippet for the ``<c:val>`` element describing
        this series, containing the series values and their spreadsheet range
        reference.
        """
        return self._val_tmpl.format(**{
            'nsdecls':       '',
            'values_ref':    self._series.values_ref,
            'number_format': self._series.number_format,
            'val_count':     len(self._series),
            'val_pt_xml':    self._val_pt_xml,
        })

    @property
    def _cat_num_pt_xml(self):
        """
        The unicode XML snippet for the ``<c:pt>`` elements when category
        labels are numeric (including date type).
        """
        xml = ''
        for idx, category in enumerate(self._series.categories):
            xml += (
                '                <c:pt idx="{cat_idx}">\n'
                '                  <c:v>{cat_lbl_str}</c:v>\n'
                '                </c:pt>\n'
            ).format(**{
                'cat_idx':     idx,
                'cat_lbl_str': category.numeric_str_val(self._date_1904),
            })
        return xml

    @property
    def _cat_pt_xml(self):
        """
        The unicode XML snippet for the ``<c:pt>`` elements containing the
        category names for this series.
        """
        xml = ''
        for idx, category in enumerate(self._series.categories):
            xml += (
                '                <c:pt idx="{cat_idx}">\n'
                '                  <c:v>{cat_label}</c:v>\n'
                '                </c:pt>\n'
            ).format(**{
                'cat_idx':   idx,
                'cat_label': escape(to_unicode(category.label)),
            })
        return xml

    @property
    def _cat_tmpl(self):
        """
        The template for the ``<c:cat>`` element for this series, containing
        the category labels and spreadsheet reference.
        """
        return (
            '          <c:cat{nsdecls}>\n'
            '            <c:strRef>\n'
            '              <c:f>{wksht_ref}</c:f>\n'
            '              <c:strCache>\n'
            '                <c:ptCount val="{cat_count}"/>\n'
            '{cat_pt_xml}'
            '              </c:strCache>\n'
            '            </c:strRef>\n'
            '          </c:cat>\n'
        )

    def _lvl_xml(self, categories):
        """
        The unicode XML snippet for the ``<c:lvl>`` elements containing
        multi-level category names.
        """
        def lvl_pt_xml(level):
            xml = ''
            for idx, name in level:
                xml += (
                    '                  <c:pt idx="%d">\n'
                    '                    <c:v>%s</c:v>\n'
                    '                  </c:pt>\n'
                ) % (idx, escape('%s' % name))
            return xml

        xml = ''
        for level in categories.levels:
            xml += (
                '                <c:lvl>\n'
                '{lvl_pt_xml}'
                '                </c:lvl>\n'
            ).format(**{
                'lvl_pt_xml': lvl_pt_xml(level),
            })
        return xml

    @property
    def _multiLvl_cat_tmpl(self):
        """
        The template for the ``<c:cat>`` element for this series when there
        are multi-level (nested) categories.
        """
        return (
            '          <c:cat{nsdecls}>\n'
            '            <c:multiLvlStrRef>\n'
            '              <c:f>{wksht_ref}</c:f>\n'
            '              <c:multiLvlStrCache>\n'
            '                <c:ptCount val="{cat_count}"/>\n'
            '{lvl_xml}'
            '              </c:multiLvlStrCache>\n'
            '            </c:multiLvlStrRef>\n'
            '          </c:cat>\n'
        )

    @property
    def _numRef_cat_tmpl(self):
        """
        The template for the ``<c:cat>`` element for this series when the
        labels are numeric (or date) values.
        """
        return (
            '          <c:cat{nsdecls}>\n'
            '            <c:numRef>\n'
            '              <c:f>{wksht_ref}</c:f>\n'
            '              <c:numCache>\n'
            '                <c:formatCode>{number_format}</c:formatCode>\n'
            '                <c:ptCount val="{cat_count}"/>\n'
            '{cat_pt_xml}'
            '              </c:numCache>\n'
            '            </c:numRef>\n'
            '          </c:cat>\n'
        )

    @property
    def _val_pt_xml(self):
        """
        The unicode XML snippet containing the ``<c:pt>`` elements containing
        the values for this series.
        """
        xml = ''
        for idx, value in enumerate(self._series.values):
            if value is None:
                continue
            xml += (
                '                <c:pt idx="{val_idx:d}">\n'
                '                  <c:v>{value}</c:v>\n'
                '                </c:pt>\n'
            ).format(**{
                'val_idx': idx,
                'value':   value,
            })
        return xml

    @property
    def _val_tmpl(self):
        """
        The template for the ``<c:val>`` element for this series, containing
        the series values and their spreadsheet range reference.
        """
        return (
            '          <c:val{nsdecls}>\n'
            '            <c:numRef>\n'
            '              <c:f>{values_ref}</c:f>\n'
            '              <c:numCache>\n'
            '                <c:formatCode>{number_format}</c:formatCode>\n'
            '                <c:ptCount val="{val_count}"/>\n'
            '{val_pt_xml}'
            '              </c:numCache>\n'
            '            </c:numRef>\n'
            '          </c:val>\n'
        )


class _XySeriesXmlWriter(_BaseSeriesXmlWriter):
    """
    Generates XML snippets particular to an XY series.
    """
    @property
    def xVal(self):
        """
        Return the ``<c:xVal>`` element for this series as an oxml element.
        This element contains the X values for this series.
        """
        xml = self._xVal_tmpl.format(**{
            'nsdecls':    ' %s' % nsdecls('c'),
            'numRef_xml': self.numRef_xml(
                self._series.x_values_ref, self._series.number_format,
                self._series.x_values
            ),
        })
        return parse_xml(xml)

    @property
    def xVal_xml(self):
        """
        Return the ``<c:xVal>`` element for this series as unicode text. This
        element contains the X values for this series.
        """
        return self._xVal_tmpl.format(**{
            'nsdecls':    '',
            'numRef_xml': self.numRef_xml(
                self._series.x_values_ref, self._series.number_format,
                self._series.x_values
            ),
        })

    @property
    def yVal(self):
        """
        Return the ``<c:yVal>`` element for this series as an oxml element.
        This element contains the Y values for this series.
        """
        xml = self._yVal_tmpl.format(**{
            'nsdecls':    ' %s' % nsdecls('c'),
            'numRef_xml': self.numRef_xml(
                self._series.y_values_ref, self._series.number_format,
                self._series.y_values
            ),
        })
        return parse_xml(xml)

    @property
    def yVal_xml(self):
        """
        Return the ``<c:yVal>`` element for this series as unicode text. This
        element contains the Y values for this series.
        """
        return self._yVal_tmpl.format(**{
            'nsdecls':    '',
            'numRef_xml': self.numRef_xml(
                self._series.y_values_ref, self._series.number_format,
                self._series.y_values
            ),
        })

    @property
    def _xVal_tmpl(self):
        """
        The template for the ``<c:xVal>`` element for this series, containing
        the X values and their spreadsheet range reference.
        """
        return (
            '          <c:xVal{nsdecls}>\n'
            '{numRef_xml}'
            '          </c:xVal>\n'
        )

    @property
    def _yVal_tmpl(self):
        """
        The template for the ``<c:yVal>`` element for this series, containing
        the Y values and their spreadsheet range reference.
        """
        return (
            '          <c:yVal{nsdecls}>\n'
            '{numRef_xml}'
            '          </c:yVal>\n'
        )


class _BubbleSeriesXmlWriter(_XySeriesXmlWriter):
    """
    Generates XML snippets particular to a bubble chart series.
    """
    @property
    def bubbleSize(self):
        """
        Return the ``<c:bubbleSize>`` element for this series as an oxml
        element. This element contains the bubble size values for this
        series.
        """
        xml = self._bubbleSize_tmpl.format(**{
            'nsdecls':    ' %s' % nsdecls('c'),
            'numRef_xml': self.numRef_xml(
                self._series.bubble_sizes_ref, self._series.number_format,
                self._series.bubble_sizes
            ),
        })
        return parse_xml(xml)

    @property
    def bubbleSize_xml(self):
        """
        Return the ``<c:bubbleSize>`` element for this series as unicode
        text. This element contains the bubble size values for all the
        data points in the chart.
        """
        return self._bubbleSize_tmpl.format(**{
            'nsdecls':    '',
            'numRef_xml': self.numRef_xml(
                self._series.bubble_sizes_ref, self._series.number_format,
                self._series.bubble_sizes
            ),
        })

    @property
    def _bubbleSize_tmpl(self):
        """
        The template for the ``<c:bubbleSize>`` element for this series,
        containing the bubble size values and their spreadsheet range
        reference.
        """
        return (
            '          <c:bubbleSize{nsdecls}>\n'
            '{numRef_xml}'
            '          </c:bubbleSize>\n'
        )


class _BubbleSeriesXmlRewriter(_BaseSeriesXmlRewriter):
    """
    A series rewriter suitable for bubble charts.
    """
    def _rewrite_ser_data(self, ser, series_data, date_1904):
        """
        Rewrite the ``<c:tx>``, ``<c:cat>`` and ``<c:val>`` child elements
        of *ser* based on the values in *series_data*.
        """
        ser._remove_tx()
        ser._remove_xVal()
        ser._remove_yVal()
        ser._remove_bubbleSize()

        xml_writer = _BubbleSeriesXmlWriter(series_data)

        ser._insert_tx(xml_writer.tx)
        ser._insert_xVal(xml_writer.xVal)
        ser._insert_yVal(xml_writer.yVal)
        ser._insert_bubbleSize(xml_writer.bubbleSize)


class _CategorySeriesXmlRewriter(_BaseSeriesXmlRewriter):
    """
    A series rewriter suitable for category charts.
    """
    def _rewrite_ser_data(self, ser, series_data, date_1904):
        """
        Rewrite the ``<c:tx>``, ``<c:cat>`` and ``<c:val>`` child elements
        of *ser* based on the values in *series_data*.
        """
        ser._remove_tx()
        ser._remove_cat()
        ser._remove_val()

        xml_writer = _CategorySeriesXmlWriter(series_data, date_1904)

        ser._insert_tx(xml_writer.tx)
        ser._insert_cat(xml_writer.cat)
        ser._insert_val(xml_writer.val)


class _XySeriesXmlRewriter(_BaseSeriesXmlRewriter):
    """
    A series rewriter suitable for XY (aka. scatter) charts.
    """
    def _rewrite_ser_data(self, ser, series_data, date_1904):
        """
        Rewrite the ``<c:tx>``, ``<c:xVal>`` and ``<c:yVal>`` child elements
        of *ser* based on the values in *series_data*.
        """
        ser._remove_tx()
        ser._remove_xVal()
        ser._remove_yVal()

        xml_writer = _XySeriesXmlWriter(series_data)

        ser._insert_tx(xml_writer.tx)
        ser._insert_xVal(xml_writer.xVal)
        ser._insert_yVal(xml_writer.yVal)
