# encoding: utf-8

"""Shared oxml objects for charts."""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
from pptx.oxml.simpletypes import (
    ST_LayoutMode, XsdBoolean, XsdDouble, XsdString, XsdUnsignedInt
)
from pptx.oxml.xmlchemy import (
    BaseOxmlElement, OptionalAttribute, RequiredAttribute, ZeroOrOne
)


class CT_Boolean(BaseOxmlElement):
    """
    Common complex type used for elements having a True/False value.
    """
    val = OptionalAttribute('val', XsdBoolean, default=True)


class CT_Boolean_Explicit(BaseOxmlElement):
    """Always spells out the `val` attribute, e.g. `val=1`.

    At least one boolean element is improperly interpreted by one or more
    versions of PowerPoint. The `c:overlay` element is interpreted as |False|
    when no `val` attribute is present, contrary to the behavior described in
    the schema. A remedy for this is to interpret a missing `val` attribute
    as |True| (consistent with the spec), but always write the attribute
    whenever there is occasion for changing the element.
    """
    _val = OptionalAttribute('val', XsdBoolean, default=True)

    @property
    def val(self):
        return self._val

    @val.setter
    def val(self, value):
        val_str = '1' if bool(value) is True else '0'
        self.set('val', val_str)


class CT_Double(BaseOxmlElement):
    """
    Used for floating point values.
    """
    val = RequiredAttribute('val', XsdDouble)


class CT_Layout(BaseOxmlElement):
    """
    ``<c:layout>`` custom element class
    """
    manualLayout = ZeroOrOne('c:manualLayout', successors=('c:extLst',))

    @property
    def horz_offset(self):
        """
        The float value in ./c:manualLayout/c:x when
        c:layout/c:manualLayout/c:xMode@val == "factor". 0.0 if that XPath
        expression finds no match.
        """
        manualLayout = self.manualLayout
        if manualLayout is None:
            return 0.0
        return manualLayout.horz_offset

    @horz_offset.setter
    def horz_offset(self, offset):
        """
        Set the value of ./c:manualLayout/c:x@val to *offset* and
        ./c:manualLayout/c:xMode@val to "factor". Remove ./c:manualLayout if
        *offset* == 0.
        """
        if offset == 0.0:
            self._remove_manualLayout()
            return
        manualLayout = self.get_or_add_manualLayout()
        manualLayout.horz_offset = offset

    @property
    def x(self):
        """
        The float value in ./c:manualLayout/c:x when
        c:layout/c:manualLayout/c:xMode@val == "factor". 0.0 if that XPath
        expression finds no match.
        """
        manualLayout = self.manualLayout
        if manualLayout is None:
            return 0.0
        return manualLayout.x

    @x.setter
    def x(self, value):
        """
        Set the value of ./c:manualLayout/c:x@val to *offset* and
        ./c:manualLayout/c:xMode@val to "factor". Remove ./c:manualLayout if
        *offset* == 0.
        """
        manualLayout = self.get_or_add_manualLayout()
        manualLayout.x = value

    @property
    def y(self):
        """
        The float value in ./c:manualLayout/c:x when
        c:layout/c:manualLayout/c:xMode@val == "factor". 0.0 if that XPath
        expression finds no match.
        """
        return manualLayout.y

    @y.setter
    def y(self, value):
        """
        Set the value of ./c:manualLayout/c:x@val to *offset* and
        ./c:manualLayout/c:xMode@val to "factor". Remove ./c:manualLayout if
        *offset* == 0.
        """
        manualLayout = self.get_or_add_manualLayout()
        manualLayout.y = value

    @property
    def width(self):
        """
        The float value in ./c:manualLayout/c:x when
        c:layout/c:manualLayout/c:xMode@val == "factor". 0.0 if that XPath
        expression finds no match.
        """
        manualLayout = self.manualLayout
        if manualLayout is None:
            return 0.0
        return manualLayout.width

    @width.setter
    def width(self, value):
        """
        Set the value of ./c:manualLayout/c:x@val to *offset* and
        ./c:manualLayout/c:xMode@val to "factor". Remove ./c:manualLayout if
        *offset* == 0.
        """
        manualLayout = self.get_or_add_manualLayout()
        manualLayout.width = value
        
    @property
    def height(self):
        """
        The float value in ./c:manualLayout/c:x when
        c:layout/c:manualLayout/c:xMode@val == "factor". 0.0 if that XPath
        expression finds no match.
        """
        manualLayout = self.manualLayout
        if manualLayout is None:
            return 0.0
        return manualLayout.height

    @height.setter
    def height(self, value):
        """
        Set the value of ./c:manualLayout/c:x@val to *offset* and
        ./c:manualLayout/c:xMode@val to "factor". Remove ./c:manualLayout if
        *offset* == 0.
        """
        manualLayout = self.get_or_add_manualLayout()
        manualLayout.height = value

class CT_LayoutMode(BaseOxmlElement):
    """
    Used for ``<c:xMode>``, ``<c:yMode>``, ``<c:wMode>``, and ``<c:hMode>``
    child elements of CT_ManualLayout.
    """
    val = OptionalAttribute(
        'val', ST_LayoutMode, default=ST_LayoutMode.FACTOR
    )


class CT_ManualLayout(BaseOxmlElement):
    """
    ``<c:manualLayout>`` custom element class
    """
    _tag_seq = (
        'c:layoutTarget', 'c:xMode', 'c:yMode', 'c:wMode', 'c:hMode', 'c:x',
        'c:y', 'c:w', 'c:h', 'c:extLst'
    )
    xMode = ZeroOrOne('c:xMode', successors=_tag_seq[2:])
    yMode = ZeroOrOne('c:yMode', successors=_tag_seq[3:])
    x_ = ZeroOrOne('c:x', successors=_tag_seq[6:])
    y_ = ZeroOrOne('c:y', successors=_tag_seq[7:])
    w = ZeroOrOne('c:w', successors=_tag_seq[8:])
    h = ZeroOrOne('c:h', successors=_tag_seq[9:])
    del _tag_seq

    @property
    def horz_offset(self):
        """
        The float value in ./c:x@val when ./c:xMode@val == "factor". 0.0 when
        ./c:x is not present or ./c:xMode@val != "factor".
        """
        x, xMode = self.x, self.xMode
        if x is None or xMode is None or xMode.val != ST_LayoutMode.FACTOR:
            return 0.0
        return x.val

    @horz_offset.setter
    def horz_offset(self, offset):
        """
        Set the value of ./c:x@val to *offset* and ./c:xMode@val to "factor".
        """
        self.get_or_add_xMode().val = ST_LayoutMode.FACTOR
        self.get_or_add_x_().val = offset

    @property
    def x(self):
        x = self.x
        if x is None:
            return 0.0

        return x.val

    @x.setter
    def x(self, value):
        self.get_or_add_xMode().val = ST_LayoutMode.EDGE
        self.get_or_add_x_().val = value

    @property
    def y(self):
        y = self.y_
        if y is None:
            return 0.0

        return y.val

    @y.setter
    def y(self, value):
        self.get_or_add_yMode().val = ST_LayoutMode.EDGE
        self.get_or_add_y_().val = value

    @property
    def width(self):
        w = self.w
        if w is None:
            return 0.0

        return w.val

    @width.setter
    def width(self, value):
        self.get_or_add_w().val = value

    @property
    def height(self):
        h = self.h
        if h is None:
            return 0.0

        return h.val

    @height.setter
    def height(self, value):
        self.get_or_add_h().val = value

class CT_NumFmt(BaseOxmlElement):
    """
    ``<c:numFmt>`` element specifying the formatting for number labels on a
    tick mark or data point.
    """
    formatCode = RequiredAttribute('formatCode', XsdString)
    sourceLinked = OptionalAttribute('sourceLinked', XsdBoolean)


class CT_Title(BaseOxmlElement):
    """`c:title` custom element class."""

    _tag_seq = (
        'c:tx', 'c:layout', 'c:overlay', 'c:spPr', 'c:txPr', 'c:extLst'
    )
    tx = ZeroOrOne('c:tx', successors=_tag_seq[1:])
    spPr = ZeroOrOne('c:spPr', successors=_tag_seq[4:])
    del _tag_seq

    def get_or_add_tx_rich(self):
        """Return `c:tx/c:rich`, newly created if not present.

        Return the `c:rich` grandchild at `c:tx/c:rich`. Both the `c:tx` and
        `c:rich` elements are created if not already present. Any
        `c:tx/c:strRef` element is removed. (Such an element would contain
        a cell reference for the axis title text in the chart's Excel
        worksheet.)
        """
        tx = self.get_or_add_tx()
        tx._remove_strRef()
        return tx.get_or_add_rich()

    @property
    def tx_rich(self):
        """Return `c:tx/c:rich` or |None| if not present."""
        richs = self.xpath('c:tx/c:rich')
        if not richs:
            return None
        return richs[0]

    @staticmethod
    def new_title():
        """Return "loose" `c:title` element containing default children."""
        return parse_xml(
            '<c:title %s>'
            '  <c:layout/>'
            '  <c:overlay val="0"/>'
            '</c:title>' % nsdecls('c')
        )


class CT_Tx(BaseOxmlElement):
    """
    ``<c:tx>`` element containing the text for a label on a data point or
    other chart item.
    """
    strRef = ZeroOrOne('c:strRef')
    rich = ZeroOrOne('c:rich')

    def _new_rich(self):
        return parse_xml(
            '<c:rich %s>'
            '  <a:bodyPr/>'
            '  <a:lstStyle/>'
            '  <a:p>'
            '    <a:pPr>'
            '      <a:defRPr/>'
            '    </a:pPr>'
            '  </a:p>'
            '</c:rich>' % nsdecls('c', 'a')
        )


class CT_UnsignedInt(BaseOxmlElement):
    """
    ``<c:idx>`` element and others.
    """
    val = RequiredAttribute('val', XsdUnsignedInt)
