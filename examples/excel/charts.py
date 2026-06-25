#!/usr/bin/env python3
"""
Beautiful Charts Showcase — generates charts.xlsx with 8 chart types built from
raw chart XML: combo (bar+line dual axis), 3D cylinder bar, scatter+trendline,
exploded 3D pie, 3D bubble, stock OHLC candlestick (red up / green down), filled
radar, and a multi-ring (nested) doughnut. 4 data sheets feed them: monthly
sales (Sheet1), spend/sales analysis (Analysis), OHLC stock data (StockData),
and capability assessment (Assessment).

SDK twin of charts.sh (officecli CLI). Both produce an equivalent charts.xlsx.
This one drives the **officecli Python SDK** (`pip install officecli-sdk`): one
resident is started and every command is shipped over the named pipe. Cell data
goes out per-sheet in a single `doc.batch(...)` round-trip; each chart is then a
short sequence of `doc.send(...)` calls — `add-part` (whose returned relId is
captured and substituted into the drawing anchor) plus two `raw-set` calls that
replace the chartSpace XML and append the worksheet drawing anchor. Each item is
the same `{"command",...,"props"}` dict you'd put in an `officecli batch` list;
`add-part`/`raw-set` forward their `type`/`xpath`/`action`/`xml` fields verbatim.

Usage:
  pip install officecli-sdk          # plus the `officecli` binary on PATH
  python3 charts.py
"""

import os
import re
import sys

# --- locate the SDK: prefer an installed `officecli-sdk`, else the in-repo copy
try:
    import officecli  # pip install officecli-sdk
except ImportError:
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    "..", "..", "sdk", "python"))
    import officecli

FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "charts.xlsx")

CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart"


# ---------------------------------------------------------------- batch helpers
def cell(path, value, **props):
    """One `set <cell>` item in batch-shape."""
    return {"command": "set", "path": path, "props": {"value": str(value), **props}}


def add_sheet(name):
    return {"command": "add", "parent": "/", "type": "sheet", "props": {"name": name}}


HDR = {"font.bold": "true", "alignment.horizontal": "center"}


def header(path, value, fill, font_color, size=None):
    p = {"value": value, "fill": fill, "font.color": font_color, **HDR}
    if size is not None:
        p["font.size"] = str(size)
    return {"command": "set", "path": path, "props": p}


# ---------------------------------------------------------------- chart helpers
def add_chart_part(doc, parent):
    """`add-part --type chart`; return the created relationship id.
    as_json=False yields the plain 'Created chart part: relId=... path=...' line
    (same string the .sh greps), from which we pull relId."""
    msg = doc.send({"command": "add-part", "parent": parent, "type": "chart"},
                   as_json=False)
    m = re.search(r"relId=(\S+)", msg if isinstance(msg, str) else str(msg))
    if not m:
        raise RuntimeError(f"add-part did not return a relId: {msg!r}")
    return m.group(1)


def set_chart_xml(doc, chart_path, xml):
    """`raw-set` the whole chartSpace (replace /c:chartSpace). raw-set's target
    part is the `part` field (the CLI positional arg), not `path`."""
    doc.send({"command": "raw-set", "part": chart_path,
              "xpath": "/c:chartSpace", "action": "replace", "xml": xml})


def add_anchor(doc, sheet, from_col, from_row, to_col, to_row, cnvpr_id, name, rel_id):
    """`raw-set` append a twoCellAnchor graphicFrame referencing the chart."""
    xml = (
        '<xdr:twoCellAnchor>'
        f'<xdr:from><xdr:col>{from_col}</xdr:col><xdr:colOff>0</xdr:colOff>'
        f'<xdr:row>{from_row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
        f'<xdr:to><xdr:col>{to_col}</xdr:col><xdr:colOff>0</xdr:colOff>'
        f'<xdr:row>{to_row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
        '<xdr:graphicFrame macro="">'
        f'<xdr:nvGraphicFramePr><xdr:cNvPr id="{cnvpr_id}" name="{name}" />'
        '<xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>'
        '<xdr:xfrm><a:off x="0" y="0" /><a:ext cx="0" cy="0" /></xdr:xfrm>'
        f'<a:graphic><a:graphicData uri="{CHART_URI}">'
        f'<c:chart r:id="{rel_id}" /></a:graphicData></a:graphic>'
        '</xdr:graphicFrame><xdr:clientData />'
        '</xdr:twoCellAnchor>'
    )
    doc.send({"command": "raw-set", "part": f"/{sheet}/drawing",
              "xpath": "//xdr:wsDr", "action": "append", "xml": xml})


# ---------------------------------------------------------------- chart XML
CHART1_XML = '''
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr rot="0" /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="1F4E79" /></a:solidFill><a:latin typeface="Microsoft YaHei" /><a:ea typeface="Microsoft YaHei" /></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1600" b="1"><a:solidFill><a:srgbClr val="1F4E79" /></a:solidFill></a:rPr><a:t>Monthly Sales and YoY Growth Trend</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:barChart>
        <c:barDir val="col" /><c:grouping val="clustered" /><c:varyColors val="0" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:gradFill rotWithShape="1"><a:gsLst>
              <a:gs pos="0"><a:srgbClr val="1F4E79" /></a:gs>
              <a:gs pos="100000"><a:srgbClr val="2E75B6" /></a:gs>
            </a:gsLst><a:lin ang="5400000" /></a:gradFill>
            <a:ln w="0"><a:noFill /></a:ln>
            <a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000" /></a:srgbClr></a:outerShdw></a:effectLst>
          </c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1" /><c:order val="1" />
          <c:tx><c:strRef><c:f>Sheet1!$C$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:gradFill rotWithShape="1"><a:gsLst>
              <a:gs pos="0"><a:srgbClr val="C55A11" /></a:gs>
              <a:gs pos="100000"><a:srgbClr val="ED7D31" /></a:gs>
            </a:gsLst><a:lin ang="5400000" /></a:gradFill>
            <a:ln w="0"><a:noFill /></a:ln>
            <a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000" /></a:srgbClr></a:outerShdw></a:effectLst>
          </c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$C$2:$C$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="2" /><c:order val="2" />
          <c:tx><c:strRef><c:f>Sheet1!$D$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:gradFill rotWithShape="1"><a:gsLst>
              <a:gs pos="0"><a:srgbClr val="548235" /></a:gs>
              <a:gs pos="100000"><a:srgbClr val="70AD47" /></a:gs>
            </a:gsLst><a:lin ang="5400000" /></a:gradFill>
            <a:ln w="0"><a:noFill /></a:ln>
            <a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000" /></a:srgbClr></a:outerShdw></a:effectLst>
          </c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$D$2:$D$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="1" /><c:axId val="2" />
      </c:barChart>
      <c:lineChart>
        <c:grouping val="standard" /><c:varyColors val="0" />
        <c:ser>
          <c:idx val="3" /><c:order val="3" />
          <c:tx><c:strRef><c:f>Sheet1!$F$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="38100" cap="rnd"><a:solidFill><a:srgbClr val="FF0000" /></a:solidFill><a:prstDash val="solid" /><a:round /></a:ln></c:spPr>
          <c:marker><c:symbol val="circle" /><c:size val="8" />
            <c:spPr><a:solidFill><a:srgbClr val="FF0000" /></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF" /></a:solidFill></a:ln></c:spPr>
          </c:marker>
          <c:dLbls>
            <c:numFmt formatCode="0.0&quot;%&quot;" sourceLinked="0" />
            <c:spPr><a:noFill /><a:ln><a:noFill /></a:ln></c:spPr>
            <c:txPr><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="900" b="1"><a:solidFill><a:srgbClr val="FF0000" /></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US" /></a:p></c:txPr>
            <c:showLegendKey val="0" /><c:showVal val="1" /><c:showCatName val="0" /><c:showSerName val="0" /><c:showPercent val="0" />
          </c:dLbls>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$F$2:$F$13</c:f></c:numRef></c:val>
          <c:smooth val="1" />
        </c:ser>
        <c:marker val="1" />
        <c:axId val="1" /><c:axId val="3" />
      </c:lineChart>
      <c:catAx>
        <c:axId val="1" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" />
        <c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF" /></a:solidFill></a:ln></c:spPr>
        <c:txPr><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="404040" /></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US" /></a:p></c:txPr>
        <c:crossAx val="2" />
      </c:catAx>
      <c:valAx>
        <c:axId val="2" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" />
        <c:title><c:tx><c:rich><a:bodyPr rot="-5400000" /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="404040" /></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr lang="en-US" sz="1000" /><a:t>Sales (10K)</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:numFmt formatCode="#,##0" sourceLinked="0" />
        <c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF" /></a:solidFill></a:ln></c:spPr>
        <c:crossAx val="1" />
      </c:valAx>
      <c:valAx>
        <c:axId val="3" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="r" />
        <c:title><c:tx><c:rich><a:bodyPr rot="5400000" /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="FF0000" /></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr lang="en-US" sz="1000" /><a:t>YoY Growth (%)</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:numFmt formatCode="0.0&quot;%&quot;" sourceLinked="0" />
        <c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000" /></a:srgbClr></a:solidFill></a:ln></c:spPr>
        <c:crossAx val="1" /><c:crosses val="max" />
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" />
      <c:txPr><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="404040" /></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US" /></a:p></c:txPr>
    </c:legend>
    <c:plotVisOnly val="1" />
  </c:chart>
</c:chartSpace>'''

CHART2_XML = '''
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="1F4E79" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1600" b="1" /><a:t>3D Regional Sales Comparison</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:view3D>
      <c:rotX val="15" /><c:rotY val="20" /><c:depthPercent val="100" /><c:rAngAx val="1" /><c:perspective val="30" />
    </c:view3D>
    <c:plotArea>
      <c:layout />
      <c:bar3DChart>
        <c:barDir val="col" /><c:grouping val="clustered" /><c:varyColors val="0" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:gradFill><a:gsLst>
              <a:gs pos="0"><a:srgbClr val="4472C4" /></a:gs>
              <a:gs pos="50000"><a:srgbClr val="5B9BD5" /></a:gs>
              <a:gs pos="100000"><a:srgbClr val="9DC3E6" /></a:gs>
            </a:gsLst><a:lin ang="5400000" /></a:gradFill>
          </c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1" /><c:order val="1" />
          <c:tx><c:strRef><c:f>Sheet1!$C$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:gradFill><a:gsLst>
              <a:gs pos="0"><a:srgbClr val="ED7D31" /></a:gs>
              <a:gs pos="50000"><a:srgbClr val="F4B183" /></a:gs>
              <a:gs pos="100000"><a:srgbClr val="F8CBAD" /></a:gs>
            </a:gsLst><a:lin ang="5400000" /></a:gradFill>
          </c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$C$2:$C$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="2" /><c:order val="2" />
          <c:tx><c:strRef><c:f>Sheet1!$D$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:gradFill><a:gsLst>
              <a:gs pos="0"><a:srgbClr val="70AD47" /></a:gs>
              <a:gs pos="50000"><a:srgbClr val="A9D18E" /></a:gs>
              <a:gs pos="100000"><a:srgbClr val="C5E0B4" /></a:gs>
            </a:gsLst><a:lin ang="5400000" /></a:gradFill>
          </c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$D$2:$D$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:shape val="cylinder" />
        <c:axId val="10" /><c:axId val="20" /><c:axId val="30" />
      </c:bar3DChart>
      <c:catAx><c:axId val="10" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" /><c:crossAx val="20" /></c:catAx>
      <c:valAx><c:axId val="20" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" /><c:numFmt formatCode="#,##0" sourceLinked="0" /><c:crossAx val="10" /></c:valAx>
      <c:serAx><c:axId val="30" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" /><c:crossAx val="20" /></c:serAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
    <c:plotVisOnly val="1" />
  </c:chart>
</c:chartSpace>'''

CHART3_XML = '''
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="7030A0" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1600" b="1" /><a:t>Ad Spend vs Sales Correlation</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:scatterChart>
        <c:scatterStyle val="lineMarker" />
        <c:varyColors val="0" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>Analysis!$B$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="0"><a:noFill /></a:ln></c:spPr>
          <c:marker><c:symbol val="circle" /><c:size val="10" />
            <c:spPr>
              <a:solidFill><a:srgbClr val="7030A0"><a:alpha val="70000" /></a:srgbClr></a:solidFill>
              <a:ln w="19050"><a:solidFill><a:srgbClr val="7030A0" /></a:solidFill></a:ln>
              <a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000"><a:srgbClr val="000000"><a:alpha val="30000" /></a:srgbClr></a:outerShdw></a:effectLst>
            </c:spPr>
          </c:marker>
          <c:trendline>
            <c:spPr><a:ln w="25400" cap="rnd"><a:solidFill><a:srgbClr val="FF0000" /></a:solidFill><a:prstDash val="dash" /><a:round /></a:ln></c:spPr>
            <c:trendlineType val="linear" />
            <c:dispRSqr val="1" /><c:dispEq val="1" />
          </c:trendline>
          <c:xVal><c:numRef><c:f>Analysis!$A$2:$A$16</c:f></c:numRef></c:xVal>
          <c:yVal><c:numRef><c:f>Analysis!$B$2:$B$16</c:f></c:numRef></c:yVal>
          <c:smooth val="0" />
        </c:ser>
        <c:axId val="100" /><c:axId val="200" />
      </c:scatterChart>
      <c:valAx>
        <c:axId val="100" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" />
        <c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1000" /></a:pPr><a:r><a:rPr lang="en-US" sz="1000" /><a:t>Ad Spend (10K)</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:numFmt formatCode="#,##0" sourceLinked="0" />
        <c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF" /></a:solidFill></a:ln></c:spPr>
        <c:crossAx val="200" />
      </c:valAx>
      <c:valAx>
        <c:axId val="200" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" />
        <c:title><c:tx><c:rich><a:bodyPr rot="-5400000" /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1000" /></a:pPr><a:r><a:rPr lang="en-US" sz="1000" /><a:t>Sales (10K)</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:numFmt formatCode="#,##0" sourceLinked="0" />
        <c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF" /></a:solidFill></a:ln></c:spPr>
        <c:crossAx val="100" />
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
    <c:plotVisOnly val="1" />
  </c:chart>
</c:chartSpace>'''

CHART4_XML = '''
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="1F4E79" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1600" b="1" /><a:t>Annual Regional Sales Share (3D)</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:view3D>
      <c:rotX val="30" /><c:rotY val="70" /><c:rAngAx val="0" /><c:perspective val="30" />
    </c:view3D>
    <c:plotArea>
      <c:layout />
      <c:pie3DChart>
        <c:varyColors val="1" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:explosion val="10" />
          <c:dPt><c:idx val="0" />
            <c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="1F4E79" /></a:gs><a:gs pos="100000"><a:srgbClr val="4472C4" /></a:gs></a:gsLst><a:lin ang="5400000" /></a:gradFill>
            <a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="5400000"><a:srgbClr val="000000"><a:alpha val="40000" /></a:srgbClr></a:outerShdw></a:effectLst></c:spPr>
          </c:dPt>
          <c:dPt><c:idx val="1" />
            <c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="C55A11" /></a:gs><a:gs pos="100000"><a:srgbClr val="ED7D31" /></a:gs></a:gsLst><a:lin ang="5400000" /></a:gradFill>
            <a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="5400000"><a:srgbClr val="000000"><a:alpha val="40000" /></a:srgbClr></a:outerShdw></a:effectLst></c:spPr>
          </c:dPt>
          <c:dPt><c:idx val="2" />
            <c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="548235" /></a:gs><a:gs pos="100000"><a:srgbClr val="70AD47" /></a:gs></a:gsLst><a:lin ang="5400000" /></a:gradFill>
            <a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="5400000"><a:srgbClr val="000000"><a:alpha val="40000" /></a:srgbClr></a:outerShdw></a:effectLst></c:spPr>
          </c:dPt>
          <c:dLbls>
            <c:numFmt formatCode="0.0&quot;%&quot;" sourceLinked="0" />
            <c:spPr><a:noFill /><a:ln><a:noFill /></a:ln></c:spPr>
            <c:txPr><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1100" b="1"><a:solidFill><a:srgbClr val="FFFFFF" /></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US" /></a:p></c:txPr>
            <c:showLegendKey val="0" /><c:showVal val="0" /><c:showCatName val="1" /><c:showSerName val="0" /><c:showPercent val="1" />
          </c:dLbls>
          <c:cat><c:strRef><c:f>Sheet1!$B$1:$D$1</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$8:$D$8</c:f></c:numRef></c:val>
        </c:ser>
      </c:pie3DChart>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
  </c:chart>
</c:chartSpace>'''

CHART5_XML = '''
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="7030A0" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1600" b="1" /><a:t>Spend-Revenue-Market Share Bubble</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:bubbleChart>
        <c:varyColors val="0" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>Analysis!$D$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:solidFill><a:srgbClr val="7030A0"><a:alpha val="60000" /></a:srgbClr></a:solidFill>
            <a:ln w="19050"><a:solidFill><a:srgbClr val="7030A0" /></a:solidFill></a:ln>
            <a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000"><a:srgbClr val="000000"><a:alpha val="25000" /></a:srgbClr></a:outerShdw></a:effectLst>
          </c:spPr>
          <c:xVal><c:numRef><c:f>Analysis!$A$2:$A$16</c:f></c:numRef></c:xVal>
          <c:yVal><c:numRef><c:f>Analysis!$B$2:$B$16</c:f></c:numRef></c:yVal>
          <c:bubbleSize><c:numRef><c:f>Analysis!$D$2:$D$16</c:f></c:numRef></c:bubbleSize>
          <c:bubble3D val="1" />
        </c:ser>
        <c:axId val="300" /><c:axId val="400" />
      </c:bubbleChart>
      <c:valAx>
        <c:axId val="300" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" />
        <c:title><c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1000" /></a:pPr><a:r><a:rPr lang="en-US" sz="1000" /><a:t>Ad Spend (10K)</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:numFmt formatCode="#,##0" sourceLinked="0" /><c:crossAx val="400" />
      </c:valAx>
      <c:valAx>
        <c:axId val="400" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" />
        <c:title><c:tx><c:rich><a:bodyPr rot="-5400000" /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="1000" /></a:pPr><a:r><a:rPr lang="en-US" sz="1000" /><a:t>Sales (10K)</a:t></a:r></a:p></c:rich></c:tx></c:title>
        <c:numFmt formatCode="#,##0" sourceLinked="0" /><c:crossAx val="300" />
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
    <c:plotVisOnly val="1" />
  </c:chart>
</c:chartSpace>'''

CHART6_XML = '''
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="C00000" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1600" b="1" /><a:t>Stock Candlestick Chart (OHLC)</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:stockChart>
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>StockData!$B$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="0"><a:noFill /></a:ln></c:spPr>
          <c:marker><c:symbol val="none" /></c:marker>
          <c:cat><c:strRef><c:f>StockData!$A$2:$A$21</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>StockData!$B$2:$B$21</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1" /><c:order val="1" />
          <c:tx><c:strRef><c:f>StockData!$C$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="0"><a:noFill /></a:ln></c:spPr>
          <c:marker><c:symbol val="none" /></c:marker>
          <c:cat><c:strRef><c:f>StockData!$A$2:$A$21</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>StockData!$C$2:$C$21</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="2" /><c:order val="2" />
          <c:tx><c:strRef><c:f>StockData!$D$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="0"><a:noFill /></a:ln></c:spPr>
          <c:marker><c:symbol val="none" /></c:marker>
          <c:cat><c:strRef><c:f>StockData!$A$2:$A$21</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>StockData!$D$2:$D$21</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="3" /><c:order val="3" />
          <c:tx><c:strRef><c:f>StockData!$E$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="0"><a:noFill /></a:ln></c:spPr>
          <c:marker><c:symbol val="none" /></c:marker>
          <c:cat><c:strRef><c:f>StockData!$A$2:$A$21</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>StockData!$E$2:$E$21</c:f></c:numRef></c:val>
        </c:ser>
        <c:hiLowLines>
          <c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="404040" /></a:solidFill></a:ln></c:spPr>
        </c:hiLowLines>
        <c:upDownBars>
          <c:gapWidth val="100" />
          <c:upBars><c:spPr><a:solidFill><a:srgbClr val="FF0000" /></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="C00000" /></a:solidFill></a:ln></c:spPr></c:upBars>
          <c:downBars><c:spPr><a:solidFill><a:srgbClr val="00B050" /></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="006400" /></a:solidFill></a:ln></c:spPr></c:downBars>
        </c:upDownBars>
        <c:axId val="500" /><c:axId val="600" />
      </c:stockChart>
      <c:catAx>
        <c:axId val="500" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" />
        <c:txPr><a:bodyPr rot="-5400000" /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="800" /></a:pPr><a:endParaRPr lang="en-US" /></a:p></c:txPr>
        <c:crossAx val="600" />
      </c:catAx>
      <c:valAx>
        <c:axId val="600" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" />
        <c:numFmt formatCode="0.00" sourceLinked="0" />
        <c:crossAx val="500" />
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
    <c:plotVisOnly val="1" />
  </c:chart>
</c:chartSpace>'''

CHART7_XML = '''
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="002060" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1600" b="1" /><a:t>Product Capability Radar Comparison</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:radarChart>
        <c:radarStyle val="filled" /><c:varyColors val="0" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>Assessment!$B$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:solidFill><a:srgbClr val="4472C4"><a:alpha val="40000" /></a:srgbClr></a:solidFill>
            <a:ln w="28575"><a:solidFill><a:srgbClr val="4472C4" /></a:solidFill></a:ln>
          </c:spPr>
          <c:cat><c:strRef><c:f>Assessment!$A$2:$A$9</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Assessment!$B$2:$B$9</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1" /><c:order val="1" />
          <c:tx><c:strRef><c:f>Assessment!$C$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:solidFill><a:srgbClr val="00B050"><a:alpha val="40000" /></a:srgbClr></a:solidFill>
            <a:ln w="28575"><a:solidFill><a:srgbClr val="00B050" /></a:solidFill></a:ln>
          </c:spPr>
          <c:cat><c:strRef><c:f>Assessment!$A$2:$A$9</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Assessment!$C$2:$C$9</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="2" /><c:order val="2" />
          <c:tx><c:strRef><c:f>Assessment!$D$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:solidFill><a:srgbClr val="FFC000"><a:alpha val="40000" /></a:srgbClr></a:solidFill>
            <a:ln w="28575"><a:solidFill><a:srgbClr val="FFC000" /></a:solidFill></a:ln>
          </c:spPr>
          <c:cat><c:strRef><c:f>Assessment!$A$2:$A$9</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Assessment!$D$2:$D$9</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="700" /><c:axId val="800" />
      </c:radarChart>
      <c:catAx><c:axId val="700" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" /><c:crossAx val="800" /></c:catAx>
      <c:valAx><c:axId val="800" /><c:scaling><c:orientation val="minMax" /><c:max val="100" /><c:min val="0" /></c:scaling><c:delete val="0" /><c:axPos val="l" /><c:crossAx val="700" /></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
  </c:chart>
</c:chartSpace>'''

CHART8_XML = '''
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1600" b="1"><a:solidFill><a:srgbClr val="1F4E79" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1600" b="1" /><a:t>Q3 vs Q4 Regional Sales Multi-Ring</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:doughnutChart>
        <c:varyColors val="1" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:v>Q3</c:v></c:tx>
          <c:dPt><c:idx val="0" /><c:spPr><a:solidFill><a:srgbClr val="1F4E79" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="1" /><c:spPr><a:solidFill><a:srgbClr val="C55A11" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="2" /><c:spPr><a:solidFill><a:srgbClr val="548235" /></a:solidFill></c:spPr></c:dPt>
          <c:dLbls>
            <c:numFmt formatCode="0.0&quot;%&quot;" sourceLinked="0" />
            <c:spPr><a:noFill /><a:ln><a:noFill /></a:ln></c:spPr>
            <c:txPr><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="900" b="1"><a:solidFill><a:srgbClr val="FFFFFF" /></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US" /></a:p></c:txPr>
            <c:showLegendKey val="0" /><c:showVal val="0" /><c:showCatName val="0" /><c:showSerName val="0" /><c:showPercent val="1" />
          </c:dLbls>
          <c:cat><c:strRef><c:f>Sheet1!$B$1:$D$1</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$9:$D$9</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1" /><c:order val="1" />
          <c:tx><c:v>Q4</c:v></c:tx>
          <c:dPt><c:idx val="0" /><c:spPr><a:solidFill><a:srgbClr val="4472C4" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="1" /><c:spPr><a:solidFill><a:srgbClr val="ED7D31" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="2" /><c:spPr><a:solidFill><a:srgbClr val="70AD47" /></a:solidFill></c:spPr></c:dPt>
          <c:dLbls>
            <c:numFmt formatCode="0.0&quot;%&quot;" sourceLinked="0" />
            <c:spPr><a:noFill /><a:ln><a:noFill /></a:ln></c:spPr>
            <c:txPr><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz="900" b="1"><a:solidFill><a:srgbClr val="FFFFFF" /></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US" /></a:p></c:txPr>
            <c:showLegendKey val="0" /><c:showVal val="0" /><c:showCatName val="1" /><c:showSerName val="0" /><c:showPercent val="1" />
          </c:dLbls>
          <c:cat><c:strRef><c:f>Sheet1!$B$1:$D$1</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$13:$D$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:holeSize val="40" />
      </c:doughnutChart>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
  </c:chart>
</c:chartSpace>'''


print("\n==========================================")
print(f"Generating beautiful charts document: {FILE}")
print("==========================================")

with officecli.create(FILE, "--force") as doc:

    # ======================================================================
    # Sheet1: Monthly sales data
    # ======================================================================
    print("  -> Populating Sheet1: Monthly sales data")
    s1 = [
        header("/Sheet1/A1", "Month", "1F4E79", "FFFFFF", 11),
        header("/Sheet1/B1", "East Sales", "2E75B6", "FFFFFF", 11),
        header("/Sheet1/C1", "South Sales", "9DC3E6", "1F4E79", 11),
        header("/Sheet1/D1", "North Sales", "BDD7EE", "1F4E79", 11),
        header("/Sheet1/E1", "Total", "C55A11", "FFFFFF", 11),
        header("/Sheet1/F1", "YoY Growth %", "548235", "FFFFFF", 11),
    ]
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    east = [120, 135, 148, 162, 155, 178, 195, 210, 188, 172, 165, 198]
    south = [95, 108, 115, 128, 142, 155, 168, 175, 160, 148, 135, 158]
    north = [88, 92, 105, 118, 125, 138, 145, 152, 140, 130, 122, 142]
    total = [303, 335, 368, 408, 422, 471, 508, 537, 488, 450, 422, 498]
    growth = [5.2, 8.1, 12.3, 15.6, 10.2, 18.5, 22.1, 25.3, 16.8, 11.2, 7.5, 19.8]
    for i in range(12):
        r = i + 2
        s1.append(cell(f"/Sheet1/A{r}", months[i], **{"alignment.horizontal": "center"}))
        s1.append(cell(f"/Sheet1/B{r}", east[i], numFmt="#,##0", **{"alignment.horizontal": "center"}))
        s1.append(cell(f"/Sheet1/C{r}", south[i], numFmt="#,##0", **{"alignment.horizontal": "center"}))
        s1.append(cell(f"/Sheet1/D{r}", north[i], numFmt="#,##0", **{"alignment.horizontal": "center"}))
        s1.append(cell(f"/Sheet1/E{r}", total[i], numFmt="#,##0",
                       **{"font.bold": "true", "alignment.horizontal": "center"}))
        s1.append(cell(f"/Sheet1/F{r}", growth[i], numFmt='0.0"%"', **{"alignment.horizontal": "center"}))
    doc.batch(s1)
    print("  Done: Sheet1 data")

    # ======================================================================
    # Sheet2: Analysis (scatter/bubble) data
    # ======================================================================
    print("  -> Populating Sheet2: Analysis data")
    s2 = [add_sheet("Analysis")]
    for col, title in zip("ABCD", ["Ad Spend (10K)", "Sales (10K)", "Margin %", "Market Share %"]):
        s2.append(header(f"/Analysis/{col}1", title, "7030A0", "FFFFFF"))
    ad_spend = [10, 15, 22, 28, 35, 42, 50, 58, 65, 72, 80, 88, 95, 105, 115]
    sales_rev = [45, 68, 95, 120, 155, 180, 220, 260, 290, 335, 370, 410, 445, 500, 550]
    profit = [8.5, 10.2, 12.1, 14.5, 16.8, 15.2, 18.3, 20.1, 19.5, 22.3, 21.8, 24.5, 23.1, 26.8, 28.2]
    mkt_share = [2.1, 3.2, 4.5, 5.8, 7.2, 8.5, 10.1, 11.8, 12.5, 14.2, 15.8, 17.5, 18.2, 20.5, 22.1]
    for i in range(15):
        r = i + 2
        for col, vals in zip("ABCD", [ad_spend, sales_rev, profit, mkt_share]):
            s2.append(cell(f"/Analysis/{col}{r}", vals[i], **{"alignment.horizontal": "center"}))
    doc.batch(s2)
    print("  Done: Sheet2 data")

    # ======================================================================
    # Sheet3: StockData (red up / green down coloring)
    # ======================================================================
    print("  -> Populating Sheet3: Stock data")
    s3 = [add_sheet("StockData")]
    for col, title in zip("ABCDEF", ["Date", "Open", "High", "Low", "Close", "Volume (10K)"]):
        s3.append(header(f"/StockData/{col}1", title, "C00000", "FFFFFF"))
    dates = ["3/1", "3/2", "3/3", "3/4", "3/5", "3/6", "3/7", "3/8", "3/9", "3/10",
             "3/11", "3/12", "3/13", "3/14", "3/15", "3/16", "3/17", "3/18", "3/19", "3/20"]
    s_open = [52.3, 53.1, 52.8, 54.2, 55.1, 54.5, 56.2, 57.8, 58.5, 57.2,
              56.8, 58.3, 59.5, 60.2, 59.8, 61.5, 62.3, 61.8, 63.5, 64.2]
    s_high = [53.8, 54.2, 54.5, 55.8, 56.3, 56.8, 58.1, 59.2, 59.8, 58.5,
              58.2, 59.8, 61.2, 61.5, 61.8, 63.2, 63.8, 63.5, 65.2, 65.8]
    s_low = [51.5, 52.2, 51.8, 53.5, 54.2, 53.8, 55.5, 56.8, 57.2, 56.1,
             55.8, 57.5, 58.8, 59.2, 58.5, 60.8, 61.2, 60.5, 62.8, 63.5]
    s_close = [53.1, 52.8, 54.2, 55.1, 54.5, 56.2, 57.8, 58.5, 57.2, 56.8,
               58.3, 59.5, 60.2, 59.8, 61.5, 62.3, 61.8, 63.5, 64.2, 65.1]
    volume = [285, 312, 268, 345, 298, 378, 425, 468, 395, 310,
              352, 415, 485, 442, 368, 512, 548, 478, 562, 598]
    for i in range(20):
        r = i + 2
        if s_close[i] > s_open[i]:
            color, bg = "FF0000", "FFF2F2"   # Up: red
        elif s_close[i] < s_open[i]:
            color, bg = "008000", "F2FFF2"   # Down: green
        else:
            color, bg = "666666", "F5F5F5"   # Flat: gray
        common = {"alignment.horizontal": "center", "font.color": color, "fill": bg}
        s3.append(cell(f"/StockData/A{r}", dates[i], **common))
        s3.append(cell(f"/StockData/B{r}", s_open[i], numFmt="0.00", **common))
        s3.append(cell(f"/StockData/C{r}", s_high[i], numFmt="0.00", **common))
        s3.append(cell(f"/StockData/D{r}", s_low[i], numFmt="0.00", **common))
        s3.append(cell(f"/StockData/E{r}", s_close[i], numFmt="0.00", **common))
        s3.append(cell(f"/StockData/F{r}", volume[i], numFmt="#,##0", **common))
    doc.batch(s3)
    print("  Done: Sheet3 stock data (with red/green coloring)")

    # ======================================================================
    # Sheet4: Assessment (radar) data
    # ======================================================================
    print("  -> Populating Sheet4: Capability assessment")
    s4 = [add_sheet("Assessment")]
    s4.append(header("/Assessment/A1", "Dimension", "002060", "FFFFFF"))
    s4.append(header("/Assessment/B1", "Product A", "0070C0", "FFFFFF"))
    s4.append(header("/Assessment/C1", "Product B", "00B050", "FFFFFF"))
    s4.append(header("/Assessment/D1", "Product C", "FFC000", "000000"))
    dims = ["Performance", "Stability", "Usability", "Security",
            "Scalability", "Value", "Ecosystem", "Docs"]
    pa = [92, 88, 75, 95, 82, 70, 85, 78]
    pb = [78, 92, 88, 80, 90, 85, 72, 82]
    pc = [85, 76, 92, 72, 78, 92, 88, 70]
    for i in range(8):
        r = i + 2
        for col, vals in zip("ABCD", [dims, pa, pb, pc]):
            s4.append(cell(f"/Assessment/{col}{r}", vals[i], **{"alignment.horizontal": "center"}))
    doc.batch(s4)
    print("  Done: Sheet4 data")

    # ======================================================================
    # Charts — each: add-part (capture relId) → replace chartSpace → anchor
    # ======================================================================
    print("  -> Chart 1: Combo chart (bar + line dual axis)")
    rel = add_chart_part(doc, "/Sheet1")
    set_chart_xml(doc, "/Sheet1/chart[1]", CHART1_XML)
    add_anchor(doc, "Sheet1", 7, 0, 18, 18, 2, "Chart 1", rel)
    print("  Done: Chart 1 combo chart")

    print("  -> Chart 2: 3D bar chart")
    rel = add_chart_part(doc, "/Sheet1")
    set_chart_xml(doc, "/Sheet1/chart[2]", CHART2_XML)
    add_anchor(doc, "Sheet1", 7, 19, 18, 37, 3, "Chart 2", rel)
    print("  Done: Chart 2 3D bar chart")

    print("  -> Chart 3: Scatter plot + trendline")
    rel = add_chart_part(doc, "/Analysis")
    set_chart_xml(doc, "/Analysis/chart[1]", CHART3_XML)
    add_anchor(doc, "Analysis", 5, 0, 16, 18, 2, "Chart 3", rel)
    print("  Done: Chart 3 scatter plot")

    print("  -> Chart 4: 3D pie chart (exploded)")
    rel = add_chart_part(doc, "/Sheet1")
    set_chart_xml(doc, "/Sheet1/chart[3]", CHART4_XML)
    add_anchor(doc, "Sheet1", 19, 0, 28, 18, 4, "Chart 4", rel)
    print("  Done: Chart 4 3D pie chart")

    print("  -> Chart 5: Bubble chart")
    rel = add_chart_part(doc, "/Analysis")
    set_chart_xml(doc, "/Analysis/chart[2]", CHART5_XML)
    add_anchor(doc, "Analysis", 5, 19, 16, 37, 3, "Chart 5", rel)
    print("  Done: Chart 5 bubble chart")

    print("  -> Chart 6: Stock OHLC chart")
    rel = add_chart_part(doc, "/StockData")
    set_chart_xml(doc, "/StockData/chart[1]", CHART6_XML)
    add_anchor(doc, "StockData", 7, 0, 20, 22, 2, "Chart 6", rel)
    print("  Done: Chart 6 stock OHLC chart")

    print("  -> Chart 7: Filled radar chart")
    rel = add_chart_part(doc, "/Assessment")
    set_chart_xml(doc, "/Assessment/chart[1]", CHART7_XML)
    add_anchor(doc, "Assessment", 5, 0, 16, 20, 2, "Chart 7", rel)
    print("  Done: Chart 7 radar chart")

    print("  -> Chart 8: Multi-ring doughnut chart")
    rel = add_chart_part(doc, "/Sheet1")
    set_chart_xml(doc, "/Sheet1/chart[4]", CHART8_XML)
    add_anchor(doc, "Sheet1", 19, 19, 28, 37, 5, "Chart 8", rel)
    print("  Done: Chart 8 multi-ring doughnut chart")

    doc.send({"command": "save"})
# context exit closes the resident, flushing the workbook to disk.

print(f"\nGenerated: {FILE}")
