#!/usr/bin/env python3
"""
Excel Chart Showcase — generates charts-demo.xlsx with 6 chart types:
clustered bar, smooth line, pie, stacked area, radar, and doughnut. Each
chart is built from a shared 12-month, 4-region sales table on Sheet1.

SDK twin of charts-demo.sh (officecli CLI). Both produce an equivalent
charts-demo.xlsx. This one drives the **officecli Python SDK**
(`pip install officecli-sdk`): one resident is started and the data fill +
chart bodies + drawing anchors are shipped over the named pipe.

Because each chart's drawing anchor must reference the relId returned by the
`add-part` that created the chart part, the build interleaves single
`doc.send(...)` calls (add-part, read back relId) with `doc.batch(...)` for
the bulk data fill. Every item is the same `{"command",...,"props"}` dict you'd
put in an `officecli batch` list — `add-part` / `raw-set` are forwarded verbatim
exactly like the matching CLI commands.

Usage:
  pip install officecli-sdk          # plus the `officecli` binary on PATH
  python3 charts-demo.py
"""

import os
import re
import sys

# --- locate the SDK: prefer an installed `officecli-sdk`, else the in-repo copy
try:
    import officecli  # pip install officecli-sdk
except ImportError:
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    "..", "..", "..", "sdk", "python"))
    import officecli

FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "charts-demo.xlsx")

# --- shared 12-month sales data (10K units) ----------------------------------
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
EAST = [120, 135, 148, 162, 155, 178, 195, 210, 188, 172, 165, 198]
SOUTH = [95, 108, 115, 128, 142, 155, 168, 175, 160, 148, 135, 158]
NORTH = [88, 92, 105, 118, 125, 138, 145, 152, 140, 130, 122, 142]
WEST = [72, 78, 85, 95, 102, 115, 125, 132, 120, 110, 98, 118]


def cell(ref, **props):
    """One `set` item in batch-shape for a Sheet1 cell."""
    return {"command": "set", "path": f"/Sheet1/{ref}", "props": props}


def add_chart_part(doc):
    """Run `add-part --type chart` and return the new chart's relId.
    Mirrors the .sh `grep -o 'relId=...'` step over the JSON envelope."""
    resp = doc.send({"command": "add-part", "parent": "/Sheet1", "type": "chart"})
    data = resp.get("data", "") if isinstance(resp, dict) else str(resp)
    m = re.search(r"relId=(\S+)", data)
    if not m:
        raise RuntimeError(f"could not parse relId from add-part response: {resp!r}")
    return m.group(1)


def chart_body(chart_index, xml):
    """`raw-set` item replacing the chartSpace of /Sheet1/chart[N].
    raw-set takes the target as `part` (the CLI's positional path arg)."""
    return {"command": "raw-set", "part": f"/Sheet1/chart[{chart_index}]",
            "xpath": "/c:chartSpace", "action": "replace", "xml": xml}


def anchor(rel_id, from_col, from_row, to_col, to_row, cnvpr_id, name):
    """`raw-set` item appending a twoCellAnchor graphicFrame to the drawing."""
    xml = (
        "<xdr:twoCellAnchor>"
        f"<xdr:from><xdr:col>{from_col}</xdr:col><xdr:colOff>0</xdr:colOff>"
        f"<xdr:row>{from_row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
        f"<xdr:to><xdr:col>{to_col}</xdr:col><xdr:colOff>0</xdr:colOff>"
        f"<xdr:row>{to_row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>"
        '<xdr:graphicFrame macro="">'
        f'<xdr:nvGraphicFramePr><xdr:cNvPr id="{cnvpr_id}" name="{name}" />'
        "<xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr>"
        '<xdr:xfrm><a:off x="0" y="0" /><a:ext cx="0" cy="0" /></xdr:xfrm>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        f'<c:chart r:id="{rel_id}" /></a:graphicData></a:graphic>'
        "</xdr:graphicFrame><xdr:clientData /></xdr:twoCellAnchor>"
    )
    return {"command": "raw-set", "part": "/Sheet1/drawing",
            "xpath": "//xdr:wsDr", "action": "append", "xml": xml}


# ---- chart bodies (verbatim translations of the .sh chartSpace XML) ----------
BAR_XML = '''
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1400" b="1"><a:solidFill><a:srgbClr val="333333" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1400" b="1" /><a:t>2025 Monthly Sales by Region (10K)</a:t></a:r></a:p>
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
          <c:spPr><a:solidFill><a:srgbClr val="4472C4" /></a:solidFill><a:ln w="0"><a:noFill /></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1" /><c:order val="1" />
          <c:tx><c:strRef><c:f>Sheet1!$C$1</c:f></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="ED7D31" /></a:solidFill><a:ln w="0"><a:noFill /></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$C$2:$C$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="2" /><c:order val="2" />
          <c:tx><c:strRef><c:f>Sheet1!$D$1</c:f></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="70AD47" /></a:solidFill><a:ln w="0"><a:noFill /></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$D$2:$D$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="3" /><c:order val="3" />
          <c:tx><c:strRef><c:f>Sheet1!$E$1</c:f></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="FFC000" /></a:solidFill><a:ln w="0"><a:noFill /></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$E$2:$E$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="1" /><c:axId val="2" />
      </c:barChart>
      <c:catAx><c:axId val="1" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" /><c:crossAx val="2" /></c:catAx>
      <c:valAx><c:axId val="2" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" /><c:numFmt formatCode="#,##0" sourceLinked="0" /><c:crossAx val="1" /></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
    <c:plotVisOnly val="1" />
  </c:chart>
</c:chartSpace>'''

LINE_XML = '''
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1400" b="1"><a:solidFill><a:srgbClr val="333333" /></a:solidFill></a:defRPr></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1400" b="1" /><a:t>Sales Trend Line Chart</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:lineChart>
        <c:grouping val="standard" /><c:varyColors val="0" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="4472C4" /></a:solidFill><a:round /></a:ln></c:spPr>
          <c:marker><c:symbol val="circle" /><c:size val="6" /><c:spPr><a:solidFill><a:srgbClr val="4472C4" /></a:solidFill></c:spPr></c:marker>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$13</c:f></c:numRef></c:val>
          <c:smooth val="1" />
        </c:ser>
        <c:ser>
          <c:idx val="1" /><c:order val="1" />
          <c:tx><c:strRef><c:f>Sheet1!$C$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="ED7D31" /></a:solidFill><a:round /></a:ln></c:spPr>
          <c:marker><c:symbol val="diamond" /><c:size val="6" /><c:spPr><a:solidFill><a:srgbClr val="ED7D31" /></a:solidFill></c:spPr></c:marker>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$C$2:$C$13</c:f></c:numRef></c:val>
          <c:smooth val="1" />
        </c:ser>
        <c:ser>
          <c:idx val="2" /><c:order val="2" />
          <c:tx><c:strRef><c:f>Sheet1!$D$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="70AD47" /></a:solidFill><a:round /></a:ln></c:spPr>
          <c:marker><c:symbol val="triangle" /><c:size val="6" /><c:spPr><a:solidFill><a:srgbClr val="70AD47" /></a:solidFill></c:spPr></c:marker>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$D$2:$D$13</c:f></c:numRef></c:val>
          <c:smooth val="1" />
        </c:ser>
        <c:ser>
          <c:idx val="3" /><c:order val="3" />
          <c:tx><c:strRef><c:f>Sheet1!$E$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="FFC000" /></a:solidFill><a:round /></a:ln></c:spPr>
          <c:marker><c:symbol val="square" /><c:size val="6" /><c:spPr><a:solidFill><a:srgbClr val="FFC000" /></a:solidFill></c:spPr></c:marker>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$E$2:$E$13</c:f></c:numRef></c:val>
          <c:smooth val="1" />
        </c:ser>
        <c:marker val="1" />
        <c:axId val="10" /><c:axId val="20" />
      </c:lineChart>
      <c:catAx><c:axId val="10" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" /><c:crossAx val="20" /></c:catAx>
      <c:valAx><c:axId val="20" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" /><c:numFmt formatCode="#,##0" sourceLinked="0" /><c:crossAx val="10" /></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
    <c:plotVisOnly val="1" />
  </c:chart>
</c:chartSpace>'''

PIE_XML = '''
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1400" b="1" /></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1400" b="1" /><a:t>Annual Regional Sales Share</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:pieChart>
        <c:varyColors val="1" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:dPt><c:idx val="0" /><c:spPr><a:solidFill><a:srgbClr val="4472C4" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="1" /><c:spPr><a:solidFill><a:srgbClr val="ED7D31" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="2" /><c:spPr><a:solidFill><a:srgbClr val="70AD47" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="3" /><c:spPr><a:solidFill><a:srgbClr val="FFC000" /></a:solidFill></c:spPr></c:dPt>
          <c:dLbls>
            <c:showLegendKey val="0" /><c:showVal val="0" /><c:showCatName val="1" /><c:showSerName val="0" /><c:showPercent val="1" />
          </c:dLbls>
          <c:cat><c:strRef><c:f>Sheet1!$B$1:$E$1</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$E$2</c:f></c:numRef></c:val>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
  </c:chart>
</c:chartSpace>'''

AREA_XML = '''
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1400" b="1" /></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1400" b="1" /><a:t>Stacked Area - Sales Composition</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:areaChart>
        <c:grouping val="stacked" /><c:varyColors val="0" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="4472C4"><a:alpha val="80000" /></a:srgbClr></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="4472C4" /></a:solidFill></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1" /><c:order val="1" />
          <c:tx><c:strRef><c:f>Sheet1!$C$1</c:f></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="ED7D31"><a:alpha val="80000" /></a:srgbClr></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="ED7D31" /></a:solidFill></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$C$2:$C$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="2" /><c:order val="2" />
          <c:tx><c:strRef><c:f>Sheet1!$D$1</c:f></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="70AD47"><a:alpha val="80000" /></a:srgbClr></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="70AD47" /></a:solidFill></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$D$2:$D$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="3" /><c:order val="3" />
          <c:tx><c:strRef><c:f>Sheet1!$E$1</c:f></c:strRef></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="FFC000"><a:alpha val="80000" /></a:srgbClr></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="FFC000" /></a:solidFill></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$13</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$E$2:$E$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="30" /><c:axId val="40" />
      </c:areaChart>
      <c:catAx><c:axId val="30" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" /><c:crossAx val="40" /></c:catAx>
      <c:valAx><c:axId val="40" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" /><c:numFmt formatCode="#,##0" sourceLinked="0" /><c:crossAx val="30" /></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
    <c:plotVisOnly val="1" />
  </c:chart>
</c:chartSpace>'''

RADAR_XML = '''
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1400" b="1" /></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1400" b="1" /><a:t>Regional Capability Radar (Q3)</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:radarChart>
        <c:radarStyle val="marker" /><c:varyColors val="0" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="4472C4" /></a:solidFill></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$8:$A$10</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$8:$B$10</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1" /><c:order val="1" />
          <c:tx><c:strRef><c:f>Sheet1!$C$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="ED7D31" /></a:solidFill></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$8:$A$10</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$C$8:$C$10</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="2" /><c:order val="2" />
          <c:tx><c:strRef><c:f>Sheet1!$D$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="70AD47" /></a:solidFill></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$8:$A$10</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$D$8:$D$10</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="3" /><c:order val="3" />
          <c:tx><c:strRef><c:f>Sheet1!$E$1</c:f></c:strRef></c:tx>
          <c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="FFC000" /></a:solidFill></a:ln></c:spPr>
          <c:cat><c:strRef><c:f>Sheet1!$A$8:$A$10</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$E$8:$E$10</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="50" /><c:axId val="60" />
      </c:radarChart>
      <c:catAx><c:axId val="50" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="b" /><c:crossAx val="60" /></c:catAx>
      <c:valAx><c:axId val="60" /><c:scaling><c:orientation val="minMax" /></c:scaling><c:delete val="0" /><c:axPos val="l" /><c:crossAx val="50" /></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
  </c:chart>
</c:chartSpace>'''

DOUGHNUT_XML = '''
<c:chartSpace>
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr /><a:lstStyle />
        <a:p><a:pPr><a:defRPr sz="1400" b="1" /></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1400" b="1" /><a:t>Q4 Regional Sales Doughnut</a:t></a:r></a:p>
      </c:rich></c:tx>
      <c:overlay val="0" />
    </c:title>
    <c:plotArea>
      <c:layout />
      <c:doughnutChart>
        <c:varyColors val="1" />
        <c:ser>
          <c:idx val="0" /><c:order val="0" />
          <c:dPt><c:idx val="0" /><c:spPr><a:solidFill><a:srgbClr val="4472C4" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="1" /><c:spPr><a:solidFill><a:srgbClr val="ED7D31" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="2" /><c:spPr><a:solidFill><a:srgbClr val="70AD47" /></a:solidFill></c:spPr></c:dPt>
          <c:dPt><c:idx val="3" /><c:spPr><a:solidFill><a:srgbClr val="FFC000" /></a:solidFill></c:spPr></c:dPt>
          <c:dLbls>
            <c:showLegendKey val="0" /><c:showVal val="0" /><c:showCatName val="1" /><c:showSerName val="0" /><c:showPercent val="1" />
          </c:dLbls>
          <c:cat><c:strRef><c:f>Sheet1!$B$1:$E$1</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$13:$E$13</c:f></c:numRef></c:val>
        </c:ser>
        <c:holeSize val="50" />
      </c:doughnutChart>
    </c:plotArea>
    <c:legend><c:legendPos val="b" /><c:overlay val="0" /></c:legend>
  </c:chart>
</c:chartSpace>'''


print("\n==========================================")
print(f"Generating Excel chart showcase: {FILE}")
print("==========================================")

with officecli.create(FILE, "--force") as doc:

    # ---------------------------------------------------------------- 1. Data
    print("  -> Populating sales data")
    items = [
        # Header row (different colors per region)
        cell("A1", value="Month", **{"font.bold": "true", "fill": "2F5496",
             "font.color": "FFFFFF", "alignment.horizontal": "center"}),
        cell("B1", value="East Region", **{"font.bold": "true", "fill": "4472C4",
             "font.color": "FFFFFF", "alignment.horizontal": "center"}),
        cell("C1", value="South Region", **{"font.bold": "true", "fill": "5B9BD5",
             "font.color": "FFFFFF", "alignment.horizontal": "center"}),
        cell("D1", value="North Region", **{"font.bold": "true", "fill": "70AD47",
             "font.color": "FFFFFF", "alignment.horizontal": "center"}),
        cell("E1", value="West Region", **{"font.bold": "true", "fill": "FFC000",
             "font.color": "000000", "alignment.horizontal": "center"}),
    ]
    for i in range(12):
        row = i + 2
        items.append(cell(f"A{row}", value=MONTHS[i], **{"alignment.horizontal": "center"}))
        items.append(cell(f"B{row}", value=str(EAST[i]), numFmt="#,##0", **{"alignment.horizontal": "center"}))
        items.append(cell(f"C{row}", value=str(SOUTH[i]), numFmt="#,##0", **{"alignment.horizontal": "center"}))
        items.append(cell(f"D{row}", value=str(NORTH[i]), numFmt="#,##0", **{"alignment.horizontal": "center"}))
        items.append(cell(f"E{row}", value=str(WEST[i]), numFmt="#,##0", **{"alignment.horizontal": "center"}))
    doc.batch(items)
    print("  Done: Data populated")

    # Each chart: add-part (gives a relId) -> raw-set chartSpace -> raw-set anchor.
    # (chart_index is the 1-based /Sheet1/chart[N] ordinal in add-part order.)

    # -------------------------------------------------- 2. Clustered bar chart
    print("  -> Chart 1: Clustered bar chart")
    rel = add_chart_part(doc)
    doc.batch([
        chart_body(1, BAR_XML),
        anchor(rel, 6, 0, 15, 15, 2, "Chart 1"),
    ])
    print("  Done: Clustered bar chart")

    # -------------------------------------------------- 3. Smooth line chart
    print("  -> Chart 2: Smooth line chart")
    rel = add_chart_part(doc)
    doc.batch([
        chart_body(2, LINE_XML),
        anchor(rel, 6, 16, 15, 31, 3, "Chart 2"),
    ])
    print("  Done: Line chart")

    # -------------------------------------------------- 4. Pie chart
    print("  -> Chart 3: Pie chart")
    rel = add_chart_part(doc)
    doc.batch([
        chart_body(3, PIE_XML),
        anchor(rel, 6, 32, 13, 47, 4, "Chart 3"),
    ])
    print("  Done: Pie chart")

    # -------------------------------------------------- 5. Stacked area chart
    print("  -> Chart 4: Stacked area chart")
    rel = add_chart_part(doc)
    doc.batch([
        chart_body(4, AREA_XML),
        anchor(rel, 6, 48, 15, 63, 5, "Chart 4"),
    ])
    print("  Done: Stacked area chart")

    # -------------------------------------------------- 6. Radar chart
    print("  -> Chart 5: Radar chart")
    rel = add_chart_part(doc)
    doc.batch([
        chart_body(5, RADAR_XML),
        anchor(rel, 6, 64, 13, 79, 6, "Chart 5"),
    ])
    print("  Done: Radar chart")

    # -------------------------------------------------- 7. Doughnut chart
    print("  -> Chart 6: Doughnut chart")
    rel = add_chart_part(doc)
    doc.batch([
        chart_body(6, DOUGHNUT_XML),
        anchor(rel, 14, 32, 21, 47, 7, "Chart 6"),
    ])
    print("  Done: Doughnut chart")

    doc.send({"command": "save"})
# context exit closes the resident, flushing the workbook to disk.

print(f"\nGenerated: {FILE}")
