"""Local Word visual inventory and conservative, optional vector export.

Native vectors are extracted byte for byte. Raster images remain raster images.
Only simple, fully cached native Word charts can be rebuilt as actual SVG paths
and text by ReportLab. A reconstruction is a new plot of cached data, not a
pixel-identical conversion or a guarantee of APA compliance. Nothing is uploaded
and the source document is never modified.
"""
from __future__ import annotations

import hashlib
import json
import math
from pathlib import Path, PurePosixPath
import posixpath
import re
from zipfile import ZipFile

from lxml import etree

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "v": "urn:schemas-microsoft-com:vml",
    "svg": "http://www.w3.org/2000/svg",
}
VECTOR_EXTENSIONS = {".svg", ".emf", ".wmf"}
RASTER_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp", ".dib", ".ico"}
CHART_TAGS = {"areaChart", "area3DChart", "barChart", "bar3DChart", "bubbleChart",
              "doughnutChart", "lineChart", "line3DChart", "ofPieChart", "pieChart",
              "pie3DChart", "radarChart", "scatterChart", "stockChart", "surfaceChart", "surface3DChart"}


class UnsupportedChart(ValueError):
    """A chart cannot be reconstructed without an unsupported interpretation."""


def _xml(data: bytes):
    return etree.fromstring(data, parser=etree.XMLParser(resolve_entities=False, no_network=True))


def _local(node):
    return etree.QName(node).localname


def _val(node, path, default=None):
    found = node.find(path, NS)
    return default if found is None else found.get("val", default)


def _sha(data):
    return hashlib.sha256(data).hexdigest()


def _relations(package, part):
    path = PurePosixPath(part)
    rel_part = str(path.parent / "_rels" / (path.name + ".rels"))
    if rel_part not in package.namelist():
        return {}
    result = {}
    for rel in _xml(package.read(rel_part)):
        target = rel.get("Target", "")
        external = rel.get("TargetMode") == "External"
        result[rel.get("Id")] = {
            "target": target if external else posixpath.normpath(posixpath.join(str(path.parent), target)).lstrip("/"),
            "external": external,
            "type": rel.get("Type", "").rsplit("/", 1)[-1],
        }
    return result


def _media_kind(part):
    extension = PurePosixPath(part).suffix.lower()
    return "vector_container" if extension in VECTOR_EXTENSIONS else "raster" if extension in RASTER_EXTENSIONS else "other"


def _inspect(package):
    tables, drawings, chart_placements, legacy_images, parse_warnings = [], [], [], [], []
    parts = sorted(package.namelist())
    for part in parts:
        if not part.startswith("word/") or not part.endswith(".xml") or "/charts/" in part:
            continue
        try:
            root = _xml(package.read(part))
            rels = _relations(package, part)
        except etree.XMLSyntaxError as exc:
            parse_warnings.append({"source_part": part, "reason": str(exc)})
            continue
        for index, table in enumerate(root.findall(".//w:tbl", NS), 1):
            depth = sum(parent.tag == "{" + NS["w"] + "}tbl" for parent in table.iterancestors())
            tables.append({"source_part": part, "table_index": index, "nested_depth": depth,
                           "rows": len(table.findall("w:tr", NS)),
                           "grid_columns": len(table.findall("w:tblGrid/w:gridCol", NS))})
        for kind in ("inline", "anchor"):
            for index, drawing in enumerate(root.findall(".//wp:" + kind, NS), 1):
                image_refs = []
                for blip in drawing.findall(".//a:blip", NS):
                    rid = blip.get("{" + NS["r"] + "}embed") or blip.get("{" + NS["r"] + "}link")
                    image_refs.append({"relationship_id": rid, **rels.get(rid, {"target": None, "external": False})})
                # Office SVG drawings usually have a raster fallback plus an SVG extension.
                for svg_blip in drawing.xpath(".//*[local-name()='svgBlip']"):
                    rid = svg_blip.get("{" + NS["r"] + "}embed")
                    image_refs.append({"relationship_id": rid, **rels.get(rid, {"target": None, "external": False}), "svg_extension": True})
                chart_refs = []
                for chart in drawing.findall(".//c:chart", NS):
                    rid = chart.get("{" + NS["r"] + "}id")
                    reference = {"source_part": part, "placement": kind, "relationship_id": rid,
                                 **rels.get(rid, {"target": None, "external": False})}
                    chart_refs.append(reference)
                    chart_placements.append(reference)
                drawings.append({"source_part": part, "placement": kind, "drawing_index": index,
                                 "images": image_refs, "charts": chart_refs})
        for image in root.findall(".//v:imagedata", NS):
            rid = image.get("{" + NS["r"] + "}id")
            legacy_images.append({"source_part": part, "relationship_id": rid,
                                  **rels.get(rid, {"target": None, "external": False})})
    media = [{"source_part": part, "kind": _media_kind(part), "bytes": package.getinfo(part).file_size}
             for part in parts if part.startswith("word/media/") and not part.endswith("/")]
    charts = [{"source_part": part, "placements": [p for p in chart_placements if p["target"] == part]}
              for part in parts if re.fullmatch(r"word/charts/chart[^/]*\.xml", part)]
    return {
        "tables": tables, "drawings": drawings, "native_charts": charts, "media": media,
        "legacy_images": legacy_images, "warnings": parse_warnings,
        "counts": {"native_tables": len(tables), "nested_tables": sum(t["nested_depth"] > 0 for t in tables),
                   "top_level_tables": sum(t["nested_depth"] == 0 for t in tables),
                   "inline_images": sum(d["placement"] == "inline" and bool(d["images"]) for d in drawings),
                   "floating_images": sum(d["placement"] == "anchor" and bool(d["images"]) for d in drawings),
                   "inline_drawings": sum(d["placement"] == "inline" for d in drawings),
                   "floating_drawings": sum(d["placement"] == "anchor" for d in drawings),
                   "native_charts": len(charts), "vector_media": sum(m["kind"] == "vector_container" for m in media),
                   "raster_media": sum(m["kind"] == "raster" for m in media), "legacy_images": len(legacy_images)},
        "scope": "Counts cover packaged Word XML, including headers/footers and nested tables. Drawing alternatives may both be counted; screenshots are images, not inferred native tables/charts.",
    }


def inspect_visuals(docx_path: Path) -> dict:
    """Read native objects and packaged assets without modifying the document."""
    with ZipFile(Path(docx_path)) as package:
        return _inspect(package)


def _cache(container, numeric):
    """Return a dense cache indexed by pt/@idx; never pair by XML order."""
    if container is None:
        raise UnsupportedChart("Missing explicit data cache; no data are inferred from formulas or workbook cells.")
    choices = container.xpath("./c:numLit | ./c:strLit | ./c:numRef/c:numCache | ./c:strRef/c:strCache", namespaces=NS)
    if len(choices) != 1:
        raise UnsupportedChart("Missing or ambiguous single-level data cache (including formula cache gaps).")
    cache = choices[0]
    if numeric and _local(cache) not in {"numLit", "numCache"}:
        raise UnsupportedChart("A numerical axis requires an explicit numeric cache.")
    cache_format = cache.find("c:formatCode", NS)
    if cache_format is not None and cache_format.text not in {None, "General", "0", "0.0", "0.00"}:
        raise UnsupportedChart("Formatted date/percentage/custom cached values require unsupported label interpretation.")
    try:
        count = int(_val(cache, "c:ptCount"))
    except (TypeError, ValueError):
        raise UnsupportedChart("Cache has no valid explicit point count.") from None
    if not 0 < count <= 10000:
        raise UnsupportedChart("Empty or excessively large cache; supported range is 1–10,000 points per series.")
    points = {}
    for pt in cache.findall("c:pt", NS):
        try:
            index = int(pt.get("idx"))
        except (TypeError, ValueError):
            raise UnsupportedChart("Cache point has an invalid index.") from None
        value = pt.find("c:v", NS)
        if index in points or value is None or value.text is None:
            raise UnsupportedChart("Duplicate cache index or missing explicit value.")
        text = value.text
        if numeric:
            try:
                text = float(text)
            except ValueError:
                raise UnsupportedChart("Cache contains a non-numerical or spreadsheet-error value.") from None
            if not math.isfinite(text):
                raise UnsupportedChart("Cache contains a non-finite value.")
        points[index] = text
    if set(points) != set(range(count)):
        raise UnsupportedChart("Cache is incomplete: point indices do not exactly cover 0 through ptCount−1.")
    return [points[index] for index in range(count)]


def _text(node):
    if node is None:
        return ""
    rich = node.findall(".//a:t", NS)
    if rich:
        return "".join(item.text or "" for item in rich)
    reference = node.find("c:tx", NS)
    if reference is not None and reference.find("c:strRef", NS) is not None:
        values = _cache(reference, False)
        if len(values) != 1:
            raise UnsupportedChart("Title/name cache must contain exactly one value.")
        return values[0]
    return ""


def _model(chart_bytes):
    """Parse only a deliberately narrow set of unambiguous 2D plots."""
    root = _xml(chart_bytes)
    chart = root.find("c:chart", NS)
    if chart is None:
        raise UnsupportedChart("Not a supported standard Word chart part (ChartEx is not supported).")
    plot = chart.find("c:plotArea", NS)
    if plot is None:
        raise UnsupportedChart("Missing plot area.")
    plots = [child for child in plot if _local(child) in CHART_TAGS]
    if len(plots) != 1 or _local(plots[0]) not in {"barChart", "lineChart", "scatterChart"}:
        raise UnsupportedChart("Only a single 2D bar, line, or scatter plot is supported; combined, 3D, and other charts are skipped.")
    for tag, reason in {
        "dateAx": "Date axes require calendar-aware plotting.", "logBase": "Logarithmic axes are unsupported.",
        "errBars": "Error bars require source-data-aware reconstruction.", "trendline": "Trendlines are unsupported.",
        "view3D": "3D chart views are unsupported.", "multiLvlStrRef": "Multilevel categories are unsupported.",
        "dLbls": "Data labels/custom annotations are unsupported.", "pivotSource": "Pivot charts are unsupported.",
        "dispUnits": "Scaled display units are unsupported.", "dataTable": "Embedded chart data tables are unsupported.",
        "upDownBars": "Up/down bars are unsupported.", "hiLowLines": "High/low lines are unsupported.",
        "dropLines": "Drop lines are unsupported.", "extLst": "Chart extensions may alter the plotted data or rendering and are unsupported.",
    }.items():
        if root.find(".//c:" + tag, NS) is not None:
            raise UnsupportedChart(reason)
    if any(node.get("val", "1") not in {"0", "false", "off"} for node in root.findall(".//c:smooth", NS)):
        raise UnsupportedChart("Smoothed curves are unsupported; a straight-line substitute would change the figure.")
    axes = [child for child in plot if _local(child) in {"catAx", "valAx", "dateAx", "serAx"}]
    kind = _local(plots[0])
    if len(axes) != 2 or sorted(_local(axis) for axis in axes) != (["valAx", "valAx"] if kind == "scatterChart" else ["catAx", "valAx"]):
        raise UnsupportedChart("Exactly one ordinary pair of axes is required; missing, secondary, or multiple axes are unsupported.")
    ids = [_val(axis, "c:axId") for axis in axes]
    declared = [axis.get("val") for axis in plots[0].findall("c:axId", NS)]
    if None in ids or len(set(ids)) != 2 or len(declared) != 2 or set(ids) != set(declared):
        raise UnsupportedChart("Plot axes cannot be linked unambiguously.")
    for axis in axes:
        if _val(axis, "c:crossAx") not in ids or _val(axis, "c:crossAx") == _val(axis, "c:axId"):
            raise UnsupportedChart("Axis crossing references are incomplete or ambiguous.")
        if _val(axis, "c:scaling/c:orientation", "minMax") != "minMax":
            raise UnsupportedChart("Reversed axes are unsupported.")
        if any(axis.find("c:scaling/c:" + bound, NS) is not None for bound in ("min", "max")):
            raise UnsupportedChart("Explicit axis limits are unsupported; automatic limits must be acceptable.")
        if axis.find("c:crossesAt", NS) is not None:
            raise UnsupportedChart("Custom numerical axis-crossing positions are unsupported.")
        fmt = axis.find("c:numFmt", NS)
        if fmt is not None and fmt.get("formatCode", "General") not in {"General", "0", "0.0", "0.00"}:
            raise UnsupportedChart("Custom/percentage/date axis number formats are unsupported.")
    plot_node = plots[0]
    if kind in {"barChart", "lineChart"}:
        grouping = _val(plot_node, "c:grouping", "clustered" if kind == "barChart" else "standard")
        allowed = {"clustered"} if kind == "barChart" else {"standard"}
        if grouping not in allowed:
            raise UnsupportedChart("Stacked, percentage-stacked, and other nonstandard grouping are unsupported.")
    direction = _val(plot_node, "c:barDir", "col")
    if kind == "barChart" and direction not in {"col", "bar"}:
        raise UnsupportedChart("Unsupported bar direction.")
    scatter_style = _val(plot_node, "c:scatterStyle", "marker")
    if kind == "scatterChart" and scatter_style not in {"marker", "line", "lineMarker"}:
        raise UnsupportedChart("Only unsmoothed marker/line/line-with-marker scatter styles are supported.")
    series_nodes = plot_node.findall("c:ser", NS)
    if not 1 <= len(series_nodes) <= 12:
        raise UnsupportedChart("Supported chart series count is 1–12.")
    try:
        orders = [int(_val(series, "c:order")) for series in series_nodes]
    except (ValueError, TypeError):
        raise UnsupportedChart("Every series must declare an explicit order.") from None
    if len(set(orders)) != len(orders):
        raise UnsupportedChart("Duplicate series order values.")
    series, categories = [], None
    for order, node in sorted(zip(orders, series_nodes), key=lambda item: item[0]):
        name_node = node.find("c:tx", NS)
        name = ""
        if name_node is not None:
            literal = name_node.find("c:v", NS)
            if literal is not None:
                name = literal.text or ""
            else:
                names = _cache(name_node, False)
                if len(names) != 1:
                    raise UnsupportedChart("Series-name cache must contain one explicit value.")
                name = names[0]
        y = _cache(node.find("c:yVal" if kind == "scatterChart" else "c:val", NS), True)
        x = _cache(node.find("c:xVal", NS), True) if kind == "scatterChart" else _cache(node.find("c:cat", NS), False)
        if len(x) != len(y):
            raise UnsupportedChart("Category/X and Y cache lengths differ; pairing would be ambiguous.")
        if kind != "scatterChart":
            if categories is not None and categories != x:
                raise UnsupportedChart("Series have different category caches; no category alignment is inferred.")
            categories = x
        marker = node.find("c:marker/c:symbol", NS)
        series.append({"name": name, "order": order, "x": x, "y": y,
                       "points": list(zip(x, y)), "marker": marker.get("val") if marker is not None else None})
    if categories and (len(categories) > 40 or any(len(label) > 60 for label in categories)):
        raise UnsupportedChart("Category labels exceed the compact reconstruction layout; use the original chart editor.")
    # Use physical axis positions to associate labels; axis IDs alone do not encode X vs Y.
    axis_labels = {"x": "", "y": ""}
    used_positions = set()
    for axis in axes:
        position = _val(axis, "c:axPos")
        logical = "x" if position in {"b", "t"} else "y" if position in {"l", "r"} else None
        if logical is None or logical in used_positions:
            raise UnsupportedChart("Axes do not specify one horizontal and one vertical position.")
        used_positions.add(logical)
        axis_labels[logical] = _text(axis.find("c:title", NS))
    return {"chart_type": kind, "bar_direction": direction, "scatter_style": scatter_style,
            "series": series, "categories": categories, "title": _text(chart.find("c:title", NS)),
            "axis_labels": axis_labels, "legend": chart.find("c:legend", NS) is not None}


def _drawing(model):
    """Build actual ReportLab vector chart objects from validated paired data."""
    from reportlab.graphics.charts.barcharts import HorizontalBarChart, VerticalBarChart
    from reportlab.graphics.charts.linecharts import HorizontalLineChart
    from reportlab.graphics.charts.lineplots import LinePlot
    from reportlab.graphics.charts.legends import Legend
    from reportlab.graphics.shapes import Drawing, String
    from reportlab.graphics.widgets.markers import makeMarker
    from reportlab.lib import colors
    from reportlab.pdfbase.pdfmetrics import stringWidth

    def label_width(label, size=10):
        # SVG retains Unicode text. CJK fallback fonts often need a full em per
        # character, wider than ReportLab's Helvetica substitute metrics.
        return sum(size if ord(char) > 255 else stringWidth(char, "Helvetica", size) for char in label)

    kind, series = model["chart_type"], model["series"]
    horizontal_bar = kind == "barChart" and model["bar_direction"] == "bar"
    category_width = max((label_width(label) for label in (model["categories"] or [])), default=0)
    category_angle = 30 if kind != "scatterChart" and not horizontal_bar and (
        len(model["categories"]) > 8 or category_width > 70) else 0
    legend_visible = model["legend"] and any(entry["name"] for entry in series)
    legend_height = 15 * len(series) if legend_visible else 0
    bottom_margin = max(90, category_width * math.sin(math.radians(category_angle)) + 48) + legend_height
    left_margin = max(100, category_width + 45) if horizontal_bar else 100
    plot_width = 495 if horizontal_bar or kind == "scatterChart" else max(495, len(model["categories"]) * 27)
    plot_height = max(270, len(model["categories"]) * 18) if horizontal_bar else 270
    if label_width(model["axis_labels"]["x"]) > plot_width or label_width(model["axis_labels"]["y"]) > plot_height:
        raise UnsupportedChart("An axis title exceeds the supported layout; shorten it or export using the original chart editor.")
    if legend_visible and any(label_width(entry["name"]) > plot_width - 35 for entry in series):
        raise UnsupportedChart("A legend entry exceeds the supported layout; shorten it or export using the original chart editor.")
    title_lines = []
    remaining = model["title"]
    while remaining:
        limit = len(remaining)
        while limit > 1 and label_width(remaining[:limit], 12) > left_margin + plot_width - 10:
            limit -= 1
        if limit < len(remaining) and " " in remaining[:limit]:
            limit = remaining.rfind(" ", 0, limit) + 1
        title_lines.append(remaining[:limit].strip())
        remaining = remaining[limit:].lstrip()
    width = left_margin + plot_width + 45
    height = bottom_margin + plot_height + 40 + len(title_lines) * 16
    drawing = Drawing(width, height)
    if kind == "scatterChart":
        plot = LinePlot()
        plot.data = [entry["points"] for entry in series]
        plot.joinedLines = model["scatter_style"] in {"line", "lineMarker"}
        if model["scatter_style"] in {"marker", "lineMarker"}:
            plot.lines.symbol = makeMarker("FilledCircle")
        value_axes = [plot.xValueAxis, plot.yValueAxis]
    elif kind == "lineChart":
        plot = HorizontalLineChart()
        plot.data = [entry["y"] for entry in series]
        plot.categoryAxis.categoryNames = model["categories"]
        value_axes = [plot.valueAxis]
    else:
        plot = HorizontalBarChart() if model["bar_direction"] == "bar" else VerticalBarChart()
        plot.data = [entry["y"] for entry in series]
        plot.categoryAxis.categoryNames = model["categories"]
        value_axes = [plot.valueAxis]
        # A zero baseline is necessary to retain the visual interpretation of bar lengths.
        all_values = [value for entry in series for value in entry["y"]]
        plot.valueAxis.valueMin = min(0, min(all_values))
        if max(all_values) <= 0:
            plot.valueAxis.valueMax = 0 if min(all_values) < 0 else 1
    plot.x, plot.y, plot.width, plot.height = left_margin, bottom_margin, plot_width, plot_height
    palette = [colors.HexColor(value) for value in ("#174A6E", "#A4402C", "#236E55", "#78528A", "#805E19", "#424242")]
    for index, entry in enumerate(series):
        if kind == "barChart":
            plot.bars[index].fillColor = palette[index % len(palette)]
            plot.bars[index].strokeColor = None
        else:
            plot.lines[index].strokeColor = palette[index % len(palette)]
            plot.lines[index].strokeWidth = 1.2
            if kind == "lineChart" and entry["marker"] not in {None, "none"}:
                plot.lines[index].symbol = makeMarker("FilledCircle")
    for axis in value_axes:
        axis.labels.fontName = "Helvetica"
        axis.labels.fontSize = 10
        axis.labelTextFormat = lambda value: f"{value:g}"
    if kind != "scatterChart":
        plot.categoryAxis.labels.fontName = "Helvetica"
        plot.categoryAxis.labels.fontSize = 10
        # Negative bars put the zero axis at the top/right. Keep labels outside
        # the plot so black text is not painted over dark bars.
        plot.categoryAxis.labelAxisMode = "low"
        if model["bar_direction"] != "bar" or kind != "barChart":
            plot.categoryAxis.labels.angle = category_angle
            plot.categoryAxis.labels.boxAnchor = "ne" if plot.categoryAxis.labels.angle else "n"
    drawing.add(plot)
    for index, line in enumerate(title_lines):
        drawing.add(String(width / 2, height - 20 - index * 16, line, fontName="Helvetica", fontSize=12, textAnchor="middle"))
    for key, label in model["axis_labels"].items():
        if label:
            if key == "x":
                drawing.add(String(left_margin + plot_width / 2, legend_height + 32, label, fontName="Helvetica", fontSize=10, textAnchor="middle"))
            else:
                from reportlab.graphics.shapes import Group
                group = Group(String(0, 0, label, fontName="Helvetica", fontSize=10, textAnchor="middle"))
                group.rotate(90)
                group.shift(22, bottom_margin + plot_height / 2)
                drawing.add(group)
    if legend_visible:
        legend = Legend()
        legend.x, legend.y = left_margin, legend_height + 9
        legend.fontName, legend.fontSize = "Helvetica", 10
        legend.columnMaximum = len(series)
        legend.colorNamePairs = [(palette[index % len(palette)], entry["name"] or f"Series {index + 1}") for index, entry in enumerate(series)]
        drawing.add(legend)
    return drawing


def export_visuals(docx_path: Path, out_dir: Path) -> dict:
    """Export original assets and supported new SVGs into a NEW directory.

    Existing directories (including empty ones and symlinks) are refused. The
    returned manifest is also saved as manifest.json. Unsupported charts are
    recorded individually; no OCR, bitmap tracing, workbook evaluation, or
    native-table-to-image conversion is performed.
    """
    docx_path, out_dir = Path(docx_path), Path(out_dir)
    if out_dir.exists() or out_dir.is_symlink():
        raise FileExistsError(f"Output directory already exists: {out_dir}")
    source_bytes = docx_path.read_bytes()
    with ZipFile(docx_path) as package:
        inventory = _inspect(package)
        out_dir.mkdir(parents=True, exist_ok=False)
        manifest = {
            "schema_version": 1, "source_document": str(docx_path.resolve()), "source_sha256": _sha(source_bytes),
            "output_directory": str(out_dir.resolve()), "output_dir": str(out_dir.resolve()), "inventory": inventory, "items": [],
            "limitations": [
                "Not an APA compliance certification. APA formatting and scientific accuracy still require review.",
                "Reconstructed SVGs use complete explicit chart caches, which may be stale; workbooks/formulas are not evaluated.",
                "Reconstruction preserves cached data pairing but uses a new style and automatic axis ranges; it is not an appearance-preserving conversion.",
                "Review figures at the FINAL insertion size: SVG scaling can move labels outside APA's 8–14 pt figure-text range, regardless of the original 10-unit font settings.",
                "SVG text is retained as text, with Helvetica and local font fallback. Chinese/other Unicode glyph coverage depends on the viewing computer; inspect for missing glyphs, label overlap, and clipping.",
                "Raster images are copied unchanged, never wrapped in an SVG and called vectors; screenshot table/chart reconstruction requires source data or separate reviewed OCR.",
                "SVG/EMF/WMF are vector-capable containers and may themselves contain bitmap content; exact extraction does not certify all-vector content.",
                "Original native tables/charts and document assets are not modified. Exported figures are separate files.",
            ],
        }
        for index, media in enumerate(inventory["media"], 1):
            part = media["source_part"]
            data = package.read(part)
            item = {"source_part": part, "source_sha256": _sha(data), "kind": media["kind"]}
            if media["kind"] == "other":
                item.update(method="unsupported", reason="Unrecognized media format; no automatic conversion.")
            else:
                filename = f"media_{index:03d}_" + re.sub(r"[^A-Za-z0-9_.-]", "_", PurePosixPath(part).name)
                with (out_dir / filename).open("xb") as output:
                    output.write(data)
                item.update(method="extracted_vector" if media["kind"] == "vector_container" else "extracted_raster",
                            output=filename, output_sha256=_sha(data), bytes=len(data))
                if media["kind"] == "vector_container":
                    item["vector_purity"] = "not_verified"
                if PurePosixPath(part).suffix.lower() == ".svg":
                    try:
                        svg = _xml(data)
                        item["contains_svg_image_elements"] = bool(svg.findall(".//svg:image", NS))
                        if item["contains_svg_image_elements"]:
                            item["vector_purity"] = "contains_embedded_or_linked_images"
                    except etree.XMLSyntaxError:
                        item["warning"] = "SVG XML could not be parsed; original bytes were preserved without validation."
            manifest["items"].append(item)
        for index, native in enumerate(inventory["native_charts"], 1):
            part = native["source_part"]
            data = package.read(part)
            item = {"source_part": part, "source_sha256": _sha(data), "kind": "native_chart"}
            filename = f"chart_{index:03d}_reconstructed.svg"
            try:
                model = _model(data)
                from reportlab.graphics import renderSVG
                drawing = _drawing(model)
                svg_data = renderSVG.drawToString(drawing)
                if isinstance(svg_data, str):
                    svg_data = svg_data.encode("utf-8")
                # The renderer must produce geometry/text, never an image-only wrapper.
                svg = _xml(svg_data)
                if svg.findall(".//svg:image", NS):
                    raise UnsupportedChart("Unexpected raster element in vector renderer output.")
                with (out_dir / filename).open("xb") as output:
                    output.write(svg_data)
                item.update(method="reconstructed_from_chart_cache", output=filename,
                            output_sha256=_sha(svg_data), bytes=len(svg_data), data=model,
                            vector_purity="vector_shapes_and_text_no_image_elements",
                            review_required=True, reason="New vector chart built from cached values; original formatting is not reproduced.")
            except (UnsupportedChart, etree.XMLSyntaxError, ImportError, ValueError, TypeError, AttributeError, ZeroDivisionError) as exc:
                item.update(method="unsupported", reason=f"{type(exc).__name__}: {exc}")
            manifest["items"].append(item)
        with (out_dir / "manifest.json").open("x", encoding="utf-8") as output:
            json.dump(manifest, output, ensure_ascii=False, indent=2)
    return manifest
