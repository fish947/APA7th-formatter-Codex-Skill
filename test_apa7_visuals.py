"""Synthetic OOXML tests: vector preservation, data pairing and refusal paths."""
from __future__ import annotations

import copy
import json
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile

from lxml import etree

import apa7_visuals as visuals


def cache(tag, values, numeric=True, order=None):
    container = etree.Element("{" + visuals.NS["c"] + "}" + tag)
    kind = "num" if numeric else "str"
    reference = etree.SubElement(container, "{" + visuals.NS["c"] + "}" + kind + "Ref")
    etree.SubElement(reference, "{" + visuals.NS["c"] + "}f").text = "Sheet1!$A$2:$A$4"
    stored = etree.SubElement(reference, "{" + visuals.NS["c"] + "}" + kind + "Cache")
    etree.SubElement(stored, "{" + visuals.NS["c"] + "}ptCount", val=str(len(values)))
    for index in (range(len(values)) if order is None else order):
        point = etree.SubElement(stored, "{" + visuals.NS["c"] + "}pt", idx=str(index))
        etree.SubElement(point, "{" + visuals.NS["c"] + "}v").text = str(values[index])
    return container


def chart(kind="scatterChart", x=(3, 1, 2), y=(30, 10, 20), order=None):
    C = "{" + visuals.NS["c"] + "}"
    A = "{" + visuals.NS["a"] + "}"
    root = etree.Element(C + "chartSpace", nsmap={"c": visuals.NS["c"], "a": visuals.NS["a"]})
    parent = etree.SubElement(root, C + "chart")
    plot = etree.SubElement(parent, C + "plotArea")
    graph = etree.SubElement(plot, C + kind)
    if kind == "scatterChart":
        etree.SubElement(graph, C + "scatterStyle", val="lineMarker")
    else:
        etree.SubElement(graph, C + "grouping", val="clustered" if kind == "barChart" else "standard")
        if kind == "barChart":
            etree.SubElement(graph, C + "barDir", val="col")
    series = etree.SubElement(graph, C + "ser")
    etree.SubElement(series, C + "order", val="0")
    name = etree.SubElement(series, C + "tx")
    etree.SubElement(name, C + "v").text = "Measured values"
    series.append(cache("xVal" if kind == "scatterChart" else "cat", x, kind == "scatterChart", order))
    series.append(cache("yVal" if kind == "scatterChart" else "val", y, True, tuple(reversed(range(len(y))))))
    etree.SubElement(graph, C + "axId", val="1")
    etree.SubElement(graph, C + "axId", val="2")
    for axis_id, position, cross, label in (("1", "b", "2", "Condition / X"), ("2", "l", "1", "Measured response")):
        axis = etree.SubElement(plot, C + ("catAx" if kind != "scatterChart" and axis_id == "1" else "valAx"))
        etree.SubElement(axis, C + "axId", val=axis_id)
        etree.SubElement(axis, C + "axPos", val=position)
        etree.SubElement(axis, C + "crossAx", val=cross)
        title = etree.SubElement(axis, C + "title")
        tx = etree.SubElement(title, C + "tx")
        rich = etree.SubElement(tx, C + "rich")
        paragraph = etree.SubElement(rich, A + "p")
        run = etree.SubElement(paragraph, A + "r")
        etree.SubElement(run, A + "t").text = label
    etree.SubElement(parent, C + "legend")
    return root


class VisualExportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="apa7_visuals_")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def package(self, parts):
        source = self.root / "source.docx"
        with ZipFile(source, "w") as archive:
            for name, data in parts.items():
                archive.writestr(name, data if isinstance(data, bytes) else etree.tostring(data))
        return source

    def test_native_vector_bytes_and_raster_label_preserved(self):
        svg = b'<svg xmlns="http://www.w3.org/2000/svg"><path d="M0 0L5 10"/></svg>'
        mixed_svg = b'<svg xmlns="http://www.w3.org/2000/svg"><image href="data:image/png;base64,AA=="/></svg>'
        png = b"\x89PNG\r\n\x1a\nfixture-bytes-preserved"
        source = self.package({"word/media/pure.svg": svg, "word/media/mixed.svg": mixed_svg,
                               "word/media/photo.png": png, "word/media/drawing.emf": b"emf-fixture"})
        before = source.read_bytes()
        destination = self.root / "vectors"
        result = visuals.export_visuals(source, destination)
        self.assertEqual(source.read_bytes(), before)
        by_part = {item["source_part"]: item for item in result["items"]}
        for part, expected in (("word/media/pure.svg", svg), ("word/media/mixed.svg", mixed_svg), ("word/media/photo.png", png)):
            self.assertEqual((destination / by_part[part]["output"]).read_bytes(), expected)
        self.assertEqual(by_part["word/media/photo.png"]["method"], "extracted_raster")
        self.assertTrue(by_part["word/media/photo.png"]["output"].endswith(".png"))
        self.assertEqual(by_part["word/media/pure.svg"]["method"], "extracted_vector")
        self.assertFalse(by_part["word/media/pure.svg"]["contains_svg_image_elements"])
        self.assertTrue(by_part["word/media/mixed.svg"]["contains_svg_image_elements"])
        self.assertEqual(json.loads((destination / "manifest.json").read_text())["source_sha256"], result["source_sha256"])

    def test_scatter_index_pairing_not_xml_order_or_sorted_x(self):
        source = self.package({"word/charts/chart1.xml": chart(order=(2, 0, 1))})
        result = visuals.export_visuals(source, self.root / "out")
        item = result["items"][0]
        self.assertEqual(item["method"], "reconstructed_from_chart_cache", item)
        self.assertEqual(item["data"]["series"][0]["points"], [(3.0, 30.0), (1.0, 10.0), (2.0, 20.0)])
        drawing = visuals._drawing(item["data"])
        self.assertEqual(drawing.contents[0].data[0], [(3.0, 30.0), (1.0, 10.0), (2.0, 20.0)])
        svg = etree.fromstring((self.root / "out" / item["output"]).read_bytes())
        self.assertFalse(svg.findall(".//svg:image", visuals.NS))
        self.assertTrue(svg.findall(".//svg:polyline", visuals.NS) or svg.findall(".//svg:path", visuals.NS))

    def test_categorical_line_and_bar_keep_paired_labels(self):
        for kind in ("lineChart", "barChart"):
            with self.subTest(kind=kind):
                source = self.package({"word/charts/chart1.xml": chart(kind, ("Third", "First", "Second"), order=(1, 2, 0))})
                result = visuals.export_visuals(source, self.root / kind)
                item = result["items"][0]
                self.assertEqual(item["method"], "reconstructed_from_chart_cache", item)
                self.assertEqual(item["data"]["series"][0]["points"], [("Third", 30.0), ("First", 10.0), ("Second", 20.0)])
                drawing = visuals._drawing(item["data"])
                self.assertEqual(drawing.contents[0].categoryAxis.categoryNames, ["Third", "First", "Second"])
                self.assertEqual(drawing.contents[0].data[0], [30.0, 10.0, 20.0])

    def test_negative_bar_zero_baseline_and_labels_outside_dark_bars(self):
        model = visuals._model(etree.tostring(chart("barChart")))
        model["series"][0]["y"] = [-2.0, -5.0, -3.0]
        model["series"][0]["points"] = list(zip(model["categories"], model["series"][0]["y"]))
        plot = visuals._drawing(model).contents[0]
        self.assertEqual(plot.valueAxis.valueMax, 0)
        self.assertEqual(plot.valueAxis.valueMin, -5)
        self.assertEqual(plot.categoryAxis.labelAxisMode, "low")

    def test_refuses_incomplete_and_unsupported_charts(self):
        C = "{" + visuals.NS["c"] + "}"
        invalid = {}
        model = chart()
        missing = model.find(".//c:yVal/c:numRef/c:numCache/c:pt", visuals.NS)
        missing.getparent().remove(missing)
        invalid["gap"] = model
        model = chart()
        point = model.find(".//c:xVal/c:numRef/c:numCache/c:pt", visuals.NS)
        point.getparent().append(copy.deepcopy(point))
        invalid["duplicate"] = model
        model = chart()
        etree.SubElement(model.find(".//c:ser", visuals.NS), C + "smooth", val="1")
        invalid["smooth"] = model
        for name, tag in (("errorbars", "errBars"), ("trendline", "trendline")):
            model = chart()
            etree.SubElement(model.find(".//c:ser", visuals.NS), C + tag)
            invalid[name] = model
        model = chart()
        scaling = etree.SubElement(model.find(".//c:valAx", visuals.NS), C + "scaling")
        etree.SubElement(scaling, C + "logBase", val="10")
        invalid["log"] = model
        model = chart()
        etree.SubElement(model.find(".//c:plotArea", visuals.NS), C + "barChart")
        invalid["combined"] = model
        model = chart()
        etree.SubElement(model.find(".//c:plotArea", visuals.NS), C + "valAx")
        invalid["multiple_axes"] = model
        model = chart("barChart", ("A", "B", "C"))
        model.find(".//c:grouping", visuals.NS).set("val", "stacked")
        invalid["stacked"] = model
        model = chart("lineChart", ("A", "B", "C"))
        model.find(".//c:catAx", visuals.NS).tag = C + "dateAx"
        invalid["date"] = model
        for name, model in invalid.items():
            with self.subTest(name=name):
                source = self.package({"word/charts/chart1.xml": model})
                result = visuals.export_visuals(source, self.root / name)
                self.assertEqual(result["items"][0]["method"], "unsupported")
                self.assertTrue(result["items"][0]["reason"])
                self.assertFalse(list((self.root / name).glob("*.svg")))

    def test_existing_directory_and_source_never_overwritten(self):
        source = self.package({"word/charts/chart1.xml": chart()})
        before = source.read_bytes()
        destination = self.root / "out"
        destination.mkdir()
        sentinel = destination / "manifest.json"
        sentinel.write_bytes(b"keep existing output")
        with self.assertRaises(FileExistsError):
            visuals.export_visuals(source, destination)
        self.assertEqual(sentinel.read_bytes(), b"keep existing output")
        self.assertEqual(source.read_bytes(), before)

    def test_inventory_native_tables_inline_floating_and_chart_relationships(self):
        document = etree.fromstring(('''<w:document xmlns:w="%(w)s" xmlns:wp="%(wp)s" xmlns:a="%(a)s" xmlns:c="%(c)s" xmlns:r="%(r)s">
          <w:body><w:tbl><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:tbl><w:tr><w:tc/></w:tr></w:tbl></w:tc></w:tr></w:tbl>
          <w:p><w:r><w:drawing><wp:inline><a:blip r:embed="rImage"/></wp:inline></w:drawing></w:r></w:p>
          <w:p><w:r><w:drawing><wp:anchor><a:blip r:embed="rImage"/></wp:anchor></w:drawing></w:r></w:p>
          <w:p><w:r><w:drawing><wp:inline><c:chart r:id="rChart"/></wp:inline></w:drawing></w:r></w:p>
          </w:body></w:document>''' % visuals.NS).encode())
        rels = ('''<Relationships xmlns="%(rel)s"><Relationship Id="rImage" Target="media/image1.png" Type="%(r)s/image"/><Relationship Id="rChart" Target="charts/chart1.xml" Type="%(r)s/chart"/></Relationships>''' % visuals.NS).encode()
        source = self.package({"word/document.xml": document, "word/_rels/document.xml.rels": rels,
                               "word/media/image1.png": b"raster", "word/charts/chart1.xml": chart()})
        result = visuals.inspect_visuals(source)
        self.assertEqual(result["counts"]["native_tables"], 2)
        self.assertEqual(result["counts"]["nested_tables"], 1)
        self.assertEqual(result["counts"]["top_level_tables"], 1)
        self.assertEqual(result["counts"]["inline_images"], 1)
        self.assertEqual(result["counts"]["floating_images"], 1)
        self.assertEqual(result["counts"]["native_charts"], 1)
        self.assertEqual(result["native_charts"][0]["placements"][0]["source_part"], "word/document.xml")


if __name__ == "__main__":
    unittest.main()
