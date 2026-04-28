package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import java.io.IOException;
import java.io.StringReader;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import javax.xml.parsers.DocumentBuilderFactory;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.jupiter.api.Test;
import org.w3c.dom.Node;
import org.xml.sax.InputSource;

/** Tests for chart clone-preparation support on named-range-backed chart formulas. */
class ExcelSheetCloneChartPreparationSupportTest {
  private final ExcelSheetCloneChartPreparationSupport support =
      new ExcelSheetCloneChartPreparationSupport();

  @Test
  void rewriteDefinedNameFormulasPreservesLeadingEqualsAndRestoresOriginalText()
      throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      XSSFSheet sourceSheet = namedRangeChartSheet(workbook);
      Node categoriesFormulaNode = formulaNode(sourceSheet, "ChartCategories");
      Node valuesFormulaNode = formulaNode(sourceSheet, "ChartPlan");
      replaceFormulaText(categoriesFormulaNode, "=ChartCategories");
      replaceFormulaText(valuesFormulaNode, "=ChartPlan");

      ExcelSheetCloneChartPreparationSupport.RewrittenChartFormulas rewrites =
          support.rewriteDefinedNameFormulas(sourceSheet);

      assertEquals(2, rewrites.rewrites().size());
      assertEquals("=Source!$A$2:$A$4", formulaText(categoriesFormulaNode));
      assertEquals("=Source!$B$2:$B$4", formulaText(valuesFormulaNode));

      rewrites.restoreSourceFormulas();

      assertEquals("=ChartCategories", formulaText(categoriesFormulaNode));
      assertEquals("=ChartPlan", formulaText(valuesFormulaNode));
    }
  }

  @Test
  void rewriteDefinedNameFormulasIgnoresBlankAndPayloadlessFormulaNodes() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      XSSFSheet sourceSheet = namedRangeChartSheet(workbook);
      Node categoriesFormulaNode = formulaNode(sourceSheet, "ChartCategories");
      Node valuesFormulaNode = formulaNode(sourceSheet, "ChartPlan");
      replaceFormulaText(categoriesFormulaNode, "   ");
      replaceChildren(valuesFormulaNode, valuesFormulaNode.getParentNode().cloneNode(false));

      ExcelSheetCloneChartPreparationSupport.RewrittenChartFormulas rewrites =
          support.rewriteDefinedNameFormulas(sourceSheet);

      assertTrue(rewrites.rewrites().isEmpty());
      assertEquals("   ", formulaText(categoriesFormulaNode));
      assertNull(formulaText(valuesFormulaNode));
    }
  }

  @Test
  void restoreSourceFormulasFailsWhenARewrittenFormulaLosesItsTextPayload() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      XSSFSheet sourceSheet = namedRangeChartSheet(workbook);
      Node categoriesFormulaNode = formulaNode(sourceSheet, "ChartCategories");

      ExcelSheetCloneChartPreparationSupport.RewrittenChartFormulas rewrites =
          support.rewriteDefinedNameFormulas(sourceSheet);
      replaceChildren(categoriesFormulaNode, null);

      IllegalStateException exception =
          assertThrows(IllegalStateException.class, rewrites::restoreSourceFormulas);
      assertEquals("Chart formula node is missing its text payload", exception.getMessage());
    }
  }

  @Test
  void formulaValueNodeRecognizesCdataPayloads() throws Exception {
    DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
    factory.setNamespaceAware(true);
    Node formulaNode =
        factory
            .newDocumentBuilder()
            .parse(new InputSource(new StringReader("<f><![CDATA[ChartCategories]]></f>")))
            .getDocumentElement();

    Optional<Node> valueNode = ExcelSheetCloneChartPreparationSupport.formulaValueNode(formulaNode);

    assertTrue(valueNode.isPresent());
    assertEquals(Node.CDATA_SECTION_NODE, valueNode.orElseThrow().getNodeType());
    assertEquals("ChartCategories", valueNode.orElseThrow().getNodeValue());
  }

  private static XSSFSheet namedRangeChartSheet(ExcelWorkbook workbook) {
    ExcelSheet source = workbook.getOrCreateSheet("Source");
    ExcelChartTestSupport.seedChartData(source);
    ExcelChartTestSupport.seedChartNamedRanges(workbook, "Source");
    source.setChart(
        ExcelChartTestSupport.barChart(
            "NamedRangeChart",
            ExcelChartTestSupport.anchor(4, 1, 10, 14),
            new ExcelChartDefinition.Title.Text("Named ranges"),
            new ExcelChartDefinition.Legend.Hidden(),
            ExcelChartDisplayBlanksAs.GAP,
            true,
            false,
            ExcelChartBarDirection.COLUMN,
            List.of(
                new ExcelChartDefinition.Series(
                    new ExcelChartDefinition.Title.Text("Plan"),
                    ExcelChartTestSupport.ref("ChartCategories"),
                    ExcelChartTestSupport.ref("ChartPlan")))));
    return source.xssfSheet();
  }

  private static Node formulaNode(XSSFSheet sheet, String originalFormula) {
    return formulaNodes(sheet.getDrawingPatriarch().getCharts().getFirst()).stream()
        .filter(node -> originalFormula.equals(formulaText(node)))
        .findFirst()
        .orElseThrow(() -> new AssertionError("formula node not found: " + originalFormula));
  }

  private static List<Node> formulaNodes(XSSFChart chart) {
    List<Node> formulaNodes = new ArrayList<>();
    collectFormulaNodes(chart.getCTChart().getDomNode(), formulaNodes);
    return formulaNodes;
  }

  private static void collectFormulaNodes(Node node, List<Node> formulaNodes) {
    if (node.getNodeType() == Node.ELEMENT_NODE && "f".equals(node.getLocalName())) {
      formulaNodes.add(node);
    }
    for (Node child = node.getFirstChild(); child != null; child = child.getNextSibling()) {
      collectFormulaNodes(child, formulaNodes);
    }
  }

  private static String formulaText(Node formulaNode) {
    for (Node child = formulaNode.getFirstChild(); child != null; child = child.getNextSibling()) {
      if (child.getNodeType() == Node.TEXT_NODE || child.getNodeType() == Node.CDATA_SECTION_NODE) {
        return child.getNodeValue();
      }
    }
    return null;
  }

  private static void replaceFormulaText(Node formulaNode, String formula) {
    formulaValueNode(formulaNode)
        .orElseThrow(() -> new AssertionError("formula node is missing its text payload"))
        .setNodeValue(formula);
  }

  private static void replaceChildren(Node node, Node replacementChild) {
    while (node.getFirstChild() != null) {
      node.removeChild(node.getFirstChild());
    }
    if (replacementChild != null) {
      node.appendChild(replacementChild);
    }
  }

  private static Optional<Node> formulaValueNode(Node formulaNode) {
    return ExcelSheetCloneChartPreparationSupport.formulaValueNode(formulaNode);
  }
}
