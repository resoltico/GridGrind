package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.w3c.dom.Node;

/**
 * Rewrites defined-name chart formulas to explicit area references before POI sheet cloning.
 *
 * <p>POI's cloneSheet implementation rewrites chart references by assuming every chart formula is a
 * concrete cell range. Workbook or sheet defined names violate that assumption and can crash the
 * clone with invalid negative-column references before GridGrind repair passes begin. GridGrind
 * therefore normalizes chart formulas on the source sheet just long enough for the clone to
 * complete, then restores the original source formulas immediately afterward.
 */
final class ExcelSheetCloneChartPreparationSupport {
  RewrittenChartFormulas rewriteDefinedNameFormulas(XSSFSheet sourceSheet) {
    Objects.requireNonNull(sourceSheet, "sourceSheet must not be null");
    XSSFDrawing drawing = sourceSheet.getDrawingPatriarch();
    if (drawing == null) {
      return new RewrittenChartFormulas(List.of());
    }

    List<FormulaRewrite> rewrites = new ArrayList<>();
    for (XSSFChart chart : drawing.getCharts()) {
      collectFormulaRewrites(sourceSheet, chart, rewrites);
    }
    return new RewrittenChartFormulas(rewrites);
  }

  private static void collectFormulaRewrites(
      XSSFSheet contextSheet, XSSFChart chart, List<FormulaRewrite> rewrites) {
    collectFormulaRewrites(contextSheet, chart.getCTChart().getDomNode(), rewrites);
  }

  private static void collectFormulaRewrites(
      XSSFSheet contextSheet, Node node, List<FormulaRewrite> rewrites) {
    if (isFormulaNode(node)) {
      maybeRewriteDefinedNameFormula(contextSheet, node, rewrites);
    }
    for (Node child = node.getFirstChild(); child != null; child = child.getNextSibling()) {
      collectFormulaRewrites(contextSheet, child, rewrites);
    }
  }

  private static void maybeRewriteDefinedNameFormula(
      XSSFSheet contextSheet, Node formulaNode, List<FormulaRewrite> rewrites) {
    String originalFormula = formulaText(formulaNode);
    if (originalFormula == null || originalFormula.isBlank()) {
      return;
    }
    String normalizedFormula = ExcelChartSourceSupport.normalizeFormula(originalFormula);
    if (ExcelChartSourceSupport.resolveDefinedNameReference(contextSheet, normalizedFormula)
        == null) {
      return;
    }

    String rewrittenFormula = explicitAreaFormula(contextSheet, normalizedFormula);
    if (originalFormula.startsWith("=")) {
      rewrittenFormula = "=" + rewrittenFormula;
    }

    setFormulaText(formulaNode, rewrittenFormula);
    rewrites.add(new FormulaRewrite(formulaNode, originalFormula));
  }

  private static boolean isFormulaNode(Node node) {
    return node.getNodeType() == Node.ELEMENT_NODE && "f".equals(node.getLocalName());
  }

  private static String formulaText(Node formulaNode) {
    Node valueNode = formulaValueNode(formulaNode);
    return valueNode == null ? null : valueNode.getNodeValue();
  }

  private static void setFormulaText(Node formulaNode, String formula) {
    Node valueNode = formulaValueNode(formulaNode);
    if (valueNode == null) {
      throw new IllegalStateException("Chart formula node is missing its text payload");
    }
    valueNode.setNodeValue(formula);
  }

  static Node formulaValueNode(Node formulaNode) {
    for (Node child = formulaNode.getFirstChild(); child != null; child = child.getNextSibling()) {
      if (child.getNodeType() == Node.TEXT_NODE || child.getNodeType() == Node.CDATA_SECTION_NODE) {
        return child;
      }
    }
    return null;
  }

  static String explicitAreaFormula(XSSFSheet contextSheet, String formula) {
    Objects.requireNonNull(contextSheet, "contextSheet must not be null");
    String normalizedFormula = ExcelChartSourceSupport.normalizeFormula(formula);
    ResolvedAreaReference resolved =
        ExcelChartSourceSupport.resolveAreaReference(contextSheet, normalizedFormula);
    return formatResolvedAreaReference(resolved.sheet(), resolved.areaReference());
  }

  private static String formatResolvedAreaReference(XSSFSheet sheet, AreaReference areaReference) {
    AreaReference explicitReference =
        new AreaReference(
            absoluteCellReference(sheet, areaReference.getFirstCell()),
            absoluteCellReference(sheet, areaReference.getLastCell()),
            SpreadsheetVersion.EXCEL2007);
    return explicitReference.formatAsString();
  }

  private static CellReference absoluteCellReference(XSSFSheet sheet, CellReference reference) {
    return new CellReference(
        sheet.getSheetName(), reference.getRow(), reference.getCol(), true, true);
  }

  record RewrittenChartFormulas(List<FormulaRewrite> rewrites) {
    RewrittenChartFormulas {
      rewrites = List.copyOf(rewrites);
    }

    void restoreSourceFormulas() {
      for (FormulaRewrite rewrite : rewrites) {
        rewrite.restore();
      }
    }
  }

  private record FormulaRewrite(Node formulaNode, String originalFormula) {
    private FormulaRewrite {
      Objects.requireNonNull(formulaNode, "formulaNode must not be null");
      Objects.requireNonNull(originalFormula, "originalFormula must not be null");
    }

    void restore() {
      setFormulaText(formulaNode, originalFormula);
    }
  }
}
