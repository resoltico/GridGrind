package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchorSupport;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingBinarySupport;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.drawing.ExcelShapeDefinition;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlHashAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureDigestAlgorithm;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlEncryptionSnapshot;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlSecurityPoiBridge;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableAnalysisSupport;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableSnapshotSupport;
import dev.erst.gridgrind.excel.pivot.PivotHandle;
import java.util.List;
import java.util.Optional;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.poifs.crypt.ChainingMode;
import org.apache.poi.poifs.crypt.CipherAlgorithm;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.HashAlgorithm;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.usermodel.XSSFPivotCacheDefinition;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlObject;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDefinedName;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSortCondition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSortBy;

/** Residual direct utility coverage for workbook engine helpers left outside broader workflows. */
class ResidualUtilityCoverageTest extends ExcelPivotTableCoverageTestSupport {
  @Test
  void colorFactoriesShapeKindsAndConditionalFormattingBranchesRemainExplicit() {
    assertEquals(Optional.of(0.25d), ExcelColor.rgb("#112233", 0.25d).tint());
    assertEquals(Optional.of(0.25d), ExcelColorSnapshot.rgb("#112233", 0.25d).tint());
    assertEquals(Optional.of(-0.5d), ExcelColorSnapshot.indexed(7, -0.5d).tint());

    ExcelDrawingAnchor.TwoCell anchor =
        new ExcelDrawingAnchor.TwoCell(
            new ExcelDrawingMarker(0, 0), new ExcelDrawingMarker(2, 2), null);
    ExcelShapeDefinition.SimpleShape simpleShape =
        new ExcelShapeDefinition.SimpleShape("Queue Banner", anchor, "rect", Optional.of("Ready"));
    assertEquals(ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, simpleShape.kind());
    assertEquals(
        "name must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new ExcelShapeDefinition.Connector(" ", anchor))
            .getMessage());
    assertEquals(
        "rank must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExcelConditionalFormattingRule.Top10Rule(
                        0, false, false, false, Optional.empty()))
            .getMessage());
    assertEquals(
        "mode must not be absent when encrypted",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ExcelOoxmlEncryptionSnapshot(
                        true,
                        Optional.empty(),
                        Optional.of(ExcelOoxmlCipherAlgorithm.AES_256),
                        Optional.of(ExcelOoxmlHashAlgorithm.SHA_512),
                        Optional.of(ExcelOoxmlChainingMode.CBC),
                        Optional.of(256),
                        Optional.of(16),
                        Optional.of(1)))
            .getMessage());
  }

  @Test
  void poiBridgeAndSnapshotHelpersCoverResidualEnumAndErrorPaths() throws Exception {
    assertEquals(
        ExcelOoxmlEncryptionMode.AGILE, ExcelOoxmlSecurityPoiBridge.fromPoi(EncryptionMode.agile));
    assertEquals(
        ExcelOoxmlEncryptionMode.STANDARD,
        ExcelOoxmlSecurityPoiBridge.fromPoi(EncryptionMode.standard));
    assertEquals(
        EncryptionMode.agile, ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlEncryptionMode.AGILE));
    assertEquals(
        HashAlgorithm.sha256,
        ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlSignatureDigestAlgorithm.SHA256));
    assertEquals(
        HashAlgorithm.sha384,
        ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlSignatureDigestAlgorithm.SHA384));
    assertEquals(
        HashAlgorithm.sha512,
        ExcelOoxmlSecurityPoiBridge.toPoi(ExcelOoxmlSignatureDigestAlgorithm.SHA512));

    assertEquals(
        ExcelOoxmlCipherAlgorithm.RC4, ExcelOoxmlSecurityPoiBridge.fromPoi(CipherAlgorithm.rc4));
    assertEquals(
        ExcelOoxmlCipherAlgorithm.AES_128,
        ExcelOoxmlSecurityPoiBridge.fromPoi(CipherAlgorithm.aes128));
    assertEquals(
        ExcelOoxmlCipherAlgorithm.AES_192,
        ExcelOoxmlSecurityPoiBridge.fromPoi(CipherAlgorithm.aes192));
    assertEquals(
        ExcelOoxmlCipherAlgorithm.AES_256,
        ExcelOoxmlSecurityPoiBridge.fromPoi(CipherAlgorithm.aes256));
    assertEquals(
        ExcelOoxmlCipherAlgorithm.RC2, ExcelOoxmlSecurityPoiBridge.fromPoi(CipherAlgorithm.rc2));
    assertEquals(
        ExcelOoxmlCipherAlgorithm.DES, ExcelOoxmlSecurityPoiBridge.fromPoi(CipherAlgorithm.des));
    assertEquals(
        ExcelOoxmlCipherAlgorithm.TRIPLE_DES,
        ExcelOoxmlSecurityPoiBridge.fromPoi(CipherAlgorithm.des3));
    assertEquals(
        ExcelOoxmlCipherAlgorithm.TRIPLE_DES_112,
        ExcelOoxmlSecurityPoiBridge.fromPoi(CipherAlgorithm.des3_112));
    assertEquals(
        ExcelOoxmlCipherAlgorithm.RSA, ExcelOoxmlSecurityPoiBridge.fromPoi(CipherAlgorithm.rsa));

    assertEquals(
        ExcelOoxmlHashAlgorithm.NONE, ExcelOoxmlSecurityPoiBridge.fromPoi(HashAlgorithm.none));
    assertEquals(
        ExcelOoxmlHashAlgorithm.SHA_1, ExcelOoxmlSecurityPoiBridge.fromPoi(HashAlgorithm.sha1));
    assertEquals(
        ExcelOoxmlHashAlgorithm.SHA_224, ExcelOoxmlSecurityPoiBridge.fromPoi(HashAlgorithm.sha224));
    assertEquals(
        ExcelOoxmlHashAlgorithm.SHA_256, ExcelOoxmlSecurityPoiBridge.fromPoi(HashAlgorithm.sha256));
    assertEquals(
        ExcelOoxmlHashAlgorithm.SHA_384, ExcelOoxmlSecurityPoiBridge.fromPoi(HashAlgorithm.sha384));
    assertEquals(
        ExcelOoxmlHashAlgorithm.SHA_512, ExcelOoxmlSecurityPoiBridge.fromPoi(HashAlgorithm.sha512));
    assertEquals(
        ExcelOoxmlHashAlgorithm.MD2, ExcelOoxmlSecurityPoiBridge.fromPoi(HashAlgorithm.md2));
    assertEquals(
        ExcelOoxmlHashAlgorithm.MD4, ExcelOoxmlSecurityPoiBridge.fromPoi(HashAlgorithm.md4));
    assertEquals(
        ExcelOoxmlHashAlgorithm.MD5, ExcelOoxmlSecurityPoiBridge.fromPoi(HashAlgorithm.md5));
    assertEquals(
        ExcelOoxmlHashAlgorithm.RIPEMD_128,
        ExcelOoxmlSecurityPoiBridge.fromPoi(HashAlgorithm.ripemd128));
    assertEquals(
        ExcelOoxmlHashAlgorithm.RIPEMD_160,
        ExcelOoxmlSecurityPoiBridge.fromPoi(HashAlgorithm.ripemd160));
    assertEquals(
        ExcelOoxmlHashAlgorithm.RIPEMD_256,
        ExcelOoxmlSecurityPoiBridge.fromPoi(HashAlgorithm.ripemd256));
    assertEquals(
        ExcelOoxmlHashAlgorithm.WHIRLPOOL,
        ExcelOoxmlSecurityPoiBridge.fromPoi(HashAlgorithm.whirlpool));

    assertEquals(ExcelOoxmlChainingMode.ECB, ExcelOoxmlSecurityPoiBridge.fromPoi(ChainingMode.ecb));
    assertEquals(ExcelOoxmlChainingMode.CBC, ExcelOoxmlSecurityPoiBridge.fromPoi(ChainingMode.cbc));
    assertEquals(ExcelOoxmlChainingMode.CFB, ExcelOoxmlSecurityPoiBridge.fromPoi(ChainingMode.cfb));

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      CTColor automatic = CTColor.Factory.newInstance();
      automatic.setAuto(true);
      IllegalStateException automaticFailure =
          assertThrows(
              IllegalStateException.class,
              () -> ExcelColorSnapshotSupport.snapshot(workbook, automatic));
      assertTrue(automaticFailure.getMessage().contains("automatic color"));
    }
  }

  @Test
  void utilityFallbacksCoverAnchorPreviewCommentAndChartEdges() throws Exception {
    assertEquals(Optional.empty(), ExcelDrawingAnchorSupport.parentAnchor((XmlObject) null));
    assertEquals(
        Optional.empty(), ExcelDrawingAnchorSupport.parentAnchor(XmlObject.Factory.newInstance()));

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var packagePart =
          workbook
              .getPackagePart()
              .getPackage()
              .createPart(
                  PackagingURIHelper.createPartName("/xl/media/preview.bin"),
                  "application/octet-stream");
      assertFalse(ExcelDrawingBinarySupport.supportedPreviewImagePart(null));
      assertFalse(ExcelDrawingBinarySupport.supportedPreviewImagePart(packagePart));

      var sheet = workbook.createSheet("Budget");
      CreationHelper creationHelper = workbook.getCreationHelper();
      var anchor = creationHelper.createClientAnchor();
      anchor.setRow1(0);
      anchor.setRow2(2);
      anchor.setCol1(0);
      anchor.setCol2(2);
      var commentCell = sheet.createRow(0).createCell(0);
      var comment = sheet.createDrawingPatriarch().createCellComment(anchor);
      comment.setString(creationHelper.createRichTextString("Note"));
      commentCell.setCellComment(comment);
      new ExcelSheetCommentRepairSupport(sheet).replaceComments(List.of());
      assertTrue(sheet.getCellComments().isEmpty());

      var emptyCommentsSheet = workbook.createSheet("EmptyComments");
      var emptyCommentCell = emptyCommentsSheet.createRow(0).createCell(0);
      var emptyComment = emptyCommentsSheet.createDrawingPatriarch().createCellComment(anchor);
      emptyComment.setString(creationHelper.createRichTextString("Temp"));
      emptyCommentCell.setCellComment(emptyComment);
      CommentsTable emptyCommentsTable =
          emptyCommentsSheet.getRelations().stream()
              .filter(CommentsTable.class::isInstance)
              .map(CommentsTable.class::cast)
              .findFirst()
              .orElseThrow();
      emptyCommentsTable.removeComment(emptyCommentCell.getAddress());
      new ExcelSheetCommentRepairSupport(emptyCommentsSheet).replaceComments(List.of());

      CTSortCondition icon = CTSortCondition.Factory.newInstance();
      icon.setRef("A1:A4");
      icon.setSortBy(STSortBy.ICON);
      icon.setIconId(3L);
      assertInstanceOf(
          ExcelAutofilterSortConditionSnapshot.Icon.class,
          ExcelAutofilterOoxmlSupport.sortConditionSnapshot(workbook, icon));

      var printSheet = workbook.createSheet("Print");
      int sheetIndex = workbook.getSheetIndex(printSheet);
      assertEquals(
          Optional.empty(), ExcelPrintLayoutController.rawPrintAreaFormula(workbook, sheetIndex));
      workbook.getCTWorkbook().addNewDefinedNames();
      CTDefinedName blankPrintArea = workbook.getCTWorkbook().getDefinedNames().addNewDefinedName();
      blankPrintArea.setName("_xlnm.Print_Area");
      blankPrintArea.setLocalSheetId(sheetIndex);
      blankPrintArea.setStringValue(" ");
      assertEquals(
          Optional.empty(), ExcelPrintLayoutController.rawPrintAreaFormula(workbook, sheetIndex));

      CTDefinedName invalidPrintArea =
          workbook.getCTWorkbook().getDefinedNames().addNewDefinedName();
      invalidPrintArea.setName("_xlnm.Print_Area");
      invalidPrintArea.setLocalSheetId(sheetIndex);
      invalidPrintArea.setStringValue("Broken");
      assertEquals(Optional.empty(), ExcelPrintLayoutController.storedPrintAreaFormula(printSheet));
      var chartSheet = workbook.createSheet("Charts");
      assertEquals(
          List.of("cached-a", "cached-b"),
          ExcelChartSnapshotSupport.resolvedOrCachedReferenceValues(
              chartSheet,
              "[broken",
              XDDFDataSourcesFactory.fromArray(new String[] {"cached-a", "cached-b"})));
      assertEquals(
          List.of("cached-null"),
          ExcelChartSeriesSnapshotSupport.resolvedOrCachedReferenceValues(
              chartSheet,
              null,
              XDDFDataSourcesFactory.fromArray(new String[] {"cached-null"}),
              null));
      assertEquals(
          List.of("cached-blank"),
          ExcelChartSeriesSnapshotSupport.resolvedOrCachedReferenceValues(
              chartSheet,
              " ",
              XDDFDataSourcesFactory.fromArray(new String[] {"cached-blank"}),
              null));
      assertEquals(
          List.of("cached-no-context"),
          ExcelChartSeriesSnapshotSupport.resolvedOrCachedReferenceValues(
              null,
              "Sheet1!$A$1:$A$2",
              XDDFDataSourcesFactory.fromArray(new String[] {"cached-no-context"}),
              null));
    }
  }

  @Test
  void utilityFallbacksCoverPrintAreaAndPivotRelationEdges() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var printSheet = workbook.createSheet("Print");
      int sheetIndex = workbook.getSheetIndex(printSheet);
      workbook.setPrintArea(sheetIndex, "$A$1:$B$2");
      assertEquals(
          Optional.of(workbook.getPrintArea(sheetIndex)),
          ExcelPrintLayoutController.rawPrintAreaFormula(workbook, sheetIndex));
      assertEquals(
          Optional.of(workbook.getPrintArea(sheetIndex)),
          ExcelPrintLayoutController.storedPrintAreaFormula(printSheet));
    }

    assertEquals(
        List.of(), ExcelPivotTableSnapshotSupport.snapshotPageFields(null, List.of("Region")));
    assertTrue(
        ExcelPivotTableSnapshotSupport.firstRelation(
                new NullDocumentPart(), XSSFPivotCacheDefinition.class)
            .isEmpty());
  }

  @Test
  void utilityFallbacksCoverPivotHealthEdges() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Report");
      PivotHandle missingCacheHandle =
          new PivotHandle(
              workbook.xssfWorkbook().getSheetIndex("Report"),
              0,
              "Report",
              workbook.xssfWorkbook().getSheet("Report"),
              new CacheOnlyPivotTable(pivotTableDefinition("Missing Cache", "C5:F9", 11L), null));
      List<AnalysisFindingCode> missingCacheCodes =
          ExcelPivotTableAnalysisSupport.pivotTableHealthFindings(
                  workbook.xssfWorkbook(), missingCacheHandle)
              .stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList();
      assertTrue(
          missingCacheCodes.contains(AnalysisFindingCode.PIVOT_TABLE_MISSING_CACHE_DEFINITION));
      assertEquals(
          "Pivot table is missing its cache definition relation.",
          assertThrows(
                  IllegalArgumentException.class,
                  () ->
                      ExcelPivotTableSnapshotSupport.requiredCacheDefinition(
                          missingCacheHandle.table()))
              .getMessage());

      PivotHandle missingRecordsHandle =
          new PivotHandle(
              workbook.xssfWorkbook().getSheetIndex("Report"),
              1,
              "Report",
              workbook.xssfWorkbook().getSheet("Report"),
              new SyntheticPivotTable(
                  pivotTableDefinition("Missing Records", "C10:F14", 12L),
                  new XSSFPivotCacheDefinition()));
      List<AnalysisFindingCode> missingRecordsCodes =
          ExcelPivotTableAnalysisSupport.pivotTableHealthFindings(
                  workbook.xssfWorkbook(), missingRecordsHandle)
              .stream()
              .map(WorkbookAnalysis.AnalysisFinding::code)
              .toList();
      assertTrue(
          missingRecordsCodes.contains(AnalysisFindingCode.PIVOT_TABLE_MISSING_CACHE_RECORDS));
    }
  }
}
