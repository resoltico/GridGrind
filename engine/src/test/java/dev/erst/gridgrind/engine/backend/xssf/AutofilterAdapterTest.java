package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnSpec;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateSpec;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.ExcelAutofilterSortMethod;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.engine.autofilter.AutofilterAdapter;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;
import java.util.Optional;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDxf;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDynamicFilterType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STFilterOperator;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STIconSetType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSortBy;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSortMethod;

/** Tests for direct sheet-autofilter authoring, introspection, and health analysis. */
class AutofilterAdapterTest {
  private final AutofilterAdapter controller = new AutofilterAdapter();

  @Test
  void setSheetAutofilter_roundTripsOwnedMetadata() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);

      controller.setSheetAutofilter(sheet, "A1:B3");

      assertEquals(
          List.of(
              new AutofilterEntryReport.SheetOwned(
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"))),
          controller.sheetOwnedAutofilters(sheet));
      assertEquals(1, controller.sheetAutofilterCount(sheet));

      ByteArrayOutputStream output = new ByteArrayOutputStream();
      workbook.write(output);
      try (XSSFWorkbook reopened =
          new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()))) {
        XSSFSheet reopenedSheet = reopened.getSheet("Ops");
        assertEquals(
            List.of(
                new AutofilterEntryReport.SheetOwned(
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"))),
            controller.sheetOwnedAutofilters(reopenedSheet));
      }
    }
  }

  @Test
  void setSheetAutofilter_writesAdvancedCriteriaAndSortStateDetails() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedAutofilterWriteSheet(workbook);

      controller.setSheetAutofilter(
          sheet,
          "A1:F3",
          List.of(
              new AutofilterFilterColumnSpec(
                  0L, false, new AutofilterFilterCriterionSpec.Values(List.of("Ada"), true)),
              new AutofilterFilterColumnSpec(
                  1L,
                  true,
                  new AutofilterFilterCriterionSpec.Custom(
                      true,
                      List.of(
                          new AutofilterFilterCriterionSpec.CustomConditionSpec("greaterThan", "1"),
                          new AutofilterFilterCriterionSpec.CustomConditionSpec(
                              "equal", "Queue")))),
              new AutofilterFilterColumnSpec(
                  2L,
                  true,
                  new AutofilterFilterCriterionSpec.Dynamic(
                      "today", Optional.of(1.0d), Optional.of(2.0d))),
              new AutofilterFilterColumnSpec(
                  3L, true, new AutofilterFilterCriterionSpec.Top10(10, false, true)),
              new AutofilterFilterColumnSpec(
                  4L,
                  true,
                  new AutofilterFilterCriterionSpec.Color(
                      false, ColorSpec.rgb(new RgbHex("#AABBCC")))),
              new AutofilterFilterColumnSpec(
                  5L, true, new AutofilterFilterCriterionSpec.Icon("3TrafficLights1", 2))),
          Optional.of(
              new AutofilterSortStateSpec(
                  new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:F3"),
                  true,
                  true,
                  Optional.of(ExcelAutofilterSortMethod.STROKE),
                  List.of(
                      new AutofilterSortConditionSpec.CellColor(
                          new dev.erst.gridgrind.core.vocabulary.RangeRef("A2:A3"),
                          true,
                          ColorSpec.rgb(new RgbHex("#102030"))),
                      new AutofilterSortConditionSpec.Icon(
                          new dev.erst.gridgrind.core.vocabulary.RangeRef("B2:B3"), false, 4)))));

      AutofilterEntryReport.SheetOwned snapshot =
          assertInstanceOf(
              AutofilterEntryReport.SheetOwned.class,
              controller.sheetOwnedAutofilters(sheet).getFirst());

      assertEquals("A1:F3", snapshot.range().value());
      assertEquals(
          new AutofilterFilterCriterionReport.Values(List.of("Ada"), true),
          snapshot.filterColumns().get(0).criterion());
      assertEquals(
          new AutofilterFilterCriterionReport.Custom(
              true,
              List.of(
                  new AutofilterFilterCriterionReport.CustomConditionReport("greaterThan", "1"),
                  new AutofilterFilterCriterionReport.CustomConditionReport("equal", "Queue"))),
          snapshot.filterColumns().get(1).criterion());
      assertEquals(
          new AutofilterFilterCriterionReport.Dynamic(
              "today", Optional.of(1.0d), Optional.of(2.0d)),
          snapshot.filterColumns().get(2).criterion());
      assertEquals(
          new AutofilterFilterCriterionReport.Top10(false, true, 10.0d, Optional.empty()),
          snapshot.filterColumns().get(3).criterion());
      assertEquals(
          new AutofilterFilterCriterionReport.Color(
              false, Optional.ofNullable(CellColorReport.rgb(new RgbHex("#AABBCC")))),
          snapshot.filterColumns().get(4).criterion());
      assertEquals(
          new AutofilterFilterCriterionReport.Icon("3TrafficLights1", 2),
          snapshot.filterColumns().get(5).criterion());
      assertEquals(
          new AutofilterSortConditionReport.CellColor(
              new dev.erst.gridgrind.core.vocabulary.RangeLabel("A2:A3"),
              true,
              CellColorReport.rgb(new RgbHex("#102030"))),
          snapshot.sortState().orElseThrow().conditions().get(0));
      assertEquals(
          new AutofilterSortConditionReport.Icon(
              new dev.erst.gridgrind.core.vocabulary.RangeLabel("B2:B3"), false, 4),
          snapshot.sortState().orElseThrow().conditions().get(1));
    }
  }

  @Test
  void setSheetAutofilter_rejectsUnsupportedCriterionTokens() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedAutofilterWriteSheet(workbook);

      assertEquals(
          "unsupported autofilter custom operator: nope",
          assertThrows(
                  IllegalArgumentException.class,
                  () ->
                      controller.setSheetAutofilter(
                          sheet,
                          "A1:F3",
                          List.of(
                              new AutofilterFilterColumnSpec(
                                  0L,
                                  true,
                                  new AutofilterFilterCriterionSpec.Custom(
                                      false,
                                      List.of(
                                          new AutofilterFilterCriterionSpec.CustomConditionSpec(
                                              "nope", "1"))))),
                          Optional.empty()))
              .getMessage());
      assertEquals(
          "unsupported autofilter dynamic type: tomorrowish",
          assertThrows(
                  IllegalArgumentException.class,
                  () ->
                      controller.setSheetAutofilter(
                          sheet,
                          "A1:F3",
                          List.of(
                              new AutofilterFilterColumnSpec(
                                  0L,
                                  true,
                                  new AutofilterFilterCriterionSpec.Dynamic(
                                      "tomorrowish", Optional.empty(), Optional.empty()))),
                          Optional.empty()))
              .getMessage());
      assertEquals(
          "unsupported autofilter icon set: unknown",
          assertThrows(
                  IllegalArgumentException.class,
                  () ->
                      controller.setSheetAutofilter(
                          sheet,
                          "A1:F3",
                          List.of(
                              new AutofilterFilterColumnSpec(
                                  0L, true, new AutofilterFilterCriterionSpec.Icon("unknown", 0))),
                          Optional.empty()))
              .getMessage());
    }
  }

  @Test
  void clearSheetAutofilter_removesFilterDatabaseNames() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      controller.setSheetAutofilter(sheet, "A1:B3");

      assertTrue(
          sheet.getWorkbook().getAllNames().stream()
              .anyMatch(AutofilterAdapterTest::isFilterDatabaseName));

      controller.clearSheetAutofilter(sheet);

      assertFalse(sheet.getCTWorksheet().isSetAutoFilter());
      assertTrue(
          sheet.getWorkbook().getAllNames().stream()
              .noneMatch(AutofilterAdapterTest::isFilterDatabaseName));
    }
  }

  @Test
  void clearSheetAutofilter_removesOnlySameSheetFilterDatabaseNames() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet ops = populatedSheet(workbook);
      XSSFSheet archive = populatedSheet(workbook, "Archive");
      controller.setSheetAutofilter(ops, "A1:B3");

      Name archiveFilter = workbook.createName();
      archiveFilter.setNameName("_xlnm._FilterDatabase");
      archiveFilter.setSheetIndex(workbook.getSheetIndex(archive));
      archiveFilter.setRefersToFormula("Archive!$A$1:$B$3");

      Name unrelated = workbook.createName();
      unrelated.setNameName("OpsRange");
      unrelated.setSheetIndex(workbook.getSheetIndex(ops));
      unrelated.setRefersToFormula("Ops!$A$1:$A$3");

      controller.clearSheetAutofilter(ops);

      assertTrue(workbook.getAllNames().contains(archiveFilter));
      assertTrue(workbook.getAllNames().contains(unrelated));
      assertFalse(
          workbook.getAllNames().stream()
              .anyMatch(
                  name ->
                      name.getSheetIndex() == workbook.getSheetIndex(ops)
                          && isFilterDatabaseName(name)));
    }
  }

  @Test
  void absentSheetAutofilterReadsAsEmptyAndClearsAsNoOp() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);

      assertEquals(List.of(), controller.sheetOwnedAutofilters(sheet));
      assertEquals(0, controller.sheetAutofilterCount(sheet));

      controller.clearSheetAutofilter(sheet);

      assertFalse(sheet.getCTWorksheet().isSetAutoFilter());
      assertEquals(List.of(), controller.sheetOwnedAutofilters(sheet));
    }
  }

  @Test
  void sheetOwnedAutofilters_readsBlankRefAsEmptyString() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      sheet.getCTWorksheet().addNewAutoFilter().setRef("");

      assertEquals(
          List.of(
              new AutofilterEntryReport.SheetOwned(
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel(""))),
          controller.sheetOwnedAutofilters(sheet));
    }
  }

  @Test
  void sheetOwnedAutofilters_readsHiddenButtonOverrides() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      var autoFilter = sheet.getCTWorksheet().addNewAutoFilter();
      autoFilter.setRef("A1:B3");
      var hiddenButton = autoFilter.addNewFilterColumn();
      hiddenButton.setColId(0L);
      hiddenButton.setHiddenButton(true);
      hiddenButton.addNewFilters();
      var visibleButton = autoFilter.addNewFilterColumn();
      visibleButton.setColId(1L);
      visibleButton.setHiddenButton(false);
      visibleButton.addNewFilters();

      assertEquals(
          List.of(
              new AutofilterEntryReport.SheetOwned(
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"),
                  List.of(
                      new AutofilterFilterColumnReport(
                          0L, false, new AutofilterFilterCriterionReport.Values(List.of(), false)),
                      new AutofilterFilterColumnReport(
                          1L, true, new AutofilterFilterCriterionReport.Values(List.of(), false))),
                  java.util.Optional.empty())),
          controller.sheetOwnedAutofilters(sheet));
    }
  }

  @Test
  void sheetOwnedAutofilters_readsPersistedCriteriaDetails() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      AutofilterEntryReport.SheetOwned snapshot = persistedAutofilterSnapshot(workbook);

      assertEquals("A1:H5", snapshot.range().value());
      assertEquals(8, snapshot.filterColumns().size());
      assertEquals(
          new AutofilterFilterCriterionReport.Values(List.of("Queued", ""), true),
          snapshot.filterColumns().get(0).criterion());
      AutofilterFilterCriterionReport.Custom custom =
          assertInstanceOf(
              AutofilterFilterCriterionReport.Custom.class,
              snapshot.filterColumns().get(1).criterion());
      assertTrue(custom.and());
      assertEquals("lessThan", custom.conditions().getFirst().operator());
      assertEquals("equal", custom.conditions().get(1).operator());
      assertEquals(
          new AutofilterFilterCriterionReport.Dynamic(
              "today", Optional.of(1.0d), Optional.of(2.0d)),
          snapshot.filterColumns().get(2).criterion());
      assertEquals(
          new AutofilterFilterCriterionReport.Top10(false, true, 10.0d, Optional.of(8.0d)),
          snapshot.filterColumns().get(3).criterion());
      assertEquals(
          new AutofilterFilterCriterionReport.Color(
              true, Optional.ofNullable(CellColorReport.rgb(new RgbHex("#AABBCC")))),
          snapshot.filterColumns().get(4).criterion());
      assertEquals(
          new AutofilterFilterCriterionReport.Icon("3TrafficLights1", 2),
          snapshot.filterColumns().get(5).criterion());
      assertEquals(
          new AutofilterFilterCriterionReport.Values(List.of(), false),
          snapshot.filterColumns().get(6).criterion());
      assertEquals(
          new AutofilterFilterCriterionReport.Color(false, Optional.empty()),
          snapshot.filterColumns().get(7).criterion());
    }
  }

  @Test
  void sheetOwnedAutofilters_readsPersistedSortStateDetails() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      AutofilterEntryReport.SheetOwned snapshot = persistedAutofilterSnapshot(workbook);
      AutofilterSortStateReport sortSnapshot = snapshot.sortState().orElseThrow();
      assertEquals("A1:H5", sortSnapshot.range().value());
      assertTrue(sortSnapshot.caseSensitive());
      assertTrue(sortSnapshot.columnSort());
      assertEquals(Optional.of(ExcelAutofilterSortMethod.STROKE), sortSnapshot.sortMethod());
      assertEquals(4, sortSnapshot.conditions().size());
      assertEquals(
          new AutofilterSortConditionReport.CellColor(
              new dev.erst.gridgrind.core.vocabulary.RangeLabel("A2:A5"),
              true,
              CellColorReport.rgb(new RgbHex("#AABBCC"))),
          sortSnapshot.conditions().get(0));
      assertEquals(
          new AutofilterSortConditionReport.FontColor(
              new dev.erst.gridgrind.core.vocabulary.RangeLabel("B2:B5"),
              false,
              CellColorReport.rgb(new RgbHex("#102030"))),
          sortSnapshot.conditions().get(1));
      assertEquals(
          new AutofilterSortConditionReport.Value(
              new dev.erst.gridgrind.core.vocabulary.RangeLabel("C2:C5"), false),
          sortSnapshot.conditions().get(2));
      assertEquals(
          new AutofilterSortConditionReport.Icon(
              new dev.erst.gridgrind.core.vocabulary.RangeLabel("D2:D5"), false, 3),
          sortSnapshot.conditions().get(3));
    }
  }

  @Test
  void setSheetAutofilter_rejectsBlankHeaderRows() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = workbook.createSheet("Ops");
      sheet.createRow(1).createCell(0).setCellValue("Ada");
      sheet.getRow(1).createCell(1).setCellValue("Queue");

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class, () -> controller.setSheetAutofilter(sheet, "A1:B2"));

      assertEquals(
          "autofilter range must include a nonblank header row: A1:B2", failure.getMessage());
    }
  }

  @Test
  void helperSnapshotsPreserveDefaultsAndDxfFallbacks() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      long fillDxfId = putFillDxf(workbook, new byte[] {(byte) 0xAA, (byte) 0xBB, (byte) 0xCC});
      long fontDxfId = putFontDxf(workbook, new byte[] {0x10, 0x20, 0x30});
      long emptyDxfId = workbook.getStylesSource().putDxf(CTDxf.Factory.newInstance()) - 1L;

      var values =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilters.Factory.newInstance();
      assertEquals(
          new AutofilterFilterCriterionReport.Values(List.of(), false),
          AutofilterAdapter.valuesCriterion(values));

      var customFilters =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCustomFilters.Factory.newInstance();
      customFilters.addNewCustomFilter().setVal("Ada");
      assertEquals(
          new AutofilterFilterCriterionReport.Custom(
              false,
              List.of(new AutofilterFilterCriterionReport.CustomConditionReport("equal", "Ada"))),
          AutofilterAdapter.customCriterion(customFilters));

      var dynamic =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDynamicFilter.Factory.newInstance();
      assertEquals(
          new AutofilterFilterCriterionReport.Dynamic(
              "UNKNOWN", Optional.empty(), Optional.empty()),
          AutofilterAdapter.dynamicCriterion(dynamic));

      var top10 = org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTop10.Factory.newInstance();
      top10.setVal(5.0d);
      assertEquals(
          new AutofilterFilterCriterionReport.Top10(true, false, 5.0d, Optional.empty()),
          AutofilterAdapter.top10Criterion(top10));

      var icon =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTIconFilter.Factory.newInstance();
      assertEquals(
          new AutofilterFilterCriterionReport.Icon("UNKNOWN", 0),
          AutofilterAdapter.iconCriterion(icon));

      assertEquals(
          Optional.of(CellColorReport.rgb(new RgbHex("#AABBCC"))),
          AutofilterAdapter.dxfColor(workbook, fillDxfId, true));
      assertEquals(
          Optional.of(CellColorReport.rgb(new RgbHex("#102030"))),
          AutofilterAdapter.dxfColor(workbook, fontDxfId, false));
      assertEquals(
          Optional.of(CellColorReport.rgb(new RgbHex("#AABBCC"))),
          AutofilterAdapter.dxfColor(workbook, fillDxfId, false));
      assertEquals(
          Optional.of(CellColorReport.rgb(new RgbHex("#102030"))),
          AutofilterAdapter.dxfColor(workbook, fontDxfId, true));
      assertEquals(Optional.empty(), AutofilterAdapter.dxfColor(workbook, emptyDxfId, true));
      assertEquals(Optional.empty(), AutofilterAdapter.dxfColor(workbook, 99L, true));
      assertEquals(Optional.empty(), AutofilterAdapter.dxfAt(workbook.getStylesSource(), -1L));
      assertTrue(AutofilterAdapter.dxfAt(workbook.getStylesSource(), fillDxfId).isPresent());
      // Gradient fill (no patternFill) — covers isSetPatternFill()=false in both cellColor and
      // fallback-fill branches.
      long gradientFillDxfId = putGradientFillDxf(workbook);
      assertEquals(Optional.empty(), AutofilterAdapter.dxfColor(workbook, gradientFillDxfId, true));
      assertEquals(
          Optional.empty(), AutofilterAdapter.dxfColor(workbook, gradientFillDxfId, false));
      // PatternFill with no fgColor — covers isSetFgColor()=false in both cellColor and
      // fallback-fill branches.
      long noFgColorDxfId = putPatternFillNoFgColorDxf(workbook);
      assertEquals(Optional.empty(), AutofilterAdapter.dxfColor(workbook, noFgColorDxfId, true));
      assertEquals(Optional.empty(), AutofilterAdapter.dxfColor(workbook, noFgColorDxfId, false));
      // Font with no color array — covers sizeOfColorArray()=0 in both font branches.
      long fontNoColorDxfId = putFontNoColorDxf(workbook);
      assertEquals(Optional.empty(), AutofilterAdapter.dxfColor(workbook, fontNoColorDxfId, false));
      assertEquals(Optional.empty(), AutofilterAdapter.dxfColor(workbook, fontNoColorDxfId, true));

      var autoFilter =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter.Factory.newInstance();
      var sortState = autoFilter.addNewSortState();
      autoFilter.setRef("A1:A5");
      sortState.addNewSortCondition().setRef("A2:A5");
      AutofilterSortStateReport sortSnapshot =
          AutofilterAdapter.sortState(workbook, autoFilter).orElseThrow();
      assertEquals("A1:A5", sortSnapshot.range().value());
      assertEquals(Optional.empty(), sortSnapshot.sortMethod());
    }
  }

  @Test
  void directColorAndSortHelpersCoverRemainingAutofilterBranches() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      long fillDxfId = putFillDxf(workbook, new byte[] {(byte) 0xAA, (byte) 0xBB, (byte) 0xCC});
      long fontDxfId = putFontDxf(workbook, new byte[] {0x10, 0x20, 0x30});
      long mixedDxfId =
          putFillAndFontDxf(
              workbook,
              new byte[] {0x21, 0x43, 0x65},
              new byte[] {(byte) 0xFE, (byte) 0xDC, (byte) 0xBA});

      var absentColorFilter =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColorFilter.Factory.newInstance();
      assertEquals(
          new AutofilterFilterCriterionReport.Color(false, Optional.empty()),
          AutofilterAdapter.colorCriterion(workbook, absentColorFilter));

      var fillColorFilter =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColorFilter.Factory.newInstance();
      fillColorFilter.setCellColor(true);
      fillColorFilter.setDxfId(fillDxfId);
      assertEquals(
          new AutofilterFilterCriterionReport.Color(
              true, Optional.ofNullable(CellColorReport.rgb(new RgbHex("#AABBCC")))),
          AutofilterAdapter.colorCriterion(workbook, fillColorFilter));

      var fontColorFilter =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColorFilter.Factory.newInstance();
      fontColorFilter.setDxfId(fontDxfId);
      assertEquals(
          new AutofilterFilterCriterionReport.Color(
              false, Optional.ofNullable(CellColorReport.rgb(new RgbHex("#102030")))),
          AutofilterAdapter.colorCriterion(workbook, fontColorFilter));

      assertEquals(
          Optional.of(CellColorReport.rgb(new RgbHex("#214365"))),
          AutofilterAdapter.dxfColor(workbook, mixedDxfId, true));
      assertEquals(
          Optional.of(CellColorReport.rgb(new RgbHex("#FEDCBA"))),
          AutofilterAdapter.dxfColor(workbook, mixedDxfId, false));

      var autoFilter =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter.Factory.newInstance();
      assertEquals(Optional.empty(), AutofilterAdapter.sortState(workbook, autoFilter));

      var sortState = autoFilter.addNewSortState();
      sortState.setRef("A1:E5");

      var plainSort = sortState.addNewSortCondition();
      plainSort.setRef("A2:A5");

      var cellColorSort = sortState.addNewSortCondition();
      cellColorSort.setRef("B2:B5");
      cellColorSort.setDescending(true);
      cellColorSort.setSortBy(STSortBy.CELL_COLOR);
      cellColorSort.setDxfId(fillDxfId);

      var fontColorSort = sortState.addNewSortCondition();
      fontColorSort.setRef("C2:C5");
      fontColorSort.setSortBy(STSortBy.FONT_COLOR);
      fontColorSort.setDxfId(fontDxfId);

      var fallbackFillSort = sortState.addNewSortCondition();
      fallbackFillSort.setRef("D2:D5");
      fallbackFillSort.setDxfId(fillDxfId);

      var iconOnlySort = sortState.addNewSortCondition();
      iconOnlySort.setRef("E2:E5");
      iconOnlySort.setSortBy(STSortBy.ICON);
      iconOnlySort.setIconId(4L);

      AutofilterSortStateReport snapshot =
          AutofilterAdapter.sortState(workbook, autoFilter).orElseThrow();
      assertEquals("A1:E5", snapshot.range().value());
      assertFalse(snapshot.caseSensitive());
      assertFalse(snapshot.columnSort());
      assertEquals(Optional.empty(), snapshot.sortMethod());
      assertEquals(5, snapshot.conditions().size());
      assertEquals(
          new AutofilterSortConditionReport.Value(
              new dev.erst.gridgrind.core.vocabulary.RangeLabel("A2:A5"), false),
          snapshot.conditions().get(0));
      assertEquals(
          new AutofilterSortConditionReport.CellColor(
              new dev.erst.gridgrind.core.vocabulary.RangeLabel("B2:B5"),
              true,
              CellColorReport.rgb(new RgbHex("#AABBCC"))),
          snapshot.conditions().get(1));
      assertEquals(
          new AutofilterSortConditionReport.FontColor(
              new dev.erst.gridgrind.core.vocabulary.RangeLabel("C2:C5"),
              false,
              CellColorReport.rgb(new RgbHex("#102030"))),
          snapshot.conditions().get(2));
      assertEquals(
          new AutofilterSortConditionReport.Value(
              new dev.erst.gridgrind.core.vocabulary.RangeLabel("D2:D5"), false),
          snapshot.conditions().get(3));
      assertEquals(
          new AutofilterSortConditionReport.Icon(
              new dev.erst.gridgrind.core.vocabulary.RangeLabel("E2:E5"), false, 4),
          snapshot.conditions().get(4));

      var reflectedIconSort =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSortCondition.Factory.newInstance();
      reflectedIconSort.setRef("F2:F5");
      reflectedIconSort.setSortBy(STSortBy.ICON);
      reflectedIconSort.setIconId(2L);
      assertEquals(
          new AutofilterSortConditionReport.Icon(
              new dev.erst.gridgrind.core.vocabulary.RangeLabel("F2:F5"), false, 2),
          AutofilterAdapter.sortConditionSnapshot(workbook, reflectedIconSort));
    }
  }

  @Test
  void rawCriteriaHelpersCoverExplicitFalseFlags() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var autoFilter =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter.Factory.newInstance();
      var sortState = autoFilter.addNewSortState();
      sortState.setRef("A1:B3");
      sortState.setCaseSensitive(false);
      sortState.setColumnSort(false);
      var explicitFalseSort = sortState.addNewSortCondition();
      explicitFalseSort.setRef("A2:A3");
      explicitFalseSort.setDescending(false);

      assertEquals(
          Optional.of(
              new AutofilterSortStateReport(
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"),
                  false,
                  false,
                  Optional.empty(),
                  List.of(
                      new AutofilterSortConditionReport.Value(
                          new dev.erst.gridgrind.core.vocabulary.RangeLabel("A2:A3"), false)))),
          AutofilterAdapter.sortState(workbook, autoFilter));

      var values =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilters.Factory.newInstance();
      values.addNewFilter().setVal("Ada");
      values.setBlank(false);
      assertEquals(
          new AutofilterFilterCriterionReport.Values(List.of("Ada"), false),
          AutofilterAdapter.valuesCriterion(values));

      var customFilters =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCustomFilters.Factory.newInstance();
      customFilters.setAnd(false);
      customFilters.addNewCustomFilter().setVal("Queue");
      assertEquals(
          new AutofilterFilterCriterionReport.Custom(
              false,
              List.of(new AutofilterFilterCriterionReport.CustomConditionReport("equal", "Queue"))),
          AutofilterAdapter.customCriterion(customFilters));

      var top10 = org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTop10.Factory.newInstance();
      top10.setTop(false);
      top10.setPercent(false);
      top10.setVal(10d);
      assertEquals(
          new AutofilterFilterCriterionReport.Top10(false, false, 10d, Optional.empty()),
          AutofilterAdapter.top10Criterion(top10));

      var defaultTop =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTop10.Factory.newInstance();
      defaultTop.setVal(5d);
      assertEquals(
          new AutofilterFilterCriterionReport.Top10(true, false, 5d, Optional.empty()),
          AutofilterAdapter.top10Criterion(defaultTop));

      var explicitTop =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTop10.Factory.newInstance();
      explicitTop.setTop(true);
      explicitTop.setVal(7d);
      assertEquals(
          new AutofilterFilterCriterionReport.Top10(true, false, 7d, Optional.empty()),
          AutofilterAdapter.top10Criterion(explicitTop));

      var percentTop =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTop10.Factory.newInstance();
      percentTop.setPercent(true);
      percentTop.setVal(3d);
      percentTop.setFilterVal(2d);
      assertEquals(
          new AutofilterFilterCriterionReport.Top10(true, true, 3d, Optional.of(2d)),
          AutofilterAdapter.top10Criterion(percentTop));
    }
  }

  @Test
  void setSheetAutofilter_acceptsBlankSortBy() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);

      controller.setSheetAutofilter(
          sheet,
          "A1:B3",
          List.of(),
          Optional.of(
              new AutofilterSortStateSpec(
                  new dev.erst.gridgrind.core.vocabulary.RangeRef("A1:B3"),
                  false,
                  false,
                  Optional.empty(),
                  List.of(
                      new AutofilterSortConditionSpec.Value(
                          new dev.erst.gridgrind.core.vocabulary.RangeRef("A2:A3"), false)))));

      var snapshot = controller.sheetOwnedAutofilters(sheet).getFirst();
      assertEquals(
          new AutofilterSortConditionReport.Value(
              new dev.erst.gridgrind.core.vocabulary.RangeLabel("A2:A3"), false),
          assertInstanceOf(AutofilterEntryReport.SheetOwned.class, snapshot)
              .sortState()
              .orElseThrow()
              .conditions()
              .getFirst());
    }
  }

  @Test
  void setSheetAutofilter_omitsUnsetDynamicBounds() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedAutofilterWriteSheet(workbook);

      controller.setSheetAutofilter(
          sheet,
          "A1:C3",
          List.of(
              new AutofilterFilterColumnSpec(
                  0L,
                  true,
                  new AutofilterFilterCriterionSpec.Dynamic(
                      "today", Optional.empty(), Optional.empty()))),
          Optional.empty());

      AutofilterEntryReport.SheetOwned snapshot =
          assertInstanceOf(
              AutofilterEntryReport.SheetOwned.class,
              controller.sheetOwnedAutofilters(sheet).getFirst());
      assertEquals(
          new AutofilterFilterCriterionReport.Dynamic("today", Optional.empty(), Optional.empty()),
          snapshot.filterColumns().getFirst().criterion());
    }
  }

  @Test
  void setSheetAutofilter_rejectsOverlapWithExistingTables() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      sheet.createTable(new AreaReference("A1:B3", SpreadsheetVersion.EXCEL2007));

      IllegalArgumentException failure =
          assertThrows(
              IllegalArgumentException.class, () -> controller.setSheetAutofilter(sheet, "A1:B3"));

      assertEquals(
          "sheet-level autofilter range must not overlap an existing table range: Table1@A1:B3",
          failure.getMessage());
    }
  }

  @Test
  void sheetAutofilterHealthFlagsInvalidRangesAndTableOverlap() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      sheet.getCTWorksheet().addNewAutoFilter().setRef("A0:B3");

      List<AnalysisFindingReport> invalidRangeFindings =
          controller.sheetAutofilterHealthFindings("Ops", sheet, List.of());

      assertEquals(1, invalidRangeFindings.size());
      assertEquals(
          AnalysisFindingCode.AUTOFILTER_INVALID_RANGE, invalidRangeFindings.getFirst().code());
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      sheet.setAutoFilter(CellRangeAddress.valueOf("A1:B3"));
      sheet.createTable(new AreaReference("A1:B3", SpreadsheetVersion.EXCEL2007));

      List<AnalysisFindingReport> overlapFindings =
          controller.sheetAutofilterHealthFindings(
              "Ops",
              sheet,
              List.of(
                  new TableEntryReport(
                      new dev.erst.gridgrind.core.vocabulary.TableLabel("Table1"),
                      new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
                      new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"),
                      1,
                      0,
                      List.of("Owner", "Task"),
                      new TableStyleReport.None(),
                      false)));

      assertEquals(1, overlapFindings.size());
      assertEquals(
          AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH, overlapFindings.getFirst().code());
    }
  }

  @Test
  void sheetAutofilterHealthFlagsBlankHeaderRows() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      sheet.getRow(0).getCell(0).setBlank();
      sheet.getRow(0).getCell(1).setBlank();
      sheet.setAutoFilter(CellRangeAddress.valueOf("A1:B3"));

      List<AnalysisFindingReport> findings =
          controller.sheetAutofilterHealthFindings("Ops", sheet, List.of());

      assertEquals(1, findings.size());
      assertEquals(AnalysisFindingCode.AUTOFILTER_MISSING_HEADER_ROW, findings.getFirst().code());
    }
  }

  @Test
  void sheetAutofilterHealthIgnoresInvalidAndNonOverlappingTables() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      sheet.setAutoFilter(CellRangeAddress.valueOf("A1:B3"));

      List<AnalysisFindingReport> findings =
          controller.sheetAutofilterHealthFindings(
              "Ops",
              sheet,
              List.of(
                  new TableEntryReport(
                      new dev.erst.gridgrind.core.vocabulary.TableLabel("Invalid"),
                      new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
                      new dev.erst.gridgrind.core.vocabulary.RangeLabel("A0:B3"),
                      1,
                      0,
                      List.of("Owner", "Task"),
                      new TableStyleReport.None(),
                      false),
                  new TableEntryReport(
                      new dev.erst.gridgrind.core.vocabulary.TableLabel("Desk"),
                      new dev.erst.gridgrind.core.vocabulary.SheetLabel("Ops"),
                      new dev.erst.gridgrind.core.vocabulary.RangeLabel("D1:E3"),
                      1,
                      0,
                      List.of("Desk", "Region"),
                      new TableStyleReport.None(),
                      false)));

      assertEquals(List.of(), findings);
    }
  }

  @Test
  void setSheetAutofilter_allowsInvalidExistingTableMetadataWhenItDoesNotOverlap()
      throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      XSSFTable table = sheet.createTable(new AreaReference("D1:E3", SpreadsheetVersion.EXCEL2007));
      table.getCTTable().setRef("D0:E3");

      controller.setSheetAutofilter(sheet, "A1:B3");

      assertEquals(
          List.of(
              new AutofilterEntryReport.SheetOwned(
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"))),
          controller.sheetOwnedAutofilters(sheet));
    }
  }

  @Test
  void setSheetAutofilter_allowsValidNonOverlappingTableRanges() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      sheet.getRow(0).createCell(3).setCellValue("Desk");
      sheet.getRow(0).createCell(4).setCellValue("Region");
      sheet.getRow(1).createCell(3).setCellValue("A1");
      sheet.getRow(1).createCell(4).setCellValue("North");
      sheet.getRow(2).createCell(3).setCellValue("B1");
      sheet.getRow(2).createCell(4).setCellValue("South");
      sheet.createTable(new AreaReference("D1:E3", SpreadsheetVersion.EXCEL2007));

      controller.setSheetAutofilter(sheet, "A1:B3");

      assertEquals(
          List.of(
              new AutofilterEntryReport.SheetOwned(
                  new dev.erst.gridgrind.core.vocabulary.RangeLabel("A1:B3"))),
          controller.sheetOwnedAutofilters(sheet));
    }
  }

  private static XSSFSheet populatedSheet(XSSFWorkbook workbook) {
    return populatedSheet(workbook, "Ops");
  }

  private static XSSFSheet populatedSheet(XSSFWorkbook workbook, String name) {
    XSSFSheet sheet = workbook.createSheet(name);
    sheet.createRow(0).createCell(0).setCellValue("Owner");
    sheet.getRow(0).createCell(1).setCellValue("Task");
    sheet.createRow(1).createCell(0).setCellValue("Ada");
    sheet.getRow(1).createCell(1).setCellValue("Queue");
    sheet.createRow(2).createCell(0).setCellValue("Lin");
    sheet.getRow(2).createCell(1).setCellValue("Pack");
    return sheet;
  }

  private static XSSFSheet populatedAutofilterWriteSheet(XSSFWorkbook workbook) {
    XSSFSheet sheet = workbook.createSheet("Ops");
    sheet.createRow(0).createCell(0).setCellValue("Owner");
    sheet.getRow(0).createCell(1).setCellValue("Task");
    sheet.getRow(0).createCell(2).setCellValue("Stage");
    sheet.getRow(0).createCell(3).setCellValue("Score");
    sheet.getRow(0).createCell(4).setCellValue("Tone");
    sheet.getRow(0).createCell(5).setCellValue("Flag");
    sheet.createRow(1).createCell(0).setCellValue("Ada");
    sheet.getRow(1).createCell(1).setCellValue("Queue");
    sheet.getRow(1).createCell(2).setCellValue("Today");
    sheet.getRow(1).createCell(3).setCellValue(4);
    sheet.getRow(1).createCell(4).setCellValue("Amber");
    sheet.getRow(1).createCell(5).setCellValue("High");
    sheet.createRow(2).createCell(0).setCellValue("Lin");
    sheet.getRow(2).createCell(1).setCellValue("Pack");
    sheet.getRow(2).createCell(2).setCellValue("Today");
    sheet.getRow(2).createCell(3).setCellValue(6);
    sheet.getRow(2).createCell(4).setCellValue("Green");
    sheet.getRow(2).createCell(5).setCellValue("Low");
    return sheet;
  }

  private static boolean isFilterDatabaseName(org.apache.poi.ss.usermodel.Name name) {
    return "_XLNM._FILTERDATABASE".equalsIgnoreCase(name.getNameName());
  }

  private static long putFillDxf(XSSFWorkbook workbook, byte[] rgb) {
    CTDxf dxf = CTDxf.Factory.newInstance();
    dxf.addNewFill().addNewPatternFill().addNewFgColor().setRgb(rgb);
    return workbook.getStylesSource().putDxf(dxf) - 1L;
  }

  private static long putFontDxf(XSSFWorkbook workbook, byte[] rgb) {
    CTDxf dxf = CTDxf.Factory.newInstance();
    dxf.addNewFont().addNewColor().setRgb(rgb);
    return workbook.getStylesSource().putDxf(dxf) - 1L;
  }

  private static long putFillAndFontDxf(XSSFWorkbook workbook, byte[] fillRgb, byte[] fontRgb) {
    CTDxf dxf = CTDxf.Factory.newInstance();
    dxf.addNewFill().addNewPatternFill().addNewFgColor().setRgb(fillRgb);
    dxf.addNewFont().addNewColor().setRgb(fontRgb);
    return workbook.getStylesSource().putDxf(dxf) - 1L;
  }

  private static long putGradientFillDxf(XSSFWorkbook workbook) {
    CTDxf dxf = CTDxf.Factory.newInstance();
    dxf.addNewFill().addNewGradientFill();
    return workbook.getStylesSource().putDxf(dxf) - 1L;
  }

  private static long putPatternFillNoFgColorDxf(XSSFWorkbook workbook) {
    CTDxf dxf = CTDxf.Factory.newInstance();
    dxf.addNewFill().addNewPatternFill();
    return workbook.getStylesSource().putDxf(dxf) - 1L;
  }

  private static long putFontNoColorDxf(XSSFWorkbook workbook) {
    CTDxf dxf = CTDxf.Factory.newInstance();
    dxf.addNewFont();
    return workbook.getStylesSource().putDxf(dxf) - 1L;
  }

  private AutofilterEntryReport.SheetOwned persistedAutofilterSnapshot(XSSFWorkbook workbook) {
    XSSFSheet sheet = populatedSheet(workbook);
    long fillDxfId = putFillDxf(workbook, new byte[] {(byte) 0xAA, (byte) 0xBB, (byte) 0xCC});
    long fontDxfId = putFontDxf(workbook, new byte[] {0x10, 0x20, 0x30});
    var autoFilter = sheet.getCTWorksheet().addNewAutoFilter();
    autoFilter.setRef("A1:H5");
    addPersistedFilterColumns(autoFilter, fillDxfId);
    addPersistedSortState(autoFilter, fillDxfId, fontDxfId);
    return assertInstanceOf(
        AutofilterEntryReport.SheetOwned.class, controller.sheetOwnedAutofilters(sheet).getFirst());
  }

  private static void addPersistedFilterColumns(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter autoFilter, long fillDxfId) {
    var valuesColumn = autoFilter.addNewFilterColumn();
    valuesColumn.setColId(0L);
    valuesColumn.setShowButton(false);
    var values = valuesColumn.addNewFilters();
    values.setBlank(true);
    values.addNewFilter().setVal("Queued");
    values.addNewFilter();

    var customColumn = autoFilter.addNewFilterColumn();
    customColumn.setColId(1L);
    var customFilters = customColumn.addNewCustomFilters();
    customFilters.setAnd(true);
    var lessThan = customFilters.addNewCustomFilter();
    lessThan.setOperator(STFilterOperator.LESS_THAN);
    lessThan.setVal("10");
    customFilters.addNewCustomFilter().setVal("Ada");

    var dynamicColumn = autoFilter.addNewFilterColumn();
    dynamicColumn.setColId(2L);
    var dynamicFilter = dynamicColumn.addNewDynamicFilter();
    dynamicFilter.setType(STDynamicFilterType.TODAY);
    dynamicFilter.setVal(1.0d);
    dynamicFilter.setMaxVal(2.0d);

    var top10Column = autoFilter.addNewFilterColumn();
    top10Column.setColId(3L);
    var top10 = top10Column.addNewTop10();
    top10.setTop(false);
    top10.setPercent(true);
    top10.setVal(10.0d);
    top10.setFilterVal(8.0d);

    var colorColumn = autoFilter.addNewFilterColumn();
    colorColumn.setColId(4L);
    var colorFilter = colorColumn.addNewColorFilter();
    colorFilter.setCellColor(true);
    colorFilter.setDxfId(fillDxfId);

    var iconColumn = autoFilter.addNewFilterColumn();
    iconColumn.setColId(5L);
    var iconFilter = iconColumn.addNewIconFilter();
    iconFilter.setIconSet(STIconSetType.X_3_TRAFFIC_LIGHTS_1);
    iconFilter.setIconId(2L);

    autoFilter.addNewFilterColumn().setColId(6L);
    var missingDxfColorColumn = autoFilter.addNewFilterColumn();
    missingDxfColorColumn.setColId(7L);
    missingDxfColorColumn.addNewColorFilter().setDxfId(99L);
  }

  private static void addPersistedSortState(
      org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter autoFilter,
      long fillDxfId,
      long fontDxfId) {
    var sortState = autoFilter.addNewSortState();
    sortState.setRef("A1:H5");
    sortState.setCaseSensitive(true);
    sortState.setColumnSort(true);
    sortState.setSortMethod(STSortMethod.STROKE);

    var cellColorSort = sortState.addNewSortCondition();
    cellColorSort.setRef("A2:A5");
    cellColorSort.setDescending(true);
    cellColorSort.setSortBy(STSortBy.CELL_COLOR);
    cellColorSort.setDxfId(fillDxfId);
    cellColorSort.setIconId(1L);

    var fontColorSort = sortState.addNewSortCondition();
    fontColorSort.setRef("B2:B5");
    fontColorSort.setSortBy(STSortBy.FONT_COLOR);
    fontColorSort.setDxfId(fontDxfId);

    var fallbackFillSort = sortState.addNewSortCondition();
    fallbackFillSort.setRef("C2:C5");
    fallbackFillSort.setDxfId(fillDxfId);

    var missingDxfSort = sortState.addNewSortCondition();
    missingDxfSort.setRef("D2:D5");
    missingDxfSort.setSortBy(STSortBy.ICON);
    missingDxfSort.setDxfId(99L);
    missingDxfSort.setIconId(3L);
  }
}
