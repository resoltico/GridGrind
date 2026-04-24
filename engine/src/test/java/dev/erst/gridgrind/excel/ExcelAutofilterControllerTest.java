package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

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
class ExcelAutofilterControllerTest {
  private final ExcelAutofilterController controller = new ExcelAutofilterController();

  @Test
  void setSheetAutofilter_roundTripsOwnedMetadata() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);

      controller.setSheetAutofilter(sheet, "A1:B3");

      assertEquals(
          List.of(new ExcelAutofilterSnapshot.SheetOwned("A1:B3")),
          controller.sheetOwnedAutofilters(sheet));
      assertEquals(1, controller.sheetAutofilterCount(sheet));

      ByteArrayOutputStream output = new ByteArrayOutputStream();
      workbook.write(output);
      try (XSSFWorkbook reopened =
          new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()))) {
        XSSFSheet reopenedSheet = reopened.getSheet("Ops");
        assertEquals(
            List.of(new ExcelAutofilterSnapshot.SheetOwned("A1:B3")),
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
              new ExcelAutofilterFilterColumn(
                  0L, false, new ExcelAutofilterFilterCriterion.Values(List.of("Ada"), true)),
              new ExcelAutofilterFilterColumn(
                  1L,
                  true,
                  new ExcelAutofilterFilterCriterion.Custom(
                      true,
                      List.of(
                          new ExcelAutofilterFilterCriterion.CustomCondition("greaterThan", "1"),
                          new ExcelAutofilterFilterCriterion.CustomCondition("equal", "Queue")))),
              new ExcelAutofilterFilterColumn(
                  2L, true, new ExcelAutofilterFilterCriterion.Dynamic("today", 1.0d, 2.0d)),
              new ExcelAutofilterFilterColumn(
                  3L, true, new ExcelAutofilterFilterCriterion.Top10(10, false, true)),
              new ExcelAutofilterFilterColumn(
                  4L,
                  true,
                  new ExcelAutofilterFilterCriterion.Color(false, new ExcelColor("#AABBCC"))),
              new ExcelAutofilterFilterColumn(
                  5L, true, new ExcelAutofilterFilterCriterion.Icon("3TrafficLights1", 2))),
          new ExcelAutofilterSortState(
              "A1:F3",
              true,
              true,
              STSortMethod.STROKE.toString(),
              List.of(
                  new ExcelAutofilterSortCondition(
                      "A2:A3",
                      true,
                      STSortBy.CELL_COLOR.toString(),
                      new ExcelColor("#102030"),
                      null),
                  new ExcelAutofilterSortCondition(
                      "B2:B3", false, STSortBy.ICON.toString(), null, 4))));

      ExcelAutofilterSnapshot.SheetOwned snapshot =
          assertInstanceOf(
              ExcelAutofilterSnapshot.SheetOwned.class,
              controller.sheetOwnedAutofilters(sheet).getFirst());

      assertEquals("A1:F3", snapshot.range());
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Values(List.of("Ada"), true),
          snapshot.filterColumns().get(0).criterion());
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Custom(
              true,
              List.of(
                  new ExcelAutofilterFilterCriterionSnapshot.CustomCondition("greaterThan", "1"),
                  new ExcelAutofilterFilterCriterionSnapshot.CustomCondition("equal", "Queue"))),
          snapshot.filterColumns().get(1).criterion());
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Dynamic("today", 1.0d, 2.0d),
          snapshot.filterColumns().get(2).criterion());
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Top10(false, true, 10.0d, null),
          snapshot.filterColumns().get(3).criterion());
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Color(
              false, new ExcelColorSnapshot("#AABBCC")),
          snapshot.filterColumns().get(4).criterion());
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Icon("3TrafficLights1", 2),
          snapshot.filterColumns().get(5).criterion());
      assertEquals(
          new ExcelAutofilterSortConditionSnapshot(
              "A2:A3",
              true,
              STSortBy.CELL_COLOR.toString(),
              new ExcelColorSnapshot("#102030"),
              null),
          snapshot.sortState().orElseThrow().conditions().get(0));
      assertEquals(
          new ExcelAutofilterSortConditionSnapshot(
              "B2:B3", false, STSortBy.ICON.toString(), null, 4),
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
                              new ExcelAutofilterFilterColumn(
                                  0L,
                                  true,
                                  new ExcelAutofilterFilterCriterion.Custom(
                                      false,
                                      List.of(
                                          new ExcelAutofilterFilterCriterion.CustomCondition(
                                              "nope", "1"))))),
                          null))
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
                              new ExcelAutofilterFilterColumn(
                                  0L,
                                  true,
                                  new ExcelAutofilterFilterCriterion.Dynamic(
                                      "tomorrowish", null, null))),
                          null))
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
                              new ExcelAutofilterFilterColumn(
                                  0L, true, new ExcelAutofilterFilterCriterion.Icon("unknown", 0))),
                          null))
              .getMessage());
      assertEquals(
          "unsupported autofilter sort method: diagonal",
          assertThrows(
                  IllegalArgumentException.class,
                  () ->
                      controller.setSheetAutofilter(
                          sheet,
                          "A1:F3",
                          List.of(),
                          new ExcelAutofilterSortState(
                              "A1:F3",
                              false,
                              false,
                              "diagonal",
                              List.of(
                                  new ExcelAutofilterSortCondition(
                                      "A2:A3", false, "", null, null)))))
              .getMessage());
      assertEquals(
          "unsupported autofilter sortBy value: sideways",
          assertThrows(
                  IllegalArgumentException.class,
                  () ->
                      controller.setSheetAutofilter(
                          sheet,
                          "A1:F3",
                          List.of(),
                          new ExcelAutofilterSortState(
                              "A1:F3",
                              false,
                              false,
                              "",
                              List.of(
                                  new ExcelAutofilterSortCondition(
                                      "A2:A3", false, "sideways", null, null)))))
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
              .anyMatch(ExcelAutofilterControllerTest::isFilterDatabaseName));

      controller.clearSheetAutofilter(sheet);

      assertFalse(sheet.getCTWorksheet().isSetAutoFilter());
      assertTrue(
          sheet.getWorkbook().getAllNames().stream()
              .noneMatch(ExcelAutofilterControllerTest::isFilterDatabaseName));
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
          List.of(new ExcelAutofilterSnapshot.SheetOwned("")),
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
              new ExcelAutofilterSnapshot.SheetOwned(
                  "A1:B3",
                  List.of(
                      new ExcelAutofilterFilterColumnSnapshot(
                          0L,
                          false,
                          new ExcelAutofilterFilterCriterionSnapshot.Values(List.of(), false)),
                      new ExcelAutofilterFilterColumnSnapshot(
                          1L,
                          true,
                          new ExcelAutofilterFilterCriterionSnapshot.Values(List.of(), false))),
                  java.util.Optional.empty())),
          controller.sheetOwnedAutofilters(sheet));
    }
  }

  @Test
  void sheetOwnedAutofilters_readsPersistedCriteriaDetails() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      ExcelAutofilterSnapshot.SheetOwned snapshot = persistedAutofilterSnapshot(workbook);

      assertEquals("A1:H5", snapshot.range());
      assertEquals(8, snapshot.filterColumns().size());
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Values(List.of("Queued", ""), true),
          snapshot.filterColumns().get(0).criterion());
      ExcelAutofilterFilterCriterionSnapshot.Custom custom =
          assertInstanceOf(
              ExcelAutofilterFilterCriterionSnapshot.Custom.class,
              snapshot.filterColumns().get(1).criterion());
      assertTrue(custom.and());
      assertEquals("lessThan", custom.conditions().getFirst().operator());
      assertEquals("equal", custom.conditions().get(1).operator());
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Dynamic("today", 1.0d, 2.0d),
          snapshot.filterColumns().get(2).criterion());
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Top10(false, true, 10.0d, 8.0d),
          snapshot.filterColumns().get(3).criterion());
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Color(true, new ExcelColorSnapshot("#AABBCC")),
          snapshot.filterColumns().get(4).criterion());
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Icon("3TrafficLights1", 2),
          snapshot.filterColumns().get(5).criterion());
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Values(List.of(), false),
          snapshot.filterColumns().get(6).criterion());
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Color(false, null),
          snapshot.filterColumns().get(7).criterion());
    }
  }

  @Test
  void sheetOwnedAutofilters_readsPersistedSortStateDetails() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      ExcelAutofilterSnapshot.SheetOwned snapshot = persistedAutofilterSnapshot(workbook);
      ExcelAutofilterSortStateSnapshot sortSnapshot = snapshot.sortState().orElseThrow();
      assertEquals("A1:H5", sortSnapshot.range());
      assertTrue(sortSnapshot.caseSensitive());
      assertTrue(sortSnapshot.columnSort());
      assertEquals(STSortMethod.STROKE.toString(), sortSnapshot.sortMethod());
      assertEquals(4, sortSnapshot.conditions().size());
      assertEquals(
          new ExcelAutofilterSortConditionSnapshot(
              "A2:A5", true, STSortBy.CELL_COLOR.toString(), new ExcelColorSnapshot("#AABBCC"), 1),
          sortSnapshot.conditions().get(0));
      assertEquals(
          new ExcelAutofilterSortConditionSnapshot(
              "B2:B5",
              false,
              STSortBy.FONT_COLOR.toString(),
              new ExcelColorSnapshot("#102030"),
              null),
          sortSnapshot.conditions().get(1));
      assertEquals(
          new ExcelAutofilterSortConditionSnapshot(
              "C2:C5", false, "", new ExcelColorSnapshot("#AABBCC"), null),
          sortSnapshot.conditions().get(2));
      assertEquals(
          new ExcelAutofilterSortConditionSnapshot(
              "D2:D5", false, STSortBy.ICON.toString(), null, 3),
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
          new ExcelAutofilterFilterCriterionSnapshot.Values(List.of(), false),
          ExcelAutofilterController.valuesCriterion(values));

      var customFilters =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCustomFilters.Factory.newInstance();
      customFilters.addNewCustomFilter().setVal("Ada");
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Custom(
              false,
              List.of(new ExcelAutofilterFilterCriterionSnapshot.CustomCondition("equal", "Ada"))),
          ExcelAutofilterController.customCriterion(customFilters));

      var dynamic =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDynamicFilter.Factory.newInstance();
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Dynamic("UNKNOWN", null, null),
          ExcelAutofilterController.dynamicCriterion(dynamic));

      var top10 = org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTop10.Factory.newInstance();
      top10.setVal(5.0d);
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Top10(true, false, 5.0d, null),
          ExcelAutofilterController.top10Criterion(top10));

      var icon =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTIconFilter.Factory.newInstance();
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Icon("UNKNOWN", 0),
          ExcelAutofilterController.iconCriterion(icon));

      assertEquals(
          Optional.of(new ExcelColorSnapshot("#AABBCC")),
          ExcelAutofilterController.dxfColor(workbook, fillDxfId, true));
      assertEquals(
          Optional.of(new ExcelColorSnapshot("#102030")),
          ExcelAutofilterController.dxfColor(workbook, fontDxfId, false));
      assertEquals(
          Optional.of(new ExcelColorSnapshot("#AABBCC")),
          ExcelAutofilterController.dxfColor(workbook, fillDxfId, false));
      assertEquals(
          Optional.of(new ExcelColorSnapshot("#102030")),
          ExcelAutofilterController.dxfColor(workbook, fontDxfId, true));
      assertEquals(
          Optional.empty(), ExcelAutofilterController.dxfColor(workbook, emptyDxfId, true));
      assertEquals(Optional.empty(), ExcelAutofilterController.dxfColor(workbook, 99L, true));
      assertEquals(
          Optional.empty(), ExcelAutofilterController.dxfAt(workbook.getStylesSource(), -1L));
      assertTrue(
          ExcelAutofilterController.dxfAt(workbook.getStylesSource(), fillDxfId).isPresent());
      // Gradient fill (no patternFill) — covers isSetPatternFill()=false in both cellColor and
      // fallback-fill branches.
      long gradientFillDxfId = putGradientFillDxf(workbook);
      assertEquals(
          Optional.empty(), ExcelAutofilterController.dxfColor(workbook, gradientFillDxfId, true));
      assertEquals(
          Optional.empty(), ExcelAutofilterController.dxfColor(workbook, gradientFillDxfId, false));
      // PatternFill with no fgColor — covers isSetFgColor()=false in both cellColor and
      // fallback-fill branches.
      long noFgColorDxfId = putPatternFillNoFgColorDxf(workbook);
      assertEquals(
          Optional.empty(), ExcelAutofilterController.dxfColor(workbook, noFgColorDxfId, true));
      assertEquals(
          Optional.empty(), ExcelAutofilterController.dxfColor(workbook, noFgColorDxfId, false));
      // Font with no color array — covers sizeOfColorArray()=0 in both font branches.
      long fontNoColorDxfId = putFontNoColorDxf(workbook);
      assertEquals(
          Optional.empty(), ExcelAutofilterController.dxfColor(workbook, fontNoColorDxfId, false));
      assertEquals(
          Optional.empty(), ExcelAutofilterController.dxfColor(workbook, fontNoColorDxfId, true));

      var autoFilter =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter.Factory.newInstance();
      var sortState = autoFilter.addNewSortState();
      sortState.addNewSortCondition().setRef("A2:A5");
      ExcelAutofilterSortStateSnapshot sortSnapshot =
          controller.sortState(workbook, autoFilter).orElseThrow();
      assertEquals("", sortSnapshot.range());
      assertEquals("", sortSnapshot.sortMethod());
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
          new ExcelAutofilterFilterCriterionSnapshot.Color(false, null),
          ExcelAutofilterController.colorCriterion(workbook, absentColorFilter));

      var fillColorFilter =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColorFilter.Factory.newInstance();
      fillColorFilter.setCellColor(true);
      fillColorFilter.setDxfId(fillDxfId);
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Color(true, new ExcelColorSnapshot("#AABBCC")),
          ExcelAutofilterController.colorCriterion(workbook, fillColorFilter));

      var fontColorFilter =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColorFilter.Factory.newInstance();
      fontColorFilter.setDxfId(fontDxfId);
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Color(
              false, new ExcelColorSnapshot("#102030")),
          ExcelAutofilterController.colorCriterion(workbook, fontColorFilter));

      assertEquals(
          Optional.of(new ExcelColorSnapshot("#214365")),
          ExcelAutofilterController.dxfColor(workbook, mixedDxfId, true));
      assertEquals(
          Optional.of(new ExcelColorSnapshot("#FEDCBA")),
          ExcelAutofilterController.dxfColor(workbook, mixedDxfId, false));

      var autoFilter =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTAutoFilter.Factory.newInstance();
      assertEquals(Optional.empty(), controller.sortState(workbook, autoFilter));

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

      ExcelAutofilterSortStateSnapshot snapshot =
          controller.sortState(workbook, autoFilter).orElseThrow();
      assertEquals("A1:E5", snapshot.range());
      assertFalse(snapshot.caseSensitive());
      assertFalse(snapshot.columnSort());
      assertEquals("", snapshot.sortMethod());
      assertEquals(5, snapshot.conditions().size());
      assertEquals(
          new ExcelAutofilterSortConditionSnapshot("A2:A5", false, "", null, null),
          snapshot.conditions().get(0));
      assertEquals(
          new ExcelAutofilterSortConditionSnapshot(
              "B2:B5",
              true,
              STSortBy.CELL_COLOR.toString(),
              new ExcelColorSnapshot("#AABBCC"),
              null),
          snapshot.conditions().get(1));
      assertEquals(
          new ExcelAutofilterSortConditionSnapshot(
              "C2:C5",
              false,
              STSortBy.FONT_COLOR.toString(),
              new ExcelColorSnapshot("#102030"),
              null),
          snapshot.conditions().get(2));
      assertEquals(
          new ExcelAutofilterSortConditionSnapshot(
              "D2:D5", false, "", new ExcelColorSnapshot("#AABBCC"), null),
          snapshot.conditions().get(3));
      assertEquals(
          new ExcelAutofilterSortConditionSnapshot(
              "E2:E5", false, STSortBy.ICON.toString(), null, 4),
          snapshot.conditions().get(4));

      var reflectedIconSort =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSortCondition.Factory.newInstance();
      reflectedIconSort.setRef("F2:F5");
      reflectedIconSort.setSortBy(STSortBy.ICON);
      reflectedIconSort.setIconId(2L);
      assertEquals(
          new ExcelAutofilterSortConditionSnapshot(
              "F2:F5", false, STSortBy.ICON.toString(), null, 2),
          ExcelAutofilterController.sortConditionSnapshot(workbook, reflectedIconSort));
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
              new ExcelAutofilterSortStateSnapshot(
                  "A1:B3",
                  false,
                  false,
                  "",
                  List.of(
                      new ExcelAutofilterSortConditionSnapshot("A2:A3", false, "", null, null)))),
          controller.sortState(workbook, autoFilter));

      var values =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFilters.Factory.newInstance();
      values.addNewFilter().setVal("Ada");
      values.setBlank(false);
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Values(List.of("Ada"), false),
          ExcelAutofilterController.valuesCriterion(values));

      var customFilters =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCustomFilters.Factory.newInstance();
      customFilters.setAnd(false);
      customFilters.addNewCustomFilter().setVal("Queue");
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Custom(
              false,
              List.of(
                  new ExcelAutofilterFilterCriterionSnapshot.CustomCondition("equal", "Queue"))),
          ExcelAutofilterController.customCriterion(customFilters));

      var top10 = org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTop10.Factory.newInstance();
      top10.setTop(false);
      top10.setPercent(false);
      top10.setVal(10d);
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Top10(false, false, 10d, null),
          ExcelAutofilterController.top10Criterion(top10));

      var defaultTop =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTop10.Factory.newInstance();
      defaultTop.setVal(5d);
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Top10(true, false, 5d, null),
          ExcelAutofilterController.top10Criterion(defaultTop));

      var explicitTop =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTop10.Factory.newInstance();
      explicitTop.setTop(true);
      explicitTop.setVal(7d);
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Top10(true, false, 7d, null),
          ExcelAutofilterController.top10Criterion(explicitTop));

      var percentTop =
          org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTop10.Factory.newInstance();
      percentTop.setPercent(true);
      percentTop.setVal(3d);
      percentTop.setFilterVal(2d);
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Top10(true, true, 3d, 2d),
          ExcelAutofilterController.top10Criterion(percentTop));
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
          new ExcelAutofilterSortState(
              "A1:B3",
              false,
              false,
              "",
              List.of(new ExcelAutofilterSortCondition("A2:A3", false, "", null, null))));

      var snapshot = controller.sheetOwnedAutofilters(sheet).getFirst();
      assertEquals(
          new ExcelAutofilterSortConditionSnapshot("A2:A3", false, "", null, null),
          assertInstanceOf(ExcelAutofilterSnapshot.SheetOwned.class, snapshot)
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
              new ExcelAutofilterFilterColumn(
                  0L, true, new ExcelAutofilterFilterCriterion.Dynamic("today", null, null))),
          null);

      ExcelAutofilterSnapshot.SheetOwned snapshot =
          assertInstanceOf(
              ExcelAutofilterSnapshot.SheetOwned.class,
              controller.sheetOwnedAutofilters(sheet).getFirst());
      assertEquals(
          new ExcelAutofilterFilterCriterionSnapshot.Dynamic("today", null, null),
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

      List<WorkbookAnalysis.AnalysisFinding> invalidRangeFindings =
          controller.sheetAutofilterHealthFindings("Ops", sheet, List.of());

      assertEquals(1, invalidRangeFindings.size());
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_INVALID_RANGE,
          invalidRangeFindings.getFirst().code());
    }

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      sheet.setAutoFilter(CellRangeAddress.valueOf("A1:B3"));
      sheet.createTable(new AreaReference("A1:B3", SpreadsheetVersion.EXCEL2007));

      List<WorkbookAnalysis.AnalysisFinding> overlapFindings =
          controller.sheetAutofilterHealthFindings(
              "Ops",
              sheet,
              List.of(
                  new ExcelTableSnapshot(
                      "Table1",
                      "Ops",
                      "A1:B3",
                      1,
                      0,
                      List.of("Owner", "Task"),
                      new ExcelTableStyleSnapshot.None(),
                      false)));

      assertEquals(1, overlapFindings.size());
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH,
          overlapFindings.getFirst().code());
    }
  }

  @Test
  void sheetAutofilterHealthFlagsBlankHeaderRows() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      sheet.getRow(0).getCell(0).setBlank();
      sheet.getRow(0).getCell(1).setBlank();
      sheet.setAutoFilter(CellRangeAddress.valueOf("A1:B3"));

      List<WorkbookAnalysis.AnalysisFinding> findings =
          controller.sheetAutofilterHealthFindings("Ops", sheet, List.of());

      assertEquals(1, findings.size());
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_MISSING_HEADER_ROW,
          findings.getFirst().code());
    }
  }

  @Test
  void sheetAutofilterHealthIgnoresInvalidAndNonOverlappingTables() throws Exception {
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      XSSFSheet sheet = populatedSheet(workbook);
      sheet.setAutoFilter(CellRangeAddress.valueOf("A1:B3"));

      List<WorkbookAnalysis.AnalysisFinding> findings =
          controller.sheetAutofilterHealthFindings(
              "Ops",
              sheet,
              List.of(
                  new ExcelTableSnapshot(
                      "Invalid",
                      "Ops",
                      "A0:B3",
                      1,
                      0,
                      List.of("Owner", "Task"),
                      new ExcelTableStyleSnapshot.None(),
                      false),
                  new ExcelTableSnapshot(
                      "Desk",
                      "Ops",
                      "D1:E3",
                      1,
                      0,
                      List.of("Desk", "Region"),
                      new ExcelTableStyleSnapshot.None(),
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
          List.of(new ExcelAutofilterSnapshot.SheetOwned("A1:B3")),
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
          List.of(new ExcelAutofilterSnapshot.SheetOwned("A1:B3")),
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

  private ExcelAutofilterSnapshot.SheetOwned persistedAutofilterSnapshot(XSSFWorkbook workbook) {
    XSSFSheet sheet = populatedSheet(workbook);
    long fillDxfId = putFillDxf(workbook, new byte[] {(byte) 0xAA, (byte) 0xBB, (byte) 0xCC});
    long fontDxfId = putFontDxf(workbook, new byte[] {0x10, 0x20, 0x30});
    var autoFilter = sheet.getCTWorksheet().addNewAutoFilter();
    autoFilter.setRef("A1:H5");
    addPersistedFilterColumns(autoFilter, fillDxfId);
    addPersistedSortState(autoFilter, fillDxfId, fontDxfId);
    return assertInstanceOf(
        ExcelAutofilterSnapshot.SheetOwned.class,
        controller.sheetOwnedAutofilters(sheet).getFirst());
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
