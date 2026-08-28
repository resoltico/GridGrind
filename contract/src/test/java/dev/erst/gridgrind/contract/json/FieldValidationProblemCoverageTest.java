package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.dto.ProtocolCellAddressValidation;
import dev.erst.gridgrind.contract.dto.ProtocolDefinedNameValidation;
import dev.erst.gridgrind.excel.foundation.ExcelColumnWidthViolation;
import dev.erst.gridgrind.excel.foundation.ExcelRowHeightViolation;
import dev.erst.gridgrind.excel.foundation.ExcelSheetNameProblem;
import dev.erst.gridgrind.excel.foundation.ExcelZoomViolation;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Covers split field-validation renderers, mappers, and descriptor rebasing helpers. */
class FieldValidationProblemCoverageTest {
  @Test
  void basicRulesRenderStableMessagesAndResolutions() {
    assertRendering(
        FieldValidationProblem.atField("value", FieldValidationBasicRule.NON_BLANK),
        "value must not be blank",
        "Provide a non-blank value for field 'value'.");
    assertRendering(
        FieldValidationProblem.atField("items", FieldValidationBasicRule.NON_EMPTY),
        "items must not be empty",
        "Provide at least one value in field 'items'.");
    assertRendering(
        FieldValidationProblem.atField("rows", FieldValidationBasicRule.ROWS_NON_EMPTY),
        "rows must not be empty",
        "Provide at least one non-empty row in field 'rows'.");
    assertRendering(
        FieldValidationProblem.atField("index", FieldValidationBasicRule.NON_NEGATIVE),
        "index must not be negative",
        "Provide a non-negative integer for field 'index'.");
    assertRendering(
        FieldValidationProblem.atField("count", FieldValidationBasicRule.GREATER_THAN_ZERO),
        "count must be greater than 0",
        "Provide an integer greater than 0 for field 'count'.");
    assertRendering(
        FieldValidationProblem.atField("offset", FieldValidationBasicRule.NON_ZERO),
        "offset must not be 0",
        "Provide a non-zero integer for field 'offset'.");
    assertRendering(
        FieldValidationProblem.atField("addresses", FieldValidationBasicRule.DUPLICATES),
        "addresses must not contain duplicates",
        "Remove duplicate values from field 'addresses'.");
    assertRendering(
        FieldValidationProblem.atField("matrix", FieldValidationBasicRule.EMPTY_ROWS),
        "matrix must not contain empty rows",
        "Ensure every row in field 'matrix' contains at least one cell value.");
  }

  @Test
  void boundRulesRenderStableMessagesAndResolutions() {
    assertRendering(
        FieldValidationProblem.atField(
            "rowIndex", FieldValidationBoundRule.ROW_INDEX_BOUNDS, "1048576"),
        "rowIndex must not exceed 1048576 (Excel row limit)",
        "Provide a row index between 0 and 1048576 for field 'rowIndex'.");
    assertRendering(
        FieldValidationProblem.atField(
            "columnIndex", FieldValidationBoundRule.COLUMN_INDEX_BOUNDS, "16384"),
        "columnIndex must not exceed 16384 (Excel column limit)",
        "Provide a column index between 0 and 16384 for field 'columnIndex'.");
    assertRendering(
        FieldValidationProblem.atField(
            "addresses", FieldValidationBoundRule.COLLECTION_SIZE_LIMIT, "250000", "250001"),
        "addresses must not exceed 250000 but was 250001",
        "Reduce field 'addresses' to at most 250000 values.");
    assertRendering(
        FieldValidationProblem.detached(
            "endRow", FieldValidationBoundRule.ORDERED_SPAN, "startRow"),
        "endRow must not be less than startRow",
        "Ensure 'endRow' is greater than or equal to 'startRow'.");
    assertRendering(
        FieldValidationProblem.atField("matrix", FieldValidationBoundRule.RECTANGULAR_MATRIX),
        "matrix must describe a rectangular matrix",
        "Ensure field 'matrix' describes a rectangular matrix with the same width in every row.");
    assertRendering(
        FieldValidationProblem.detached(
            "windowSize", FieldValidationBoundRule.WINDOW_SIZE_PRODUCT, "10000", "10001"),
        "windowSize must not exceed 10000 but was 10001",
        "Reduce rowCount and columnCount so their product does not exceed 10000.");
  }

  @Test
  void namingAndAddressRulesRenderStableMessagesAndResolutions() {
    assertRendering(
        FieldValidationProblem.atField(
            "sheetName",
            FieldValidationNamingRule.SHEET_NAME_TOO_LONG,
            "12345678901234567890123456789012"),
        "sheetName must not exceed 31 characters: 12345678901234567890123456789012",
        "Provide a valid Excel sheet name for field 'sheetName'.");
    assertRendering(
        FieldValidationProblem.atField(
            "sheetName", FieldValidationNamingRule.SHEET_NAME_BOUNDARY_QUOTE, "'Budget"),
        "sheetName must not begin or end with a single quote: 'Budget",
        "Provide a valid Excel sheet name for field 'sheetName'.");
    assertRendering(
        FieldValidationProblem.atField(
            "sheetName",
            FieldValidationNamingRule.SHEET_NAME_INVALID_CHARACTER,
            "/",
            "6",
            "Budget/2026"),
        "sheetName contains invalid Excel character / at position 6: Budget/2026",
        "Provide a valid Excel sheet name for field 'sheetName'.");
    assertRendering(
        FieldValidationProblem.atField(
            "definedName", FieldValidationNamingRule.DEFINED_NAME_TOO_LONG),
        "definedName must not exceed 255 Unicode code points",
        "Provide a valid Excel defined name for field 'definedName'.");
    assertRendering(
        FieldValidationProblem.atField(
            "definedName", FieldValidationNamingRule.DEFINED_NAME_SYNTAX),
        "definedName must start with a letter, underscore, or backslash and contain only Unicode"
            + " letters, Unicode numbers, underscore, period, or backslash",
        "Provide a valid Excel defined name for field 'definedName'.");
    assertRendering(
        FieldValidationProblem.atField(
            "definedName", FieldValidationNamingRule.DEFINED_NAME_RESERVED_PREFIX),
        "definedName must not use the reserved _xlnm. prefix",
        "Provide a valid Excel defined name for field 'definedName'.");
    assertRendering(
        FieldValidationProblem.atField(
            "definedName", FieldValidationNamingRule.DEFINED_NAME_A1_COLLISION),
        "definedName must not collide with A1-style cell reference syntax",
        "Provide a valid Excel defined name for field 'definedName'.");
    assertRendering(
        FieldValidationProblem.atField(
            "definedName", FieldValidationNamingRule.DEFINED_NAME_R1C1_COLLISION),
        "definedName must not collide with R1C1-style cell reference syntax",
        "Provide a valid Excel defined name for field 'definedName'.");
    assertRendering(
        FieldValidationProblem.atField("address", FieldValidationAddressRule.ADDRESS_SYNTAX),
        "address must be a single-cell A1-style address",
        "Use a single-cell A1-style address such as A1 or BC12 within Excel .xlsx bounds for"
            + " field 'address'.");
    assertRendering(
        FieldValidationProblem.atField("address", FieldValidationAddressRule.ADDRESS_BOUNDS),
        "address must be within Excel .xlsx bounds",
        "Use a single-cell A1-style address such as A1 or BC12 within Excel .xlsx bounds for"
            + " field 'address'.");
    assertRendering(
        FieldValidationProblem.atField(
            "range", FieldValidationAddressRule.RANGE_RECTANGULAR_SYNTAX),
        "range must be a rectangular A1-style range with at most one ':'",
        "Provide a rectangular A1-style range with at most one ':' for field 'range'.");
  }

  @Test
  void layoutRulesRenderStableMessagesAndResolutions() {
    assertRendering(
        FieldValidationProblem.atField("width", FieldValidationLayoutRule.FIELD_MUST_BE_FINITE),
        "width must be finite",
        "Provide a finite numeric value for field 'width'.");
    assertRendering(
        FieldValidationProblem.atField(
            "width", FieldValidationLayoutRule.COLUMN_WIDTH_TOO_LARGE, "255.0", "300.0"),
        "width must not exceed 255.0 (Excel column width limit): got 300.0",
        "Provide a visible Excel column width greater than 0 and no more than 255.0 characters"
            + " for field 'width'.");
    assertRendering(
        FieldValidationProblem.atField(
            "width", FieldValidationLayoutRule.COLUMN_WIDTH_NOT_VISIBLE, "255.0", "0.05"),
        "width is too small to produce a visible Excel column width: got 0.05",
        "Provide a visible Excel column width greater than 0 and no more than 255.0 characters"
            + " for field 'width'.");
    assertRendering(
        FieldValidationProblem.atField(
            "height", FieldValidationLayoutRule.ROW_HEIGHT_TOO_LARGE, "409.5", "500.0"),
        "height must not exceed 409.5 (Excel row height limit): got 500.0",
        "Provide a visible Excel row height greater than 0 and no more than 409.5 points for"
            + " field 'height'.");
    assertRendering(
        FieldValidationProblem.atField(
            "height", FieldValidationLayoutRule.ROW_HEIGHT_NOT_VISIBLE, "409.5", "0.02"),
        "height is too small to produce a visible Excel row height: 0.02",
        "Provide a visible Excel row height greater than 0 and no more than 409.5 points for"
            + " field 'height'.");
    assertRendering(
        FieldValidationProblem.atField(
            "zoomPercent", FieldValidationLayoutRule.ZOOM_PERCENT_RANGE, "10", "400", "401"),
        "zoomPercent must be between 10 and 400 inclusive: 401",
        "Provide a zoom percentage between 10 and 400 inclusive for field 'zoomPercent'.");
  }

  @Test
  void mappersCoverEveryTypedViolationBranch() {
    assertMapped(
        FieldValidationProblemMappers.sheetName("sheetName", ExcelSheetNameProblem.blank()),
        FieldValidationBasicRule.NON_BLANK,
        "sheetName must not be blank");
    assertMapped(
        FieldValidationProblemMappers.sheetName(
            "sheetName", ExcelSheetNameProblem.tooLong("12345678901234567890123456789012")),
        FieldValidationNamingRule.SHEET_NAME_TOO_LONG,
        "sheetName must not exceed 31 characters: 12345678901234567890123456789012");
    assertMapped(
        FieldValidationProblemMappers.sheetName(
            "sheetName", ExcelSheetNameProblem.boundaryQuote("'Budget")),
        FieldValidationNamingRule.SHEET_NAME_BOUNDARY_QUOTE,
        "sheetName must not begin or end with a single quote: 'Budget");
    assertMapped(
        FieldValidationProblemMappers.sheetName(
            "sheetName", ExcelSheetNameProblem.invalidCharacter("/", 6, "Budget/2026")),
        FieldValidationNamingRule.SHEET_NAME_INVALID_CHARACTER,
        "sheetName contains invalid Excel character / at position 6: Budget/2026");

    assertMapped(
        FieldValidationProblemMappers.definedName(
            "definedName", ProtocolDefinedNameValidation.Violation.BLANK),
        FieldValidationBasicRule.NON_BLANK,
        "definedName must not be blank");
    assertMapped(
        FieldValidationProblemMappers.definedName(
            "definedName", ProtocolDefinedNameValidation.Violation.TOO_LONG),
        FieldValidationNamingRule.DEFINED_NAME_TOO_LONG,
        "definedName must not exceed 255 Unicode code points");
    assertMapped(
        FieldValidationProblemMappers.definedName(
            "definedName", ProtocolDefinedNameValidation.Violation.SYNTAX),
        FieldValidationNamingRule.DEFINED_NAME_SYNTAX,
        "definedName must start with a letter, underscore, or backslash and contain only Unicode"
            + " letters, Unicode numbers, underscore, period, or backslash");
    assertMapped(
        FieldValidationProblemMappers.definedName(
            "definedName", ProtocolDefinedNameValidation.Violation.RESERVED_PREFIX),
        FieldValidationNamingRule.DEFINED_NAME_RESERVED_PREFIX,
        "definedName must not use the reserved _xlnm. prefix");
    assertMapped(
        FieldValidationProblemMappers.definedName(
            "definedName", ProtocolDefinedNameValidation.Violation.A1_COLLISION),
        FieldValidationNamingRule.DEFINED_NAME_A1_COLLISION,
        "definedName must not collide with A1-style cell reference syntax");
    assertMapped(
        FieldValidationProblemMappers.definedName(
            "definedName", ProtocolDefinedNameValidation.Violation.R1C1_COLLISION),
        FieldValidationNamingRule.DEFINED_NAME_R1C1_COLLISION,
        "definedName must not collide with R1C1-style cell reference syntax");

    assertMapped(
        FieldValidationProblemMappers.address(
            "address", ProtocolCellAddressValidation.Violation.BLANK),
        FieldValidationBasicRule.NON_BLANK,
        "address must not be blank");
    assertMapped(
        FieldValidationProblemMappers.address(
            "address", ProtocolCellAddressValidation.Violation.SYNTAX),
        FieldValidationAddressRule.ADDRESS_SYNTAX,
        "address must be a single-cell A1-style address");
    assertMapped(
        FieldValidationProblemMappers.address(
            "address", ProtocolCellAddressValidation.Violation.BOUNDS),
        FieldValidationAddressRule.ADDRESS_BOUNDS,
        "address must be within Excel .xlsx bounds");

    assertMapped(
        FieldValidationProblemMappers.pivotTableName("pivotTableName"),
        FieldValidationBasicRule.NON_BLANK,
        "pivotTableName must not be blank");

    assertMapped(
        FieldValidationProblemMappers.columnWidth(
            "widthCharacters", 20.0, ExcelColumnWidthViolation.NON_FINITE),
        FieldValidationLayoutRule.FIELD_MUST_BE_FINITE,
        "widthCharacters must be finite");
    assertMapped(
        FieldValidationProblemMappers.columnWidth(
            "widthCharacters", 0.0, ExcelColumnWidthViolation.NON_POSITIVE),
        FieldValidationBasicRule.GREATER_THAN_ZERO,
        "widthCharacters must be greater than 0");
    assertMapped(
        FieldValidationProblemMappers.columnWidth(
            "widthCharacters", 300.0, ExcelColumnWidthViolation.TOO_LARGE),
        FieldValidationLayoutRule.COLUMN_WIDTH_TOO_LARGE,
        "widthCharacters must not exceed 255.0 (Excel column width limit): got 300.0");
    assertMapped(
        FieldValidationProblemMappers.columnWidth(
            "widthCharacters", 0.05, ExcelColumnWidthViolation.NOT_VISIBLE),
        FieldValidationLayoutRule.COLUMN_WIDTH_NOT_VISIBLE,
        "widthCharacters is too small to produce a visible Excel column width: got 0.05");

    assertMapped(
        FieldValidationProblemMappers.rowHeight(
            "heightPoints", 20.0, ExcelRowHeightViolation.NON_FINITE),
        FieldValidationLayoutRule.FIELD_MUST_BE_FINITE,
        "heightPoints must be finite");
    assertMapped(
        FieldValidationProblemMappers.rowHeight(
            "heightPoints", 0.0, ExcelRowHeightViolation.NON_POSITIVE),
        FieldValidationBasicRule.GREATER_THAN_ZERO,
        "heightPoints must be greater than 0");
    assertMapped(
        FieldValidationProblemMappers.rowHeight(
            "heightPoints", 500.0, ExcelRowHeightViolation.TOO_LARGE),
        FieldValidationLayoutRule.ROW_HEIGHT_TOO_LARGE,
        "heightPoints must not exceed 409.0 (Excel row height limit): got 500.0");
    assertMapped(
        FieldValidationProblemMappers.rowHeight(
            "heightPoints", 0.02, ExcelRowHeightViolation.NOT_VISIBLE),
        FieldValidationLayoutRule.ROW_HEIGHT_NOT_VISIBLE,
        "heightPoints is too small to produce a visible Excel row height: 0.02");

    assertMapped(
        FieldValidationProblemMappers.zoomPercent(
            "zoomPercent", 401, ExcelZoomViolation.OUT_OF_RANGE),
        FieldValidationLayoutRule.ZOOM_PERCENT_RANGE,
        "zoomPercent must be between 10 and 400 inclusive: 401");
  }

  @Test
  void zoomMapperRejectsUnhandledViolation() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> FieldValidationProblemMappers.zoomPercent("zoomPercent", 125, null));

    assertEquals("Unhandled zoom violation: null", failure.getMessage());
  }

  @Test
  void requestProblemDescriptorSupportRebasesEveryDescriptorVariant() {
    RequestProblemDescriptor original = new MessageShape("shape", Optional.of("old"));
    assertSame(original, RequestProblemDescriptorSupport.withJsonPath(original, Optional.empty()));

    var missingField =
        assertInstanceOf(
            MissingRequiredField.class,
            RequestProblemDescriptorSupport.withJsonPath(
                new MissingRequiredField("steps[0].type"), Optional.of("steps[1].type")));
    assertEquals("steps[1].type", missingField.jsonPathValue());

    var missingType =
        assertInstanceOf(
            MissingTypeDiscriminator.class,
            RequestProblemDescriptorSupport.withJsonPath(
                new MissingTypeDiscriminator("selector.type"), Optional.of("selector.kind")));
    assertEquals("selector.kind", missingType.jsonPathValue());

    var unknownField =
        assertInstanceOf(
            UnknownField.class,
            RequestProblemDescriptorSupport.withJsonPath(
                new UnknownField("steps[0].extra"), Optional.of("steps[2].extra")));
    assertEquals("steps[2].extra", unknownField.jsonPathValue());

    var unknownTypeValue =
        assertInstanceOf(
            UnknownTypeValue.class,
            RequestProblemDescriptorSupport.withJsonPath(
                new UnknownTypeValue(
                    "by_magic",
                    Optional.of("selector.type"),
                    List.of("by_name"),
                    Optional.of("Use by_name.")),
                Optional.of("steps[3].selector.type")));
    assertEquals("by_magic", unknownTypeValue.typeId());
    assertEquals(Optional.of("steps[3].selector.type"), unknownTypeValue.jsonPath());
    assertEquals(List.of("by_name"), unknownTypeValue.similarValues());
    assertEquals(Optional.of("Use by_name."), unknownTypeValue.specificGuidance());

    var unsupportedValue =
        assertInstanceOf(
            UnsupportedValue.class,
            RequestProblemDescriptorSupport.withJsonPath(
                new UnsupportedValue("xlsxm", Optional.of("sourceKind"), List.of("xlsx")),
                Optional.of("steps[4].sourceKind")));
    assertEquals("xlsxm", unsupportedValue.value());
    assertEquals(Optional.of("steps[4].sourceKind"), unsupportedValue.jsonPath());
    assertEquals(List.of("xlsx"), unsupportedValue.allowedValues());

    var explicitNull =
        assertInstanceOf(
            ExplicitNullField.class,
            RequestProblemDescriptorSupport.withJsonPath(
                new ExplicitNullField("execution"), Optional.of("steps[5].execution")));
    assertEquals("steps[5].execution", explicitNull.jsonPathValue());

    var actionableShape =
        assertInstanceOf(
            ActionableShapeMessage.class,
            RequestProblemDescriptorSupport.withJsonPath(
                new ActionableShapeMessage("shape", "fix shape", Optional.of("old.path")),
                Optional.of("new.path")));
    assertEquals("shape", actionableShape.message());
    assertEquals("fix shape", actionableShape.resolutionValue());
    assertEquals(Optional.of("new.path"), actionableShape.jsonPath());

    var messageShape =
        assertInstanceOf(
            MessageShape.class,
            RequestProblemDescriptorSupport.withJsonPath(
                new MessageShape("shape-only", Optional.of("old.shape")),
                Optional.of("new.shape")));
    assertEquals("shape-only", messageShape.message());
    assertEquals(Optional.of("new.shape"), messageShape.jsonPath());

    var duplicateStepId =
        assertInstanceOf(
            DuplicateStepId.class,
            RequestProblemDescriptorSupport.withJsonPath(
                new DuplicateStepId("dup", "steps[0].stepId"), Optional.of("steps[9].stepId")));
    assertEquals("dup", duplicateStepId.value());
    assertEquals("steps[9].stepId", duplicateStepId.jsonPathValue());

    var nonXlsxPath =
        assertInstanceOf(
            NonXlsxPath.class,
            RequestProblemDescriptorSupport.withJsonPath(
                new NonXlsxPath(".csv", Optional.of("source.path")),
                Optional.of("steps[6].source.path")));
    assertEquals(".csv", nonXlsxPath.actualExtension());
    assertEquals(Optional.of("steps[6].source.path"), nonXlsxPath.jsonPath());

    var fieldValidation =
        assertInstanceOf(
            FieldValidationProblem.class,
            RequestProblemDescriptorSupport.withJsonPath(
                FieldValidationProblem.detached(
                    "zoomPercent",
                    FieldValidationLayoutRule.ZOOM_PERCENT_RANGE,
                    "10",
                    "400",
                    "401"),
                Optional.of("steps[7].zoomPercent")));
    assertEquals("zoomPercent", fieldValidation.fieldName());
    assertEquals(Optional.of("steps[7].zoomPercent"), fieldValidation.jsonPath());
    assertEquals(FieldValidationLayoutRule.ZOOM_PERCENT_RANGE, fieldValidation.rule());
    assertEquals(List.of("10", "400", "401"), fieldValidation.operands());

    var actionableInvariant =
        assertInstanceOf(
            ActionableInvariantMessage.class,
            RequestProblemDescriptorSupport.withJsonPath(
                new ActionableInvariantMessage(
                    "invariant", "repair invariant", Optional.of("old.invariant")),
                Optional.of("new.invariant")));
    assertEquals("invariant", actionableInvariant.message());
    assertEquals("repair invariant", actionableInvariant.resolutionValue());
    assertEquals(Optional.of("new.invariant"), actionableInvariant.jsonPath());

    var messageInvariant =
        assertInstanceOf(
            MessageInvariant.class,
            RequestProblemDescriptorSupport.withJsonPath(
                new MessageInvariant("plain invariant", Optional.of("old.message")),
                Optional.of("new.message")));
    assertEquals("plain invariant", messageInvariant.message());
    assertEquals(Optional.of("new.message"), messageInvariant.jsonPath());
  }

  private static void assertMapped(
      FieldValidationProblem problem, FieldValidationRule rule, String message) {
    assertEquals(Optional.of(problem.fieldName()), problem.jsonPath());
    assertEquals(rule, problem.rule());
    assertEquals(message, problem.message());
  }

  private static void assertRendering(
      FieldValidationProblem problem, String message, String resolution) {
    assertEquals(message, problem.message());
    assertEquals(resolution, problem.resolution());
  }
}
