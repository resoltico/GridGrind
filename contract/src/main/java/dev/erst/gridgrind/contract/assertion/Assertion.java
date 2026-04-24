package dev.erst.gridgrind.contract.assertion;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.contract.dto.AnalysisFindingCode;
import dev.erst.gridgrind.contract.dto.AnalysisSeverity;
import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import java.util.List;
import java.util.Objects;

/** First-class verification contract evaluated against a workbook target. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(
      value = Assertion.NamedRangePresent.class,
      name = "EXPECT_NAMED_RANGE_PRESENT"),
  @JsonSubTypes.Type(value = Assertion.NamedRangeAbsent.class, name = "EXPECT_NAMED_RANGE_ABSENT"),
  @JsonSubTypes.Type(value = Assertion.TablePresent.class, name = "EXPECT_TABLE_PRESENT"),
  @JsonSubTypes.Type(value = Assertion.TableAbsent.class, name = "EXPECT_TABLE_ABSENT"),
  @JsonSubTypes.Type(
      value = Assertion.PivotTablePresent.class,
      name = "EXPECT_PIVOT_TABLE_PRESENT"),
  @JsonSubTypes.Type(value = Assertion.PivotTableAbsent.class, name = "EXPECT_PIVOT_TABLE_ABSENT"),
  @JsonSubTypes.Type(value = Assertion.ChartPresent.class, name = "EXPECT_CHART_PRESENT"),
  @JsonSubTypes.Type(value = Assertion.ChartAbsent.class, name = "EXPECT_CHART_ABSENT"),
  @JsonSubTypes.Type(value = Assertion.CellValue.class, name = "EXPECT_CELL_VALUE"),
  @JsonSubTypes.Type(value = Assertion.DisplayValue.class, name = "EXPECT_DISPLAY_VALUE"),
  @JsonSubTypes.Type(value = Assertion.FormulaText.class, name = "EXPECT_FORMULA_TEXT"),
  @JsonSubTypes.Type(value = Assertion.CellStyle.class, name = "EXPECT_CELL_STYLE"),
  @JsonSubTypes.Type(
      value = Assertion.WorkbookProtectionFacts.class,
      name = "EXPECT_WORKBOOK_PROTECTION"),
  @JsonSubTypes.Type(value = Assertion.SheetStructureFacts.class, name = "EXPECT_SHEET_STRUCTURE"),
  @JsonSubTypes.Type(value = Assertion.NamedRangeFacts.class, name = "EXPECT_NAMED_RANGE_FACTS"),
  @JsonSubTypes.Type(value = Assertion.TableFacts.class, name = "EXPECT_TABLE_FACTS"),
  @JsonSubTypes.Type(value = Assertion.PivotTableFacts.class, name = "EXPECT_PIVOT_TABLE_FACTS"),
  @JsonSubTypes.Type(value = Assertion.ChartFacts.class, name = "EXPECT_CHART_FACTS"),
  @JsonSubTypes.Type(
      value = Assertion.AnalysisMaxSeverity.class,
      name = "EXPECT_ANALYSIS_MAX_SEVERITY"),
  @JsonSubTypes.Type(
      value = Assertion.AnalysisFindingPresent.class,
      name = "EXPECT_ANALYSIS_FINDING_PRESENT"),
  @JsonSubTypes.Type(
      value = Assertion.AnalysisFindingAbsent.class,
      name = "EXPECT_ANALYSIS_FINDING_ABSENT"),
  @JsonSubTypes.Type(value = Assertion.AllOf.class, name = "ALL_OF"),
  @JsonSubTypes.Type(value = Assertion.AnyOf.class, name = "ANY_OF"),
  @JsonSubTypes.Type(value = Assertion.Not.class, name = "NOT")
})
public sealed interface Assertion
    permits Assertion.NamedRangePresent,
        Assertion.NamedRangeAbsent,
        Assertion.TablePresent,
        Assertion.TableAbsent,
        Assertion.PivotTablePresent,
        Assertion.PivotTableAbsent,
        Assertion.ChartPresent,
        Assertion.ChartAbsent,
        Assertion.CellValue,
        Assertion.DisplayValue,
        Assertion.FormulaText,
        Assertion.CellStyle,
        Assertion.WorkbookProtectionFacts,
        Assertion.SheetStructureFacts,
        Assertion.NamedRangeFacts,
        Assertion.TableFacts,
        Assertion.PivotTableFacts,
        Assertion.ChartFacts,
        Assertion.AnalysisMaxSeverity,
        Assertion.AnalysisFindingPresent,
        Assertion.AnalysisFindingAbsent,
        Assertion.AllOf,
        Assertion.AnyOf,
        Assertion.Not {

  /** Stable SCREAMING_SNAKE_CASE discriminator mirrored in catalog and result surfaces. */
  default String assertionType() {
    return GridGrindProtocolTypeNames.assertionTypeName(getClass().asSubclass(Assertion.class));
  }

  /** Expects the named-range selector to resolve to one or more matching named ranges. */
  record NamedRangePresent() implements Assertion {}

  /** Expects the named-range selector to resolve to no matching named ranges. */
  record NamedRangeAbsent() implements Assertion {}

  /** Expects the table selector to resolve to one or more matching tables. */
  record TablePresent() implements Assertion {}

  /** Expects the table selector to resolve to no matching tables. */
  record TableAbsent() implements Assertion {}

  /** Expects the pivot-table selector to resolve to one or more matching pivot tables. */
  record PivotTablePresent() implements Assertion {}

  /** Expects the pivot-table selector to resolve to no matching pivot tables. */
  record PivotTableAbsent() implements Assertion {}

  /** Expects the chart selector to resolve to one or more matching charts. */
  record ChartPresent() implements Assertion {}

  /** Expects the chart selector to resolve to no matching charts. */
  record ChartAbsent() implements Assertion {}

  /** Expects every selected cell to have the exact effective value. */
  record CellValue(ExpectedCellValue expectedValue) implements Assertion {
    public CellValue {
      Objects.requireNonNull(expectedValue, "expectedValue must not be null");
    }
  }

  /** Expects every selected cell to have the exact formatted display string. */
  record DisplayValue(String displayValue) implements Assertion {
    public DisplayValue {
      Objects.requireNonNull(displayValue, "displayValue must not be null");
    }
  }

  /** Expects every selected cell to have the exact formula text. */
  record FormulaText(String formula) implements Assertion {
    public FormulaText {
      formula = AssertionSupport.requireNonBlank(formula, "formula");
    }
  }

  /** Expects every selected cell to have the exact style snapshot. */
  record CellStyle(GridGrindResponse.CellStyleReport style) implements Assertion {
    public CellStyle {
      Objects.requireNonNull(style, "style must not be null");
    }
  }

  /** Expects the current workbook protection facts to match exactly. */
  record WorkbookProtectionFacts(WorkbookProtectionReport protection) implements Assertion {
    public WorkbookProtectionFacts {
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }

  /** Expects the selected sheet summary facts to match exactly. */
  record SheetStructureFacts(GridGrindResponse.SheetSummaryReport sheet) implements Assertion {
    public SheetStructureFacts {
      Objects.requireNonNull(sheet, "sheet must not be null");
    }
  }

  /** Expects the selected named-range facts to match exactly and in order. */
  record NamedRangeFacts(List<GridGrindResponse.NamedRangeReport> namedRanges)
      implements Assertion {
    public NamedRangeFacts {
      namedRanges = AssertionSupport.copyNamedRanges(namedRanges, "namedRanges");
    }
  }

  /** Expects the selected table facts to match exactly and in order. */
  record TableFacts(List<TableEntryReport> tables) implements Assertion {
    public TableFacts {
      tables = AssertionSupport.copyTables(tables, "tables");
    }
  }

  /** Expects the selected pivot-table facts to match exactly and in order. */
  record PivotTableFacts(List<PivotTableReport> pivotTables) implements Assertion {
    public PivotTableFacts {
      pivotTables = AssertionSupport.copyPivotTables(pivotTables, "pivotTables");
    }
  }

  /** Expects the selected chart facts to match exactly and in order. */
  record ChartFacts(List<ChartReport> charts) implements Assertion {
    public ChartFacts {
      charts = AssertionSupport.copyCharts(charts, "charts");
    }
  }

  /** Expects one analysis query to report a maximum severity no higher than the supplied level. */
  record AnalysisMaxSeverity(InspectionQuery query, AnalysisSeverity maximumSeverity)
      implements Assertion {
    public AnalysisMaxSeverity {
      query = AssertionSupport.requireAnalysisQuery(query, "query");
      maximumSeverity = AssertionSupport.requireSeverity(maximumSeverity, "maximumSeverity");
    }
  }

  /** Expects one analysis query to report at least one matching finding. */
  record AnalysisFindingPresent(
      InspectionQuery query,
      AnalysisFindingCode code,
      @JsonInclude(JsonInclude.Include.NON_NULL) AnalysisSeverity severity,
      @JsonInclude(JsonInclude.Include.NON_NULL) String messageContains)
      implements Assertion {
    public AnalysisFindingPresent {
      query = AssertionSupport.requireAnalysisQuery(query, "query");
      code = AssertionSupport.requireFindingCode(code, "code");
      if (messageContains != null && messageContains.isBlank()) {
        throw new IllegalArgumentException("messageContains must not be blank");
      }
    }
  }

  /** Expects one analysis query to report no matching finding. */
  record AnalysisFindingAbsent(
      InspectionQuery query,
      AnalysisFindingCode code,
      @JsonInclude(JsonInclude.Include.NON_NULL) AnalysisSeverity severity,
      @JsonInclude(JsonInclude.Include.NON_NULL) String messageContains)
      implements Assertion {
    public AnalysisFindingAbsent {
      query = AssertionSupport.requireAnalysisQuery(query, "query");
      code = AssertionSupport.requireFindingCode(code, "code");
      if (messageContains != null && messageContains.isBlank()) {
        throw new IllegalArgumentException("messageContains must not be blank");
      }
    }
  }

  /** Requires every nested assertion to pass. */
  record AllOf(List<Assertion> assertions) implements Assertion {
    public AllOf {
      assertions = AssertionSupport.copyAssertions(assertions, "assertions");
    }
  }

  /** Requires at least one nested assertion to pass. */
  record AnyOf(List<Assertion> assertions) implements Assertion {
    public AnyOf {
      assertions = AssertionSupport.copyAssertions(assertions, "assertions");
    }
  }

  /** Inverts the result of one nested assertion. */
  record Not(Assertion assertion) implements Assertion {
    public Not {
      Objects.requireNonNull(assertion, "assertion must not be null");
    }
  }
}
