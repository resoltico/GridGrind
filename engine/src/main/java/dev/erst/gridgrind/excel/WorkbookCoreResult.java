package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.WorkbookResultSupport.copyDistinctStrings;
import static dev.erst.gridgrind.excel.WorkbookResultSupport.copyValues;
import static dev.erst.gridgrind.excel.WorkbookResultSupport.requireNonBlank;
import static dev.erst.gridgrind.excel.WorkbookResultSupport.validateCommonWorkbookSummaryFields;

import java.util.List;
import java.util.Objects;

/** Workbook-level fact results and summary payloads. */
public sealed interface WorkbookCoreResult extends WorkbookReadIntrospectionResult
    permits WorkbookCoreResult.WorkbookSummaryResult,
        WorkbookCoreResult.PackageSecurityResult,
        WorkbookCoreResult.WorkbookProtectionResult,
        WorkbookCoreResult.CustomXmlMappingsResult,
        WorkbookCoreResult.CustomXmlExportResult,
        WorkbookCoreResult.NamedRangesResult {

  /** Returns workbook-level summary facts. */
  record WorkbookSummaryResult(String stepId, WorkbookSummary workbook)
      implements WorkbookCoreResult {
    public WorkbookSummaryResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(workbook, "workbook must not be null");
    }
  }

  /** Returns OOXML package-encryption and package-signature facts. */
  record PackageSecurityResult(String stepId, ExcelOoxmlPackageSecuritySnapshot security)
      implements WorkbookCoreResult {
    public PackageSecurityResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(security, "security must not be null");
    }
  }

  /** Returns workbook-level protection facts. */
  record WorkbookProtectionResult(String stepId, ExcelWorkbookProtectionSnapshot protection)
      implements WorkbookCoreResult {
    public WorkbookProtectionResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(protection, "protection must not be null");
    }
  }

  /** Returns factual workbook custom-XML mapping metadata. */
  record CustomXmlMappingsResult(String stepId, List<ExcelCustomXmlMappingSnapshot> mappings)
      implements WorkbookCoreResult {
    public CustomXmlMappingsResult {
      stepId = requireNonBlank(stepId, "stepId");
      mappings = copyValues(mappings, "mappings");
    }
  }

  /** Returns XML exported from one selected workbook custom-XML mapping. */
  record CustomXmlExportResult(String stepId, ExcelCustomXmlExportSnapshot export)
      implements WorkbookCoreResult {
    public CustomXmlExportResult {
      stepId = requireNonBlank(stepId, "stepId");
      Objects.requireNonNull(export, "export must not be null");
    }
  }

  /** Returns selected named ranges. */
  record NamedRangesResult(String stepId, List<ExcelNamedRangeSnapshot> namedRanges)
      implements WorkbookCoreResult {
    public NamedRangesResult {
      stepId = requireNonBlank(stepId, "stepId");
      namedRanges = copyValues(namedRanges, "namedRanges");
    }
  }

  /** Workbook-level summary facts captured after all mutations complete. */
  sealed interface WorkbookSummary permits WorkbookSummary.Empty, WorkbookSummary.WithSheets {
    /** Total sheet count after all mutations complete. */
    int sheetCount();

    /** Ordered workbook sheet names. */
    List<String> sheetNames();

    /** Count of exposed named ranges after all mutations complete. */
    int namedRangeCount();

    /** Whether the workbook is marked to recalculate formulas on open. */
    boolean forceFormulaRecalculationOnOpen();

    /** Workbook summary for a zero-sheet workbook. */
    record Empty(
        int sheetCount,
        List<String> sheetNames,
        int namedRangeCount,
        boolean forceFormulaRecalculationOnOpen)
        implements WorkbookSummary {
      public Empty {
        sheetNames = validateCommonWorkbookSummaryFields(sheetCount, sheetNames, namedRangeCount);
        if (sheetCount != 0) {
          throw new IllegalArgumentException("sheetCount must be 0 for an empty workbook");
        }
      }
    }

    /** Workbook summary for a workbook that contains one or more sheets. */
    record WithSheets(
        int sheetCount,
        List<String> sheetNames,
        String activeSheetName,
        List<String> selectedSheetNames,
        int namedRangeCount,
        boolean forceFormulaRecalculationOnOpen)
        implements WorkbookSummary {
      public WithSheets {
        sheetNames = validateCommonWorkbookSummaryFields(sheetCount, sheetNames, namedRangeCount);
        activeSheetName = requireNonBlank(activeSheetName, "activeSheetName");
        selectedSheetNames = copyDistinctStrings(selectedSheetNames, "selectedSheetNames");
        if (sheetCount == 0) {
          throw new IllegalArgumentException("sheetCount must be greater than 0");
        }
        if (!sheetNames.contains(activeSheetName)) {
          throw new IllegalArgumentException("activeSheetName must be present in sheetNames");
        }
        if (selectedSheetNames.isEmpty()) {
          throw new IllegalArgumentException("selectedSheetNames must not be empty");
        }
        for (String selectedSheetName : selectedSheetNames) {
          if (!sheetNames.contains(selectedSheetName)) {
            throw new IllegalArgumentException(
                "selectedSheetNames must only contain values present in sheetNames");
          }
        }
      }
    }
  }
}
