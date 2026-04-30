package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Optional;

/** Workbook and step location facts reused across problem-context stages. */
public interface ProblemContextWorkbookSurfaces {
  /** Source-backed authored input reference without nullable kind/path slots. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = InputReference.Unknown.class, name = "UNKNOWN"),
    @JsonSubTypes.Type(value = InputReference.KindOnly.class, name = "KIND"),
    @JsonSubTypes.Type(value = InputReference.PathReference.class, name = "PATH")
  })
  sealed interface InputReference
      permits InputReference.Unknown, InputReference.KindOnly, InputReference.PathReference {
    /** Returns the explicit unknown input-reference variant. */
    static InputReference unknown() {
      return new InputReference.Unknown();
    }

    /** Returns the input-reference variant with one authored input kind and no path. */
    static InputReference kind(String inputKind) {
      return new KindOnly(requireNonBlank(inputKind, "inputKind"));
    }

    /** Returns the input-reference variant with one authored input kind and concrete path. */
    static InputReference path(String inputKind, String inputPath) {
      return new PathReference(
          requireNonBlank(inputKind, "inputKind"), requireNonBlank(inputPath, "inputPath"));
    }

    /** Returns the authored input kind when the failing binding was identified. */
    default Optional<String> inputKindValue() {
      return switch (this) {
        case KindOnly reference -> Optional.of(reference.inputKind());
        case PathReference reference -> Optional.of(reference.inputKind());
        case InputReference.Unknown _ -> Optional.empty();
      };
    }

    /** Returns the authored input path when the failing binding was identified. */
    default Optional<String> inputPathValue() {
      return switch (this) {
        case PathReference reference -> Optional.of(reference.inputPath());
        case KindOnly _ -> Optional.empty();
        case InputReference.Unknown _ -> Optional.empty();
      };
    }

    /** No input-binding path could be derived for the resolution failure. */
    record Unknown() implements InputReference {}

    /** The failing authored input kind was known, but no concrete path applied. */
    record KindOnly(String inputKind) implements InputReference {
      public KindOnly {
        inputKind = requireNonBlank(inputKind, "inputKind");
      }
    }

    /** One concrete authored input binding resolved to one path. */
    record PathReference(String inputKind, String inputPath) implements InputReference {
      public PathReference {
        inputKind = requireNonBlank(inputKind, "inputKind");
        inputPath = requireNonBlank(inputPath, "inputPath");
      }
    }
  }

  /** Workbook open target without nullable source-path fields. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = WorkbookReference.NewWorkbook.class, name = "NEW"),
    @JsonSubTypes.Type(value = WorkbookReference.ExistingFile.class, name = "EXISTING")
  })
  sealed interface WorkbookReference
      permits WorkbookReference.NewWorkbook, WorkbookReference.ExistingFile {
    /** Returns the workbook-reference variant for a brand-new in-memory workbook. */
    static WorkbookReference newWorkbook() {
      return new NewWorkbook();
    }

    /** Returns the workbook-reference variant for one existing `.xlsx` file path. */
    static WorkbookReference existingFile(String path) {
      return new ExistingFile(requireNonBlank(path, "path"));
    }

    /** Returns the source workbook path when execution opened one existing file. */
    default Optional<String> sourceWorkbookPathValue() {
      return switch (this) {
        case ExistingFile existingFile -> Optional.of(existingFile.path());
        case NewWorkbook _ -> Optional.empty();
      };
    }

    /** Workbook creation started from a brand-new in-memory workbook. */
    record NewWorkbook() implements WorkbookReference {}

    /** Workbook opening targeted one existing `.xlsx` file. */
    record ExistingFile(String path) implements WorkbookReference {
      public ExistingFile {
        path = requireNonBlank(path, "path");
      }
    }
  }

  /** Persistence attempt reference without nullable source/output path padding. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = PersistenceReference.OverwriteSource.class, name = "OVERWRITE"),
    @JsonSubTypes.Type(value = PersistenceReference.SaveAs.class, name = "SAVE_AS")
  })
  sealed interface PersistenceReference
      permits PersistenceReference.OverwriteSource, PersistenceReference.SaveAs {
    /** Returns the persistence-reference variant that targets the opened source workbook. */
    static PersistenceReference overwriteSource(String sourceWorkbookPath) {
      return new OverwriteSource(requireNonBlank(sourceWorkbookPath, "sourceWorkbookPath"));
    }

    /** Returns the persistence-reference variant that targets one explicit save-as path. */
    static PersistenceReference saveAs(String persistencePath) {
      return new SaveAs(requireNonBlank(persistencePath, "persistencePath"));
    }

    /** Returns the overwritten source path when persistence targeted the opened workbook. */
    default Optional<String> sourceWorkbookPathValue() {
      return switch (this) {
        case OverwriteSource overwriteSource -> Optional.of(overwriteSource.sourceWorkbookPath());
        case SaveAs _ -> Optional.empty();
      };
    }

    /** Returns the explicit SAVE_AS destination path when persistence targeted one new file. */
    default Optional<String> persistencePathValue() {
      return switch (this) {
        case SaveAs saveAs -> Optional.of(saveAs.persistencePath());
        case OverwriteSource _ -> Optional.empty();
      };
    }

    /** Workbook persistence attempted to overwrite the opened source file. */
    record OverwriteSource(String sourceWorkbookPath) implements PersistenceReference {
      public OverwriteSource {
        sourceWorkbookPath = requireNonBlank(sourceWorkbookPath, "sourceWorkbookPath");
      }
    }

    /** Workbook persistence attempted to write one explicit SAVE_AS path. */
    record SaveAs(String persistencePath) implements PersistenceReference {
      public SaveAs {
        persistencePath = requireNonBlank(persistencePath, "persistencePath");
      }
    }
  }

  /** Workbook location reference used by calculation and step failures without null padding. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = ProblemLocation.Unknown.class, name = "UNKNOWN"),
    @JsonSubTypes.Type(value = ProblemLocation.Sheet.class, name = "SHEET"),
    @JsonSubTypes.Type(value = ProblemLocation.Address.class, name = "ADDRESS"),
    @JsonSubTypes.Type(value = ProblemLocation.Cell.class, name = "CELL"),
    @JsonSubTypes.Type(value = ProblemLocation.RangeOnly.class, name = "RANGE_ONLY"),
    @JsonSubTypes.Type(value = ProblemLocation.Range.class, name = "RANGE"),
    @JsonSubTypes.Type(value = ProblemLocation.NamedRange.class, name = "NAMED_RANGE"),
    @JsonSubTypes.Type(value = ProblemLocation.SheetNamedRange.class, name = "SHEET_NAMED_RANGE"),
    @JsonSubTypes.Type(value = ProblemLocation.FormulaCell.class, name = "FORMULA_CELL")
  })
  sealed interface ProblemLocation
      permits ProblemLocation.Unknown,
          ProblemLocation.Sheet,
          ProblemLocation.Address,
          ProblemLocation.Cell,
          ProblemLocation.RangeOnly,
          ProblemLocation.Range,
          ProblemLocation.NamedRange,
          ProblemLocation.SheetNamedRange,
          ProblemLocation.FormulaCell {
    /** Returns the explicit unknown workbook-location variant. */
    static ProblemLocation unknown() {
      return new ProblemLocation.Unknown();
    }

    /** Returns the workbook-location variant for one whole sheet. */
    static ProblemLocation sheet(String sheetName) {
      return new Sheet(requireNonBlank(sheetName, "sheetName"));
    }

    /** Returns the workbook-location variant for one address without a resolved sheet. */
    static ProblemLocation address(String address) {
      return new Address(requireNonBlank(address, "address"));
    }

    /** Returns the workbook-location variant for one concrete sheet-backed cell. */
    static ProblemLocation cell(String sheetName, String address) {
      return new Cell(requireNonBlank(sheetName, "sheetName"), requireNonBlank(address, "address"));
    }

    /** Returns the workbook-location variant for one range without a resolved sheet. */
    static ProblemLocation range(String range) {
      return new RangeOnly(requireNonBlank(range, "range"));
    }

    /** Returns the workbook-location variant for one concrete sheet-backed range. */
    static ProblemLocation range(String sheetName, String range) {
      return new Range(requireNonBlank(sheetName, "sheetName"), requireNonBlank(range, "range"));
    }

    /** Returns the workbook-location variant for one workbook-scoped named range. */
    static ProblemLocation namedRange(String name) {
      return new NamedRange(requireNonBlank(name, "name"));
    }

    /** Returns the workbook-location variant for one sheet-scoped named range. */
    static ProblemLocation namedRange(String sheetName, String name) {
      return new SheetNamedRange(
          requireNonBlank(sheetName, "sheetName"), requireNonBlank(name, "name"));
    }

    /** Returns the workbook-location variant for one formula-bearing cell. */
    static ProblemLocation formulaCell(String sheetName, String address, String formula) {
      return new FormulaCell(
          requireNonBlank(sheetName, "sheetName"),
          requireNonBlank(address, "address"),
          requireNonBlank(formula, "formula"));
    }

    /** Returns the sheet name when the failure was localized to one sheet-backed target. */
    default Optional<String> sheetNameValue() {
      return switch (this) {
        case Sheet sheet -> Optional.of(sheet.sheetName());
        case Address _ -> Optional.empty();
        case Cell cell -> Optional.of(cell.sheetName());
        case RangeOnly _ -> Optional.empty();
        case Range range -> Optional.of(range.sheetName());
        case NamedRange _ -> Optional.empty();
        case SheetNamedRange namedRange -> Optional.of(namedRange.sheetName());
        case FormulaCell formulaCell -> Optional.of(formulaCell.sheetName());
        case ProblemLocation.Unknown _ -> Optional.empty();
      };
    }

    /** Returns the cell address when the failure was localized to one concrete cell. */
    default Optional<String> addressValue() {
      return switch (this) {
        case Address address -> Optional.of(address.address());
        case Cell cell -> Optional.of(cell.address());
        case FormulaCell formulaCell -> Optional.of(formulaCell.address());
        default -> Optional.empty();
      };
    }

    /** Returns the range when the failure was localized to one concrete rectangular range. */
    default Optional<String> rangeValue() {
      return switch (this) {
        case RangeOnly range -> Optional.of(range.range());
        case Range range -> Optional.of(range.range());
        default -> Optional.empty();
      };
    }

    /** Returns the named-range identifier when the failure was localized to one named range. */
    default Optional<String> namedRangeNameValue() {
      return switch (this) {
        case NamedRange namedRange -> Optional.of(namedRange.name());
        case SheetNamedRange namedRange -> Optional.of(namedRange.name());
        default -> Optional.empty();
      };
    }

    /** Returns the formula when the failure was localized to one formula-bearing cell. */
    default Optional<String> formulaValue() {
      if (this instanceof FormulaCell formulaCell) {
        return Optional.of(formulaCell.formula());
      }
      return Optional.empty();
    }

    /** The failure could not be localized to one workbook target. */
    record Unknown() implements ProblemLocation {}

    /** The failure was localized to one whole sheet. */
    record Sheet(String sheetName) implements ProblemLocation {
      public Sheet {
        sheetName = requireNonBlank(sheetName, "sheetName");
      }
    }

    /** The failure was localized to one concrete address without a resolved sheet. */
    record Address(String address) implements ProblemLocation {
      public Address {
        address = requireNonBlank(address, "address");
      }
    }

    /** The failure was localized to one concrete cell. */
    record Cell(String sheetName, String address) implements ProblemLocation {
      public Cell {
        sheetName = requireNonBlank(sheetName, "sheetName");
        address = requireNonBlank(address, "address");
      }
    }

    /** The failure was localized to one concrete range without a resolved sheet. */
    record RangeOnly(String range) implements ProblemLocation {
      public RangeOnly {
        range = requireNonBlank(range, "range");
      }
    }

    /** The failure was localized to one concrete rectangular range. */
    record Range(String sheetName, String range) implements ProblemLocation {
      public Range {
        sheetName = requireNonBlank(sheetName, "sheetName");
        range = requireNonBlank(range, "range");
      }
    }

    /** The failure was localized to one named range. */
    record NamedRange(String name) implements ProblemLocation {
      public NamedRange {
        name = requireNonBlank(name, "name");
      }
    }

    /** The failure was localized to one named range scoped to one sheet. */
    record SheetNamedRange(String sheetName, String name) implements ProblemLocation {
      public SheetNamedRange {
        sheetName = requireNonBlank(sheetName, "sheetName");
        name = requireNonBlank(name, "name");
      }
    }

    /** The failure was localized to one formula-bearing cell. */
    record FormulaCell(String sheetName, String address, String formula)
        implements ProblemLocation {
      public FormulaCell {
        sheetName = requireNonBlank(sheetName, "sheetName");
        address = requireNonBlank(address, "address");
        formula = requireNonBlank(formula, "formula");
      }
    }
  }

  /** Stable step identity captured for step execution failures. */
  record StepReference(int stepIndex, String stepId, String stepKind, String stepType) {
    public StepReference {
      if (stepIndex < 0) {
        throw new IllegalArgumentException("stepIndex must be greater than or equal to 0");
      }
      stepId = requireNonBlank(stepId, "stepId");
      stepKind = requireNonBlank(stepKind, "stepKind");
      stepType = requireNonBlank(stepType, "stepType");
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    return ProblemContextSupport.requireNonBlank(value, fieldName);
  }
}
