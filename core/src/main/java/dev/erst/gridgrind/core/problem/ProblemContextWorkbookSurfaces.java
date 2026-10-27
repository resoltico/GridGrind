package dev.erst.gridgrind.core.problem;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Optional;

/** Workbook and step location facts reused across problem-context stages. */
public interface ProblemContextWorkbookSurfaces {
  /** Source-backed authored input reference without nullable kind/path slots. */
  @Union(discriminator = "type")
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
    @Variant(
        id = "UNKNOWN",
        summary = "No input-binding path could be derived for the resolution failure.")
    record Unknown() implements InputReference {}

    /** The failing authored input kind was known, but no concrete path applied. */
    @Variant(
        id = "KIND",
        summary = "The failing authored input kind was known, but no concrete path applied.")
    record KindOnly(String inputKind) implements InputReference {
      public KindOnly {
        inputKind = requireNonBlank(inputKind, "inputKind");
      }
    }

    /** One concrete authored input binding resolved to one path. */
    @Variant(id = "PATH", summary = "One concrete authored input binding resolved to one path.")
    record PathReference(String inputKind, String inputPath) implements InputReference {
      public PathReference {
        inputKind = requireNonBlank(inputKind, "inputKind");
        inputPath = requireNonBlank(inputPath, "inputPath");
      }
    }
  }

  /** Workbook open target without nullable source-path fields. */
  @Union(discriminator = "type")
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
    @Variant(id = "NEW", summary = "Workbook creation started from a brand-new in-memory workbook.")
    record NewWorkbook() implements WorkbookReference {}

    /** Workbook opening targeted one existing `.xlsx` file. */
    @Variant(id = "EXISTING", summary = "Workbook opening targeted one existing `.xlsx` file.")
    record ExistingFile(String path) implements WorkbookReference {
      public ExistingFile {
        path = requireNonBlank(path, "path");
      }
    }
  }

  /** Persistence attempt reference without nullable source/output path padding. */
  @Union(discriminator = "type")
  sealed interface PersistenceReference
      permits PersistenceReference.Overwrite, PersistenceReference.SaveAs {
    /** Returns the persistence-reference variant that targets the opened source workbook. */
    static PersistenceReference overwrite(String sourceWorkbookPath) {
      return new Overwrite(requireNonBlank(sourceWorkbookPath, "sourceWorkbookPath"));
    }

    /** Returns the persistence-reference variant that targets one explicit save-as path. */
    static PersistenceReference saveAs(String persistencePath) {
      return new SaveAs(requireNonBlank(persistencePath, "persistencePath"));
    }

    /** Returns the overwritten source path when persistence targeted the opened workbook. */
    default Optional<String> sourceWorkbookPathValue() {
      return switch (this) {
        case Overwrite overwriteSource -> Optional.of(overwriteSource.sourceWorkbookPath());
        case SaveAs _ -> Optional.empty();
      };
    }

    /** Returns the explicit SAVE_AS destination path when persistence targeted one new file. */
    default Optional<String> persistencePathValue() {
      return switch (this) {
        case SaveAs saveAs -> Optional.of(saveAs.persistencePath());
        case Overwrite _ -> Optional.empty();
      };
    }

    /** Workbook persistence attempted to overwrite the opened source file. */
    @Variant(
        id = "OVERWRITE",
        summary = "Workbook persistence attempted to overwrite the opened source file.")
    record Overwrite(String sourceWorkbookPath) implements PersistenceReference {
      public Overwrite {
        sourceWorkbookPath = requireNonBlank(sourceWorkbookPath, "sourceWorkbookPath");
      }
    }

    /** Workbook persistence attempted to write one explicit SAVE_AS path. */
    @Variant(
        id = "SAVE_AS",
        summary = "Workbook persistence attempted to write one explicit SAVE_AS path.")
    record SaveAs(String persistencePath) implements PersistenceReference {
      public SaveAs {
        persistencePath = requireNonBlank(persistencePath, "persistencePath");
      }
    }
  }

  /** Workbook location reference used by calculation and step failures without null padding. */
  @Union(discriminator = "type")
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
    static ProblemLocation sheet(SheetLabel sheetName) {
      return new Sheet(java.util.Objects.requireNonNull(sheetName, "sheetName must not be null"));
    }

    /** Returns the workbook-location variant for one address without a resolved sheet. */
    static ProblemLocation address(String address) {
      return new Address(requireNonBlank(address, "address"));
    }

    /** Returns the workbook-location variant for one concrete sheet-backed cell. */
    static ProblemLocation cell(SheetLabel sheetName, String address) {
      return new Cell(
          java.util.Objects.requireNonNull(sheetName, "sheetName must not be null"),
          requireNonBlank(address, "address"));
    }

    /** Returns the workbook-location variant for one range without a resolved sheet. */
    static ProblemLocation range(String range) {
      return new RangeOnly(requireNonBlank(range, "range"));
    }

    /** Returns the workbook-location variant for one concrete sheet-backed range. */
    static ProblemLocation range(SheetLabel sheetName, String range) {
      return new Range(
          java.util.Objects.requireNonNull(sheetName, "sheetName must not be null"),
          requireNonBlank(range, "range"));
    }

    /** Returns the workbook-location variant for one workbook-scoped named range. */
    static ProblemLocation namedRange(String name) {
      return new NamedRange(requireNonBlank(name, "name"));
    }

    /** Returns the workbook-location variant for one sheet-scoped named range. */
    static ProblemLocation namedRange(SheetLabel sheetName, String name) {
      return new SheetNamedRange(
          java.util.Objects.requireNonNull(sheetName, "sheetName must not be null"),
          requireNonBlank(name, "name"));
    }

    /** Returns the workbook-location variant for one formula-bearing cell. */
    static ProblemLocation formulaCell(SheetLabel sheetName, String address, String formula) {
      return new FormulaCell(
          java.util.Objects.requireNonNull(sheetName, "sheetName must not be null"),
          requireNonBlank(address, "address"),
          requireNonBlank(formula, "formula"));
    }

    /** Returns the sheet name when the failure was localized to one sheet-backed target. */
    default Optional<SheetLabel> sheetNameValue() {
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
    @Variant(id = "UNKNOWN", summary = "The failure could not be localized to one workbook target.")
    record Unknown() implements ProblemLocation {}

    /** The failure was localized to one whole sheet. */
    @Variant(id = "SHEET", summary = "The failure was localized to one whole sheet.")
    record Sheet(SheetLabel sheetName) implements ProblemLocation {
      public Sheet {
        java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
      }
    }

    /** The failure was localized to one concrete address without a resolved sheet. */
    @Variant(
        id = "ADDRESS",
        summary = "The failure was localized to one concrete address without a resolved sheet.")
    record Address(String address) implements ProblemLocation {
      public Address {
        address = requireNonBlank(address, "address");
      }
    }

    /** The failure was localized to one concrete cell. */
    @Variant(id = "CELL", summary = "The failure was localized to one concrete cell.")
    record Cell(SheetLabel sheetName, String address) implements ProblemLocation {
      public Cell {
        java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
        address = requireNonBlank(address, "address");
      }
    }

    /** The failure was localized to one concrete range without a resolved sheet. */
    @Variant(
        id = "RANGE_ONLY",
        summary = "The failure was localized to one concrete range without a resolved sheet.")
    record RangeOnly(String range) implements ProblemLocation {
      public RangeOnly {
        range = requireNonBlank(range, "range");
      }
    }

    /** The failure was localized to one concrete rectangular range. */
    @Variant(id = "RANGE", summary = "The failure was localized to one concrete rectangular range.")
    record Range(SheetLabel sheetName, String range) implements ProblemLocation {
      public Range {
        java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
        range = requireNonBlank(range, "range");
      }
    }

    /** The failure was localized to one named range. */
    @Variant(id = "NAMED_RANGE", summary = "The failure was localized to one named range.")
    record NamedRange(String name) implements ProblemLocation {
      public NamedRange {
        name = requireNonBlank(name, "name");
      }
    }

    /** The failure was localized to one named range scoped to one sheet. */
    @Variant(
        id = "SHEET_NAMED_RANGE",
        summary = "The failure was localized to one named range scoped to one sheet.")
    record SheetNamedRange(SheetLabel sheetName, String name) implements ProblemLocation {
      public SheetNamedRange {
        java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
        name = requireNonBlank(name, "name");
      }
    }

    /** The failure was localized to one formula-bearing cell. */
    @Variant(
        id = "FORMULA_CELL",
        summary = "The failure was localized to one formula-bearing cell.")
    record FormulaCell(SheetLabel sheetName, String address, String formula)
        implements ProblemLocation {
      public FormulaCell {
        java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
        address = requireNonBlank(address, "address");
        formula = requireNonBlank(formula, "formula");
      }
    }
  }

  /** Stable step identity captured for step execution failures. */
  record StepReference(int stepIndex, StepId stepId, String stepKind, String stepType) {
    public StepReference {
      if (stepIndex < 0) {
        throw new IllegalArgumentException("stepIndex must be greater than or equal to 0");
      }

      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      stepKind = requireNonBlank(stepKind, "stepKind");
      stepType = requireNonBlank(stepType, "stepType");
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    return TextValidation.requireNonBlank(value, fieldName);
  }
}
