package dev.erst.gridgrind.engine.links;

import dev.erst.gridgrind.core.limits.Limit;
import dev.erst.gridgrind.core.model.value.CellHyperlinkReport;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookLocation;
import dev.erst.gridgrind.engine.cells.CellAccess;
import dev.erst.gridgrind.engine.poi.CellBounds;
import java.net.URI;
import java.net.URISyntaxException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.InvalidPathException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.regex.Pattern;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

/** Owns hyperlink mutation/readback and local-file URI normalization/resolution. */
@SuppressWarnings("PMD.CommentRequired")
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_MAXIMUM_HYPERLINKS_PER_WORKSHEET
})
public final class LinkAdapter {
  private final Sheet sheet;

  public LinkAdapter(Sheet sheet) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
  }

  public LinkAdapter setHyperlink(String address, HyperlinkTarget hyperlink) {
    CellAccess.requireAddress(address);
    Objects.requireNonNull(hyperlink, "hyperlink must not be null");

    CellReference cellReference = CellAccess.parseCellReference(address);
    Cell cell = CellAccess.getOrCreateCell(sheet, cellReference.getRow(), cellReference.getCol());
    requireHyperlinkCapacity(cell);
    cell.removeHyperlink();
    org.apache.poi.ss.usermodel.Hyperlink poiHyperlink =
        sheet.getWorkbook().getCreationHelper().createHyperlink(toPoi(hyperlink));
    poiHyperlink.setAddress(toPoiTarget(hyperlink));
    cell.setHyperlink(poiHyperlink);
    return this;
  }

  public LinkAdapter clearHyperlink(String address) {
    CellAccess.requireAddress(address);
    CellAccess.optionalCell(sheet, address).ifPresent(Cell::removeHyperlink);
    return this;
  }

  public static Optional<HyperlinkTarget> hyperlink(Cell cell) {
    return cell == null ? Optional.empty() : hyperlink(cell.getHyperlink());
  }

  public static Optional<HyperlinkTarget> hyperlink(
      org.apache.poi.ss.usermodel.Hyperlink hyperlink) {
    return hyperlink == null || hyperlink.getType() == null
        ? Optional.empty()
        : hyperlink(hyperlink.getType(), hyperlink.getAddress());
  }

  public static Optional<HyperlinkTarget> hyperlink(HyperlinkType hyperlinkType, String target) {
    if (hyperlinkType == null || target == null || target.isBlank()) {
      return Optional.empty();
    }
    return switch (hyperlinkType) {
      case URL -> urlHyperlink(target);
      case EMAIL -> emailHyperlink(target);
      case FILE -> fileHyperlink(target);
      case DOCUMENT -> Optional.of(new HyperlinkTarget.Document(target));
      case NONE -> Optional.empty();
    };
  }

  public static HyperlinkType toPoi(HyperlinkTarget hyperlinkType) {
    return switch (hyperlinkType) {
      case HyperlinkTarget.Url _ -> HyperlinkType.URL;
      case HyperlinkTarget.Email _ -> HyperlinkType.EMAIL;
      case HyperlinkTarget.File _ -> HyperlinkType.FILE;
      case HyperlinkTarget.Document _ -> HyperlinkType.DOCUMENT;
    };
  }

  public static String toPoiTarget(HyperlinkTarget hyperlink) {
    return switch (hyperlink) {
      case HyperlinkTarget.Url url -> url.target();
      case HyperlinkTarget.Email email -> "mailto:" + email.email();
      case HyperlinkTarget.File file -> toPoiAddress(file.path());
      case HyperlinkTarget.Document document -> document.target();
    };
  }

  private List<CellHyperlinkReport> allUsedHyperlinks() {
    List<CellHyperlinkReport> hyperlinks = new ArrayList<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        Optional<HyperlinkTarget> hyperlink = hyperlink(cell);
        if (hyperlink.isPresent()) {
          hyperlinks.add(
              new CellHyperlinkReport(
                  new dev.erst.gridgrind.core.vocabulary.CellAddress(
                      new CellReference(cell.getRowIndex(), cell.getColumnIndex())
                          .formatAsString()),
                  hyperlink.orElseThrow()));
        }
      }
    }
    return List.copyOf(hyperlinks);
  }

  public List<CellHyperlinkReport> hyperlinks() {
    return allUsedHyperlinks();
  }

  public List<CellHyperlinkReport> hyperlinks(
      List<dev.erst.gridgrind.core.vocabulary.CellAddress> addresses) {
    return selectedHyperlinks(List.copyOf(addresses));
  }

  private List<CellHyperlinkReport> selectedHyperlinks(
      List<dev.erst.gridgrind.core.vocabulary.CellAddress> addresses) {
    List<CellHyperlinkReport> hyperlinks = new ArrayList<>();
    for (dev.erst.gridgrind.core.vocabulary.CellAddress address : addresses) {
      Cell cell = CellAccess.cellOrNull(sheet, address.value()).orElse(null);
      if (cell == null) {
        continue;
      }
      Optional<HyperlinkTarget> hyperlink = hyperlink(cell);
      if (hyperlink.isPresent()) {
        hyperlinks.add(new CellHyperlinkReport(address, hyperlink.orElseThrow()));
      }
    }
    return List.copyOf(hyperlinks);
  }

  private void requireHyperlinkCapacity(Cell cell) {
    if (hyperlink(cell).isPresent()) {
      return;
    }
    int hyperlinkCount = 0;
    for (Row row : sheet) {
      for (Cell candidate : row) {
        if (hyperlink(candidate).isPresent()) {
          hyperlinkCount++;
        }
      }
    }
    requireWorksheetHyperlinkCapacity(hyperlinkCount);
  }

  private static Optional<HyperlinkTarget> urlHyperlink(String target) {
    return isValidUrlTarget(target)
        ? Optional.of(new HyperlinkTarget.Url(target))
        : Optional.empty();
  }

  private static Optional<HyperlinkTarget> emailHyperlink(String target) {
    return isValidEmailTarget(target)
        ? Optional.of(new HyperlinkTarget.Email(target))
        : Optional.empty();
  }

  private static Optional<HyperlinkTarget> fileHyperlink(String target) {
    return isValidFileTarget(target)
        ? Optional.of(new HyperlinkTarget.File(target))
        : Optional.empty();
  }

  private static final char[] HEX_DIGITS = "0123456789ABCDEF".toCharArray();
  private static final String SAFE_URI_PATH_CODE_POINTS = "-._~/:";

  /**
   * Normalizes one request or workbook file hyperlink value to the canonical stored path string.
   */
  public static String normalizePath(String path) {
    TextValidation.requireNonBlank(path, "path");
    if (looksLikeFileUri(path)) {
      return normalizeFileUri(path);
    }
    if (looksLikeAbsoluteUri(path)) {
      throw new IllegalArgumentException("path must be a local file path or file: URI");
    }
    return decodeEscapedRelativePath(path);
  }

  /**
   * Converts one request-side file hyperlink path into the escaped address string required by POI.
   */
  public static String toPoiAddress(String path) {
    String normalizedPath = normalizePath(path);
    if (looksLikeWindowsDrivePath(normalizedPath)) {
      return encodeAbsoluteUriPath("/" + normalizedPath.replace('\\', '/'));
    }

    Path candidate;
    try {
      candidate = Path.of(normalizedPath);
    } catch (InvalidPathException exception) {
      throw new IllegalArgumentException(invalidPathMessage(exception), exception);
    }

    if (candidate.isAbsolute()) {
      return candidate.toUri().toASCIIString();
    }
    return encodeRelativePath(normalizedPath);
  }

  /**
   * Resolves one normalized file hyperlink path against the workbook's current filesystem anchor.
   */
  public static FileHyperlinkResolution resolve(String path, WorkbookLocation workbookLocation) {
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");

    String normalizedPath;
    try {
      normalizedPath = normalizePath(path);
    } catch (IllegalArgumentException exception) {
      return new FileHyperlinkResolution.MalformedPath(
          path,
          Objects.requireNonNullElse(
              exception.getMessage(), "path must be a valid local file path"));
    }

    Path candidate;
    try {
      candidate = Path.of(normalizedPath);
    } catch (InvalidPathException exception) {
      return new FileHyperlinkResolution.MalformedPath(
          normalizedPath, invalidPathMessage(exception));
    }

    if (candidate.isAbsolute()) {
      return new FileHyperlinkResolution.ResolvedPath(normalizedPath, candidate);
    }
    return switch (workbookLocation) {
      case WorkbookLocation.StoredWorkbook storedWorkbook ->
          new FileHyperlinkResolution.ResolvedPath(
              normalizedPath, storedWorkbook.baseDirectory().orElseThrow().resolve(candidate));
      case WorkbookLocation.UnsavedWorkbook _ ->
          new FileHyperlinkResolution.UnresolvedRelativePath(normalizedPath);
    };
  }

  /** Returns whether the normalized stored path is relative in GridGrind's file-target contract. */
  public static boolean isRelativeStoredPath(String path) {
    Objects.requireNonNull(path, "path must not be null");
    if (looksLikeWindowsDrivePath(path)) {
      return false;
    }
    try {
      return !Path.of(path).isAbsolute();
    } catch (InvalidPathException exception) {
      return false;
    }
  }

  private static boolean looksLikeFileUri(String path) {
    return path.regionMatches(true, 0, "file:", 0, 5);
  }

  private static String normalizeFileUri(String path) {
    try {
      URI uri = new URI(path);
      try {
        return Path.of(uri).toString();
      } catch (RuntimeException exception) {
        String authority = uri.getAuthority();
        String uriPath = uri.getPath();
        if (uriPath == null || uriPath.isBlank()) {
          throw new IllegalArgumentException("path must contain a file-system path", exception);
        }
        return authority == null ? uriPath : "//" + authority + uriPath;
      }
    } catch (URISyntaxException exception) {
      throw new IllegalArgumentException("path must be a valid file: URI", exception);
    }
  }

  private static boolean looksLikeAbsoluteUri(String path) {
    if (looksLikeWindowsDrivePath(path)) {
      return false;
    }
    try {
      return new URI(path).isAbsolute();
    } catch (URISyntaxException exception) {
      return false;
    }
  }

  private static boolean looksLikeWindowsDrivePath(String path) {
    return path.length() >= 3
        && Character.isLetter(path.charAt(0))
        && path.charAt(1) == ':'
        && (path.charAt(2) == '/' || path.charAt(2) == '\\');
  }

  private static String decodeEscapedRelativePath(String path) {
    if (!path.contains("%")) {
      return path;
    }
    try {
      URI uri = new URI(path);
      if (uri.isAbsolute()) {
        return path;
      }
      String decodedPath = Objects.requireNonNullElse(uri.getPath(), "");
      return decodedPath.isBlank() ? path : decodedPath;
    } catch (URISyntaxException exception) {
      return path;
    }
  }

  private static String encodeAbsoluteUriPath(String path) {
    return "file://" + encodePath(path);
  }

  private static String encodeRelativePath(String path) {
    return encodePath(path.replace('\\', '/'));
  }

  private static String encodePath(String path) {
    StringBuilder builder = new StringBuilder(path.length());
    path.codePoints()
        .forEach(
            codePoint -> {
              if (isPlainUriPathCodePoint(codePoint)) {
                builder.appendCodePoint(codePoint);
                return;
              }
              appendPercentEncoded(builder, codePoint);
            });
    return builder.toString();
  }

  private static boolean isPlainUriPathCodePoint(int codePoint) {
    return Character.isLetterOrDigit(codePoint)
        || SAFE_URI_PATH_CODE_POINTS.indexOf(codePoint) >= 0;
  }

  private static void appendPercentEncoded(StringBuilder builder, int codePoint) {
    for (byte value : new String(Character.toChars(codePoint)).getBytes(StandardCharsets.UTF_8)) {
      int unsigned = value & 0xFF;
      builder.append('%').append(HEX_DIGITS[unsigned >>> 4]).append(HEX_DIGITS[unsigned & 0x0F]);
    }
  }

  private static String invalidPathMessage(InvalidPathException exception) {
    return Objects.toString(exception.getReason(), "path is not valid on this runtime");
  }

  private static final Pattern ABSOLUTE_URI_PATTERN =
      Pattern.compile("^[A-Za-z][A-Za-z0-9+.-]*:\\S+$");

  public static boolean isValidUrlTarget(String target) {
    if (target == null || target.isBlank() || !ABSOLUTE_URI_PATTERN.matcher(target).matches()) {
      return false;
    }
    try {
      URI uri = URI.create(target);
      String scheme = uri.getScheme();
      return !"file".equalsIgnoreCase(scheme) && !"mailto".equalsIgnoreCase(scheme);
    } catch (IllegalArgumentException exception) {
      return false;
    }
  }

  public static boolean isValidEmailTarget(String target) {
    if (target == null) {
      return false;
    }
    String normalized = stripMailtoPrefix(target);
    if (normalized.isBlank() || normalized.contains(" ")) {
      return false;
    }
    int atIndex = normalized.indexOf('@');
    return atIndex > 0
        && atIndex == normalized.lastIndexOf('@')
        && atIndex < normalized.length() - 1;
  }

  public static boolean isValidFileTarget(String target) {
    try {
      normalizePath(target);
      return true;
    } catch (IllegalArgumentException | NullPointerException exception) {
      return false;
    }
  }

  public static boolean isValidDocumentTarget(String target) {
    return target != null && !target.isBlank();
  }

  public static Optional<String> absoluteUriScheme(String target) {
    if (target == null || target.isBlank()) {
      return Optional.empty();
    }
    try {
      URI uri = URI.create(target);
      return uri.isAbsolute() ? Optional.ofNullable(uri.getScheme()) : Optional.empty();
    } catch (IllegalArgumentException exception) {
      return Optional.empty();
    }
  }

  public static String stripMailtoPrefix(String email) {
    if (email.regionMatches(true, 0, "mailto:", 0, 7)) {
      return email.substring(7);
    }
    return email;
  }

  public static void requireWorksheetHyperlinkCapacity(int existingHyperlinkCount) {
    if (existingHyperlinkCount >= Limit.MAX_HYPERLINKS) {
      throw new IllegalArgumentException(
          "sheet cannot contain more than "
              + Limit.MAX_HYPERLINKS
              + " hyperlinks (Excel worksheet hyperlink limit)");
    }
  }

  public int hyperlinkCount() {
    int count = 0;
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (hasUsableHyperlink(cell.getHyperlink())) {
          count++;
        }
      }
    }
    return count;
  }

  public List<AnalysisFindingReport> hyperlinkHealthFindings(WorkbookLocation workbookLocation) {
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");
    List<AnalysisFindingReport> findings = new ArrayList<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        org.apache.poi.ss.usermodel.Hyperlink hyperlink = cell.getHyperlink();
        if (!hasUsableHyperlink(hyperlink)) {
          continue;
        }
        HyperlinkType hyperlinkType = hyperlink.getType();
        String address = cell.getAddress().formatAsString();
        AnalysisLocationReport.Cell location = cellLocation(sheet.getSheetName(), address);
        findings.addAll(
            hyperlinkTargetFindings(
                location, hyperlinkType, hyperlink.getAddress(), workbookLocation));
      }
    }
    return List.copyOf(findings);
  }

  public List<AnalysisFindingReport> hyperlinkHealthFindings() {
    return hyperlinkHealthFindings(new WorkbookLocation.UnsavedWorkbook());
  }

  public static List<AnalysisFindingReport> externalHyperlinkFindings(
      AnalysisLocationReport.Cell location, String target, String expectedShape, boolean valid) {
    Objects.requireNonNull(location, "location must not be null");
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(expectedShape, "expectedShape must not be null");
    return valid
        ? List.of()
        : List.of(
            malformedHyperlinkFinding(
                location,
                "target does not match the expected " + expectedShape + " shape",
                List.of(target)));
  }

  public List<AnalysisFindingReport> hyperlinkTargetFindings(
      AnalysisLocationReport.Cell location,
      HyperlinkType hyperlinkType,
      String target,
      WorkbookLocation workbookLocation) {
    Objects.requireNonNull(location, "location must not be null");
    Objects.requireNonNull(hyperlinkType, "hyperlinkType must not be null");
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");

    if (hasMissingHyperlinkTarget(target)) {
      return List.of(
          malformedHyperlinkFinding(
              location, "Hyperlink target is blank or missing.", List.of(hyperlinkType.name())));
    }
    return switch (hyperlinkType) {
      case URL ->
          externalHyperlinkFindings(
              location, target, HyperlinkType.URL.name(), LinkAdapter.isValidUrlTarget(target));
      case EMAIL ->
          externalHyperlinkFindings(
              location, target, HyperlinkType.EMAIL.name(), LinkAdapter.isValidEmailTarget(target));
      case FILE -> fileHyperlinkFindings(location, target, workbookLocation);
      case DOCUMENT -> {
        List<AnalysisFindingReport> findings = new ArrayList<>();
        validateDocumentHyperlinkTarget(location, target, findings);
        yield List.copyOf(findings);
      }
      case NONE -> List.of();
    };
  }

  public void validateDocumentHyperlinkTarget(
      AnalysisLocationReport.Cell location, String target, List<AnalysisFindingReport> findings) {
    int bangIndex = target.indexOf('!');
    if (bangIndex <= 0 || bangIndex >= target.length() - 1) {
      findings.add(
          new AnalysisFindingReport(
              AnalysisFindingCode.HYPERLINK_INVALID_DOCUMENT_TARGET,
              AnalysisSeverity.ERROR,
              "Invalid document hyperlink target",
              "Document hyperlink target must include a sheet and cell or range reference.",
              location,
              List.of(target)));
      return;
    }

    String targetSheetName = unquoteSheetName(target.substring(0, bangIndex));
    String targetAddress = target.substring(bangIndex + 1);
    if (sheet.getWorkbook().getSheet(targetSheetName) == null) {
      findings.add(
          new AnalysisFindingReport(
              AnalysisFindingCode.HYPERLINK_MISSING_DOCUMENT_SHEET,
              AnalysisSeverity.ERROR,
              "Document hyperlink targets a missing sheet",
              "Document hyperlink target sheet does not exist: " + targetSheetName,
              location,
              List.of(target)));
      return;
    }

    try {
      if (targetAddress.contains(":")) {
        CellBounds.parse(targetAddress);
      } else {
        new CellReference(targetAddress);
      }
    } catch (IllegalArgumentException exception) {
      findings.add(
          new AnalysisFindingReport(
              AnalysisFindingCode.HYPERLINK_INVALID_DOCUMENT_TARGET,
              AnalysisSeverity.ERROR,
              "Invalid document hyperlink target",
              "Document hyperlink target is not a valid cell or range reference.",
              location,
              List.of(target)));
    }
  }

  public static List<AnalysisFindingReport> fileHyperlinkFindings(
      AnalysisLocationReport.Cell location, String target, WorkbookLocation workbookLocation) {
    Objects.requireNonNull(location, "location must not be null");
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");

    return switch (LinkAdapter.resolve(target, workbookLocation)) {
      case FileHyperlinkResolution.MalformedPath malformedPath ->
          List.of(
              malformedHyperlinkFinding(
                  location, malformedPath.reason(), List.of(malformedPath.path())));
      case FileHyperlinkResolution.UnresolvedRelativePath unresolvedRelativePath ->
          List.of(
              new AnalysisFindingReport(
                  AnalysisFindingCode.HYPERLINK_UNRESOLVED_FILE_TARGET,
                  AnalysisSeverity.WARNING,
                  "Relative file hyperlink cannot be resolved yet",
                  "Relative file hyperlink targets cannot be validated until the workbook has a"
                      + " filesystem location.",
                  location,
                  List.of(unresolvedRelativePath.path())));
      case FileHyperlinkResolution.ResolvedPath resolvedPath ->
          Files.notExists(resolvedPath.resolvedPath())
              ? List.of(
                  new AnalysisFindingReport(
                      AnalysisFindingCode.HYPERLINK_MISSING_FILE_TARGET,
                      AnalysisSeverity.ERROR,
                      "File hyperlink targets a missing path",
                      missingFileTargetMessage(resolvedPath, workbookLocation),
                      location,
                      missingFileTargetEvidence(resolvedPath, workbookLocation)))
              : List.of();
    };
  }

  private static String missingFileTargetMessage(
      FileHyperlinkResolution.ResolvedPath resolvedPath, WorkbookLocation workbookLocation) {
    if (workbookLocation instanceof WorkbookLocation.StoredWorkbook storedWorkbook
        && LinkAdapter.isRelativeStoredPath(resolvedPath.path())) {
      return "Resolved relative file hyperlink path does not exist against workbook directory "
          + storedWorkbook.baseDirectory().orElseThrow()
          + ": "
          + resolvedPath.resolvedPath();
    }
    return "Resolved file hyperlink path does not exist: " + resolvedPath.resolvedPath();
  }

  private static List<String> missingFileTargetEvidence(
      FileHyperlinkResolution.ResolvedPath resolvedPath, WorkbookLocation workbookLocation) {
    if (workbookLocation instanceof WorkbookLocation.StoredWorkbook storedWorkbook
        && LinkAdapter.isRelativeStoredPath(resolvedPath.path())) {
      return List.of(
          resolvedPath.path(),
          storedWorkbook.baseDirectory().orElseThrow().toString(),
          resolvedPath.resolvedPath().toString());
    }
    return List.of(resolvedPath.path(), resolvedPath.resolvedPath().toString());
  }

  public static boolean hasUsableHyperlink(org.apache.poi.ss.usermodel.Hyperlink hyperlink) {
    return hyperlink != null && hasUsableHyperlinkType(hyperlink.getType());
  }

  public static boolean hasUsableHyperlinkType(HyperlinkType hyperlinkType) {
    return hyperlinkType != null && hyperlinkType != HyperlinkType.NONE;
  }

  public static boolean hasMissingHyperlinkTarget(String target) {
    return target == null || target.isBlank();
  }

  public static String unquoteSheetName(String sheetName) {
    if (!sheetName.startsWith("'")) {
      return sheetName;
    }
    if (!sheetName.endsWith("'")) {
      return sheetName;
    }
    if (sheetName.length() < 2) {
      return sheetName;
    }
    return sheetName.substring(1, sheetName.length() - 1).replace("''", "'");
  }

  private static AnalysisFindingReport malformedHyperlinkFinding(
      AnalysisLocationReport.Cell location, String message, List<String> evidence) {
    return new AnalysisFindingReport(
        AnalysisFindingCode.HYPERLINK_MALFORMED_TARGET,
        AnalysisSeverity.ERROR,
        "Malformed hyperlink target",
        "Hyperlink target is malformed: " + message,
        location,
        evidence);
  }

  public static AnalysisLocationReport.Cell cellLocation(String sheetName, String address) {
    return new AnalysisLocationReport.Cell(
        new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetName), address);
  }
}
