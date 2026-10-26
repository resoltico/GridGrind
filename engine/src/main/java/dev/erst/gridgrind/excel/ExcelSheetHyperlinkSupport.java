package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.value.CellHyperlinkReport;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

/** Hyperlink mutation and snapshot helpers for one sheet. */
final class ExcelSheetHyperlinkSupport {
  private final Sheet sheet;

  ExcelSheetHyperlinkSupport(Sheet sheet) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
  }

  void setHyperlink(String address, HyperlinkTarget hyperlink) {
    ExcelSheet.requireNonBlank(address, "address");
    Objects.requireNonNull(hyperlink, "hyperlink must not be null");

    CellReference cellReference = ExcelSheetAddressSupport.parseCellReference(address);
    Cell cell =
        ExcelSheetAddressSupport.getOrCreateCell(
            sheet, cellReference.getRow(), cellReference.getCol());
    requireHyperlinkCapacity(cell);
    cell.removeHyperlink();
    org.apache.poi.ss.usermodel.Hyperlink poiHyperlink =
        sheet.getWorkbook().getCreationHelper().createHyperlink(toPoi(hyperlink));
    poiHyperlink.setAddress(toPoiTarget(hyperlink));
    cell.setHyperlink(poiHyperlink);
  }

  void clearHyperlink(String address) {
    ExcelSheet.requireNonBlank(address, "address");
    ExcelSheetAddressSupport.optionalCell(sheet, address).ifPresent(Cell::removeHyperlink);
  }

  List<CellHyperlinkReport> hyperlinks(ExcelCellSelection selection) {
    Objects.requireNonNull(selection, "selection must not be null");
    return switch (selection) {
      case ExcelCellSelection.AllUsedCells _ -> allUsedHyperlinks();
      case ExcelCellSelection.Selected selected -> selectedHyperlinks(selected.addresses());
    };
  }

  static Optional<HyperlinkTarget> hyperlink(Cell cell) {
    return cell == null ? Optional.empty() : hyperlink(cell.getHyperlink());
  }

  static Optional<HyperlinkTarget> hyperlink(org.apache.poi.ss.usermodel.Hyperlink hyperlink) {
    return hyperlink == null || hyperlink.getType() == null
        ? Optional.empty()
        : hyperlink(hyperlink.getType(), hyperlink.getAddress());
  }

  static Optional<HyperlinkTarget> hyperlink(HyperlinkType hyperlinkType, String target) {
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

  static HyperlinkType toPoi(HyperlinkTarget hyperlinkType) {
    return switch (hyperlinkType) {
      case HyperlinkTarget.Url _ -> HyperlinkType.URL;
      case HyperlinkTarget.Email _ -> HyperlinkType.EMAIL;
      case HyperlinkTarget.File _ -> HyperlinkType.FILE;
      case HyperlinkTarget.Document _ -> HyperlinkType.DOCUMENT;
    };
  }

  static String toPoiTarget(HyperlinkTarget hyperlink) {
    return switch (hyperlink) {
      case HyperlinkTarget.Url url -> url.target();
      case HyperlinkTarget.Email email -> "mailto:" + email.email();
      case HyperlinkTarget.File file -> ExcelFileHyperlinkTargets.toPoiAddress(file.path());
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

  private List<CellHyperlinkReport> selectedHyperlinks(List<String> addresses) {
    List<CellHyperlinkReport> hyperlinks = new ArrayList<>();
    for (String address : addresses) {
      Cell cell = ExcelSheetAddressSupport.cellOrNull(sheet, address).orElse(null);
      if (cell == null) {
        continue;
      }
      Optional<HyperlinkTarget> hyperlink = hyperlink(cell);
      if (hyperlink.isPresent()) {
        hyperlinks.add(
            new CellHyperlinkReport(
                new dev.erst.gridgrind.core.vocabulary.CellAddress(address),
                hyperlink.orElseThrow()));
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
    ExcelHyperlinkLimits.requireWorksheetHyperlinkCapacity(hyperlinkCount); // LIM-012
  }

  private static Optional<HyperlinkTarget> urlHyperlink(String target) {
    return ExcelHyperlinkValidation.isValidUrlTarget(target)
        ? Optional.of(new HyperlinkTarget.Url(target))
        : Optional.empty();
  }

  private static Optional<HyperlinkTarget> emailHyperlink(String target) {
    return ExcelHyperlinkValidation.isValidEmailTarget(target)
        ? Optional.of(new HyperlinkTarget.Email(target))
        : Optional.empty();
  }

  private static Optional<HyperlinkTarget> fileHyperlink(String target) {
    return ExcelHyperlinkValidation.isValidFileTarget(target)
        ? Optional.of(new HyperlinkTarget.File(target))
        : Optional.empty();
  }
}
