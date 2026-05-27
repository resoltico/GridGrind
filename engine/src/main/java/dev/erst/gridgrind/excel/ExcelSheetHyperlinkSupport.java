package dev.erst.gridgrind.excel;

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

  void setHyperlink(String address, ExcelHyperlink hyperlink) {
    ExcelSheet.requireNonBlank(address, "address");
    Objects.requireNonNull(hyperlink, "hyperlink must not be null");

    CellReference cellReference = ExcelSheetAddressSupport.parseCellReference(address);
    Cell cell =
        ExcelSheetAddressSupport.getOrCreateCell(
            sheet, cellReference.getRow(), cellReference.getCol());
    requireHyperlinkCapacity(cell);
    cell.removeHyperlink();
    org.apache.poi.ss.usermodel.Hyperlink poiHyperlink =
        sheet.getWorkbook().getCreationHelper().createHyperlink(toPoi(hyperlink.type()));
    poiHyperlink.setAddress(toPoiTarget(hyperlink));
    cell.setHyperlink(poiHyperlink);
  }

  void clearHyperlink(String address) {
    ExcelSheet.requireNonBlank(address, "address");
    ExcelSheetAddressSupport.optionalCell(sheet, address).ifPresent(Cell::removeHyperlink);
  }

  List<WorkbookSheetResult.CellHyperlink> hyperlinks(ExcelCellSelection selection) {
    Objects.requireNonNull(selection, "selection must not be null");
    return switch (selection) {
      case ExcelCellSelection.AllUsedCells _ -> allUsedHyperlinks();
      case ExcelCellSelection.Selected selected -> selectedHyperlinks(selected.addresses());
    };
  }

  static Optional<ExcelHyperlink> hyperlink(Cell cell) {
    return cell == null ? Optional.empty() : hyperlink(cell.getHyperlink());
  }

  static Optional<ExcelHyperlink> hyperlink(org.apache.poi.ss.usermodel.Hyperlink hyperlink) {
    return hyperlink == null || hyperlink.getType() == null
        ? Optional.empty()
        : hyperlink(hyperlink.getType(), hyperlink.getAddress());
  }

  static Optional<ExcelHyperlink> hyperlink(HyperlinkType hyperlinkType, String target) {
    if (hyperlinkType == null || target == null || target.isBlank()) {
      return Optional.empty();
    }
    return switch (hyperlinkType) {
      case URL -> urlHyperlink(target);
      case EMAIL -> emailHyperlink(target);
      case FILE -> fileHyperlink(target);
      case DOCUMENT -> Optional.of(new ExcelHyperlink.Document(target));
      case NONE -> Optional.empty();
    };
  }

  static HyperlinkType toPoi(ExcelHyperlinkType hyperlinkType) {
    return switch (hyperlinkType) {
      case URL -> HyperlinkType.URL;
      case EMAIL -> HyperlinkType.EMAIL;
      case FILE -> HyperlinkType.FILE;
      case DOCUMENT -> HyperlinkType.DOCUMENT;
    };
  }

  static String toPoiTarget(ExcelHyperlink hyperlink) {
    return switch (hyperlink) {
      case ExcelHyperlink.Url url -> url.target();
      case ExcelHyperlink.Email email -> "mailto:" + email.target();
      case ExcelHyperlink.File file -> ExcelFileHyperlinkTargets.toPoiAddress(file.path());
      case ExcelHyperlink.Document document -> document.target();
    };
  }

  private List<WorkbookSheetResult.CellHyperlink> allUsedHyperlinks() {
    List<WorkbookSheetResult.CellHyperlink> hyperlinks = new ArrayList<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        Optional<ExcelHyperlink> hyperlink = hyperlink(cell);
        if (hyperlink.isPresent()) {
          hyperlinks.add(
              new WorkbookSheetResult.CellHyperlink(
                  new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(),
                  hyperlink.orElseThrow()));
        }
      }
    }
    return List.copyOf(hyperlinks);
  }

  private List<WorkbookSheetResult.CellHyperlink> selectedHyperlinks(List<String> addresses) {
    List<WorkbookSheetResult.CellHyperlink> hyperlinks = new ArrayList<>();
    for (String address : addresses) {
      Cell cell = ExcelSheetAddressSupport.cellOrNull(sheet, address).orElse(null);
      if (cell == null) {
        continue;
      }
      Optional<ExcelHyperlink> hyperlink = hyperlink(cell);
      if (hyperlink.isPresent()) {
        hyperlinks.add(new WorkbookSheetResult.CellHyperlink(address, hyperlink.orElseThrow()));
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

  private static Optional<ExcelHyperlink> urlHyperlink(String target) {
    return ExcelHyperlinkValidation.isValidUrlTarget(target)
        ? Optional.of(new ExcelHyperlink.Url(target))
        : Optional.empty();
  }

  private static Optional<ExcelHyperlink> emailHyperlink(String target) {
    return ExcelHyperlinkValidation.isValidEmailTarget(target)
        ? Optional.of(new ExcelHyperlink.Email(target))
        : Optional.empty();
  }

  private static Optional<ExcelHyperlink> fileHyperlink(String target) {
    return ExcelHyperlinkValidation.isValidFileTarget(target)
        ? Optional.of(new ExcelHyperlink.File(target))
        : Optional.empty();
  }
}
