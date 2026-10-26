package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.SourceBackedResolutionIdentitySupport.sameOptionalReference;
import static dev.erst.gridgrind.engine.runtime.SourceBackedResolutionIdentitySupport.sameReference;

import dev.erst.gridgrind.core.model.chart.ChartSpec;
import dev.erst.gridgrind.core.model.customxml.CustomXmlImportSpec;
import dev.erst.gridgrind.core.model.drawing.CommentSpec;
import dev.erst.gridgrind.core.model.drawing.EmbeddedObjectSpec;
import dev.erst.gridgrind.core.model.drawing.PictureDataSpec;
import dev.erst.gridgrind.core.model.drawing.PictureSpec;
import dev.erst.gridgrind.core.model.drawing.ShapeSpec;
import dev.erst.gridgrind.core.model.drawing.SignatureLineSpec;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextSpec;
import dev.erst.gridgrind.core.model.print.PrintLayoutSpec;
import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationSpec;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/** Resolves authored structured inputs whose nested values can be bound from external sources. */
final class SourceBackedStructuredInputResolver {
  private SourceBackedStructuredInputResolver() {}

  static List<RichTextRunSpec> resolveRuns(
      List<RichTextRunSpec> runs, ExecutionInputBindings bindings) throws IOException {
    List<RichTextRunSpec> resolvedRuns = new ArrayList<>(runs.size());
    boolean changed = false;
    for (RichTextRunSpec run : runs) {
      RichTextRunSpec resolvedRun = resolveRichTextRun(run, bindings);
      resolvedRuns.add(resolvedRun);
      changed |= !sameReference(resolvedRun, run);
    }
    return changed ? List.copyOf(resolvedRuns) : runs;
  }

  private static RichTextRunSpec resolveRichTextRun(
      RichTextRunSpec run, ExecutionInputBindings bindings) throws IOException {
    TextSourceSpec resolvedSource =
        SourceBackedPlanResolver.resolveTextSource(run.text(), bindings, false, "rich-text run");
    if (!(resolvedSource instanceof TextSourceSpec.Inline inline)) {
      return run;
    }
    String resolvedText = inline.text();
    if (resolvedText.isEmpty()) {
      IllegalArgumentException failure =
          new IllegalArgumentException("rich-text run must not be empty");
      if (bindings.collectInputResolutionFailure(failure, run.text())) {
        return run;
      }
      throw failure;
    }
    return sameReference(resolvedSource, run.text())
        ? run
        : new RichTextRunSpec(resolvedSource, run.font());
  }

  static CommentSpec resolveComment(CommentSpec comment, ExecutionInputBindings bindings)
      throws IOException {
    TextSourceSpec resolvedText =
        SourceBackedPlanResolver.resolveTextSource(comment.text(), bindings, true, "comment text");
    Optional<List<RichTextRunSpec>> resolvedRuns =
        comment.runs().isEmpty()
            ? Optional.empty()
            : Optional.of(resolveRuns(comment.runs().orElseThrow(), bindings));
    return sameReference(resolvedText, comment.text())
            && sameOptionalReference(resolvedRuns, comment.runs())
        ? comment
        : new CommentSpec(
            resolvedText, comment.author(), comment.visible(), resolvedRuns, comment.anchor());
  }

  static CustomXmlImportSpec resolveCustomXmlImport(
      CustomXmlImportSpec input, ExecutionInputBindings bindings) throws IOException {
    TextSourceSpec resolvedXml =
        SourceBackedPlanResolver.resolveTextSource(input.xml(), bindings, true, "custom XML");
    return sameReference(resolvedXml, input.xml())
        ? input
        : new CustomXmlImportSpec(input.locator(), resolvedXml);
  }

  static SignatureLineSpec resolveSignatureLine(
      SignatureLineSpec signatureLine, ExecutionInputBindings bindings) throws IOException {
    if (signatureLine.plainSignature().isEmpty()) {
      return signatureLine;
    }
    PictureDataSpec resolvedPlainSignature =
        resolvePictureData(signatureLine.plainSignature().orElseThrow(), bindings);
    return sameOptionalReference(
            Optional.of(resolvedPlainSignature), signatureLine.plainSignature())
        ? signatureLine
        : new SignatureLineSpec(
            signatureLine.name(),
            signatureLine.anchor(),
            signatureLine.allowComments(),
            signatureLine.signingInstructions(),
            signatureLine.suggestedSigner(),
            signatureLine.suggestedSigner2(),
            signatureLine.suggestedSignerEmail(),
            signatureLine.caption(),
            signatureLine.invalidStamp(),
            Optional.of(resolvedPlainSignature));
  }

  static PictureSpec resolvePicture(PictureSpec picture, ExecutionInputBindings bindings)
      throws IOException {
    PictureDataSpec resolvedImage = resolvePictureData(picture.image(), bindings);
    Optional<TextSourceSpec> resolvedDescription = picture.description();
    if (picture.description().isPresent()) {
      resolvedDescription =
          Optional.of(
              SourceBackedPlanResolver.resolveTextSource(
                  picture.description().orElseThrow(), bindings, true, "picture description"));
    }
    return sameReference(resolvedImage, picture.image())
            && sameOptionalReference(resolvedDescription, picture.description())
        ? picture
        : new PictureSpec(picture.name(), resolvedImage, picture.anchor(), resolvedDescription);
  }

  private static PictureDataSpec resolvePictureData(
      PictureDataSpec image, ExecutionInputBindings bindings) throws IOException {
    BinarySourceSpec resolvedSource =
        SourceBackedPlanResolver.resolveBinarySource(image.source(), bindings, "picture payload");
    return sameReference(resolvedSource, image.source())
        ? image
        : new PictureDataSpec(image.format(), resolvedSource);
  }

  static EmbeddedObjectSpec resolveEmbeddedObject(
      EmbeddedObjectSpec embeddedObject, ExecutionInputBindings bindings) throws IOException {
    BinarySourceSpec resolvedPayload =
        SourceBackedPlanResolver.resolveBinarySource(
            embeddedObject.payload(), bindings, "embedded-object payload");
    PictureDataSpec resolvedPreview = resolvePictureData(embeddedObject.previewImage(), bindings);
    return sameReference(resolvedPayload, embeddedObject.payload())
            && sameReference(resolvedPreview, embeddedObject.previewImage())
        ? embeddedObject
        : new EmbeddedObjectSpec(
            embeddedObject.name(),
            embeddedObject.label(),
            embeddedObject.fileName(),
            embeddedObject.command(),
            resolvedPayload,
            resolvedPreview,
            embeddedObject.anchor());
  }

  static ChartSpec resolveChart(ChartSpec chart, ExecutionInputBindings bindings)
      throws IOException {
    return SourceBackedChartSpecResolver.resolveChart(chart, bindings);
  }

  static ShapeSpec resolveShape(ShapeSpec shape, ExecutionInputBindings bindings)
      throws IOException {
    return switch (shape) {
      case ShapeSpec.SimpleShape simpleShape -> {
        Optional<TextSourceSpec> resolvedText = simpleShape.text();
        if (simpleShape.text().isPresent()) {
          resolvedText =
              Optional.of(
                  SourceBackedPlanResolver.resolveTextSource(
                      simpleShape.text().orElseThrow(), bindings, true, "shape text"));
        }
        yield sameOptionalReference(resolvedText, simpleShape.text())
            ? shape
            : new ShapeSpec.SimpleShape(
                simpleShape.name(),
                simpleShape.anchor(),
                simpleShape.presetGeometryToken(),
                resolvedText);
      }
      case ShapeSpec.Connector _ -> shape;
    };
  }

  static DataValidationSpec resolveDataValidation(
      DataValidationSpec validation, ExecutionInputBindings bindings) throws IOException {
    Optional<DataValidationPromptSpec> resolvedPrompt = validation.prompt();
    if (validation.prompt().isPresent()) {
      resolvedPrompt = Optional.of(resolvePrompt(validation.prompt().orElseThrow(), bindings));
    }
    Optional<DataValidationErrorAlertSpec> resolvedAlert = validation.errorAlert();
    if (validation.errorAlert().isPresent()) {
      resolvedAlert =
          Optional.of(resolveErrorAlert(validation.errorAlert().orElseThrow(), bindings));
    }
    return sameOptionalReference(resolvedPrompt, validation.prompt())
            && sameOptionalReference(resolvedAlert, validation.errorAlert())
        ? validation
        : new DataValidationSpec(
            validation.rule(),
            validation.allowBlank(),
            validation.suppressDropDownArrow(),
            resolvedPrompt,
            resolvedAlert);
  }

  private static DataValidationPromptSpec resolvePrompt(
      DataValidationPromptSpec prompt, ExecutionInputBindings bindings) throws IOException {
    TextSourceSpec resolvedTitle =
        SourceBackedPlanResolver.resolveTextSource(
            prompt.title(), bindings, true, "validation prompt title");
    TextSourceSpec resolvedText =
        SourceBackedPlanResolver.resolveTextSource(
            prompt.text(), bindings, true, "validation prompt text");
    return sameReference(resolvedTitle, prompt.title())
            && sameReference(resolvedText, prompt.text())
        ? prompt
        : new DataValidationPromptSpec(resolvedTitle, resolvedText, prompt.showPromptBox());
  }

  private static DataValidationErrorAlertSpec resolveErrorAlert(
      DataValidationErrorAlertSpec alert, ExecutionInputBindings bindings) throws IOException {
    TextSourceSpec resolvedTitle =
        SourceBackedPlanResolver.resolveTextSource(
            alert.title(), bindings, true, "validation error title");
    TextSourceSpec resolvedText =
        SourceBackedPlanResolver.resolveTextSource(
            alert.text(), bindings, true, "validation error text");
    return sameReference(resolvedTitle, alert.title()) && sameReference(resolvedText, alert.text())
        ? alert
        : new DataValidationErrorAlertSpec(
            alert.style(), resolvedTitle, resolvedText, alert.showErrorBox());
  }

  static TableSpec resolveTable(TableSpec table, ExecutionInputBindings bindings)
      throws IOException {
    TextSourceSpec resolvedComment =
        SourceBackedPlanResolver.resolveTextSource(
            table.comment(), bindings, false, "table comment");
    return sameReference(resolvedComment, table.comment())
        ? table
        : new TableSpec(
            table.name(),
            table.sheetName(),
            table.range(),
            table.showTotalsRow(),
            table.hasAutofilter(),
            table.style(),
            resolvedComment,
            table.published(),
            table.insertRow(),
            table.insertRowShift(),
            table.headerRowCellStyle(),
            table.dataCellStyle(),
            table.totalsRowCellStyle(),
            table.columns());
  }

  static PrintLayoutSpec resolvePrintLayout(
      PrintLayoutSpec printLayout, ExecutionInputBindings bindings) throws IOException {
    HeaderFooterTextSpec resolvedHeader = resolveHeaderFooter(printLayout.header(), bindings);
    HeaderFooterTextSpec resolvedFooter = resolveHeaderFooter(printLayout.footer(), bindings);
    return sameReference(resolvedHeader, printLayout.header())
            && sameReference(resolvedFooter, printLayout.footer())
        ? printLayout
        : new PrintLayoutSpec(
            printLayout.printArea(),
            printLayout.orientation(),
            printLayout.scaling(),
            printLayout.repeatingRows(),
            printLayout.repeatingColumns(),
            resolvedHeader,
            resolvedFooter,
            printLayout.setup());
  }

  static HeaderFooterTextSpec resolveHeaderFooter(
      HeaderFooterTextSpec text, ExecutionInputBindings bindings) throws IOException {
    TextSourceSpec resolvedLeft =
        SourceBackedPlanResolver.resolveTextSource(
            text.left(), bindings, false, "header/footer left");
    TextSourceSpec resolvedCenter =
        SourceBackedPlanResolver.resolveTextSource(
            text.center(), bindings, false, "header/footer center");
    TextSourceSpec resolvedRight =
        SourceBackedPlanResolver.resolveTextSource(
            text.right(), bindings, false, "header/footer right");
    return sameReference(resolvedLeft, text.left())
            && sameReference(resolvedCenter, text.center())
            && sameReference(resolvedRight, text.right())
        ? text
        : new HeaderFooterTextSpec(resolvedLeft, resolvedCenter, resolvedRight);
  }
}
