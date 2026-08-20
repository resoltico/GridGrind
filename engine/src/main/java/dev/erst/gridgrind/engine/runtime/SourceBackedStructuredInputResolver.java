package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.SourceBackedResolutionIdentitySupport.sameOptionalReference;
import static dev.erst.gridgrind.engine.runtime.SourceBackedResolutionIdentitySupport.sameReference;

import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.HeaderFooterTextInput;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.SignatureLineInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/** Resolves authored structured inputs whose nested values can be bound from external sources. */
final class SourceBackedStructuredInputResolver {
  private SourceBackedStructuredInputResolver() {}

  static List<RichTextRunInput> resolveRuns(
      List<RichTextRunInput> runs, ExecutionInputBindings bindings) throws IOException {
    List<RichTextRunInput> resolvedRuns = new ArrayList<>(runs.size());
    boolean changed = false;
    for (RichTextRunInput run : runs) {
      RichTextRunInput resolvedRun = resolveRichTextRun(run, bindings);
      resolvedRuns.add(resolvedRun);
      changed |= !sameReference(resolvedRun, run);
    }
    return changed ? List.copyOf(resolvedRuns) : runs;
  }

  private static RichTextRunInput resolveRichTextRun(
      RichTextRunInput run, ExecutionInputBindings bindings) throws IOException {
    TextSourceInput resolvedSource =
        SourceBackedPlanResolver.resolveTextSource(run.source(), bindings, false, "rich-text run");
    if (!(resolvedSource instanceof TextSourceInput.Inline inline)) {
      return run;
    }
    String resolvedText = inline.text();
    if (resolvedText.isEmpty()) {
      IllegalArgumentException failure =
          new IllegalArgumentException("rich-text run must not be empty");
      if (bindings.collectInputResolutionFailure(failure, run.source())) {
        return run;
      }
      throw failure;
    }
    return sameReference(resolvedSource, run.source())
        ? run
        : new RichTextRunInput(resolvedSource, run.font());
  }

  static CommentInput resolveComment(CommentInput comment, ExecutionInputBindings bindings)
      throws IOException {
    TextSourceInput resolvedText =
        SourceBackedPlanResolver.resolveTextSource(comment.text(), bindings, true, "comment text");
    Optional<List<RichTextRunInput>> resolvedRuns =
        comment.runs().isEmpty()
            ? Optional.empty()
            : Optional.of(resolveRuns(comment.runs().orElseThrow(), bindings));
    return sameReference(resolvedText, comment.text())
            && sameOptionalReference(resolvedRuns, comment.runs())
        ? comment
        : new CommentInput(
            resolvedText, comment.author(), comment.visible(), resolvedRuns, comment.anchor());
  }

  static CustomXmlImportInput resolveCustomXmlImport(
      CustomXmlImportInput input, ExecutionInputBindings bindings) throws IOException {
    TextSourceInput resolvedXml =
        SourceBackedPlanResolver.resolveTextSource(input.xml(), bindings, true, "custom XML");
    return sameReference(resolvedXml, input.xml())
        ? input
        : new CustomXmlImportInput(input.locator(), resolvedXml);
  }

  static SignatureLineInput resolveSignatureLine(
      SignatureLineInput signatureLine, ExecutionInputBindings bindings) throws IOException {
    if (signatureLine.plainSignature().isEmpty()) {
      return signatureLine;
    }
    PictureDataInput resolvedPlainSignature =
        resolvePictureData(signatureLine.plainSignature().orElseThrow(), bindings);
    return sameOptionalReference(
            Optional.of(resolvedPlainSignature), signatureLine.plainSignature())
        ? signatureLine
        : new SignatureLineInput(
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

  static PictureInput resolvePicture(PictureInput picture, ExecutionInputBindings bindings)
      throws IOException {
    PictureDataInput resolvedImage = resolvePictureData(picture.image(), bindings);
    Optional<TextSourceInput> resolvedDescription = picture.description();
    if (picture.description().isPresent()) {
      resolvedDescription =
          Optional.of(
              SourceBackedPlanResolver.resolveTextSource(
                  picture.description().orElseThrow(), bindings, true, "picture description"));
    }
    return sameReference(resolvedImage, picture.image())
            && sameOptionalReference(resolvedDescription, picture.description())
        ? picture
        : new PictureInput(picture.name(), resolvedImage, picture.anchor(), resolvedDescription);
  }

  private static PictureDataInput resolvePictureData(
      PictureDataInput image, ExecutionInputBindings bindings) throws IOException {
    BinarySourceInput resolvedSource =
        SourceBackedPlanResolver.resolveBinarySource(image.source(), bindings, "picture payload");
    return sameReference(resolvedSource, image.source())
        ? image
        : new PictureDataInput(image.format(), resolvedSource);
  }

  static EmbeddedObjectInput resolveEmbeddedObject(
      EmbeddedObjectInput embeddedObject, ExecutionInputBindings bindings) throws IOException {
    BinarySourceInput resolvedPayload =
        SourceBackedPlanResolver.resolveBinarySource(
            embeddedObject.payload(), bindings, "embedded-object payload");
    PictureDataInput resolvedPreview = resolvePictureData(embeddedObject.previewImage(), bindings);
    return sameReference(resolvedPayload, embeddedObject.payload())
            && sameReference(resolvedPreview, embeddedObject.previewImage())
        ? embeddedObject
        : new EmbeddedObjectInput(
            embeddedObject.name(),
            embeddedObject.label(),
            embeddedObject.fileName(),
            embeddedObject.command(),
            resolvedPayload,
            resolvedPreview,
            embeddedObject.anchor());
  }

  static ChartInput resolveChart(ChartInput chart, ExecutionInputBindings bindings)
      throws IOException {
    return SourceBackedChartInputResolver.resolveChart(chart, bindings);
  }

  static ShapeInput resolveShape(ShapeInput shape, ExecutionInputBindings bindings)
      throws IOException {
    return switch (shape) {
      case ShapeInput.SimpleShape simpleShape -> {
        Optional<TextSourceInput> resolvedText = simpleShape.text();
        if (simpleShape.text().isPresent()) {
          resolvedText =
              Optional.of(
                  SourceBackedPlanResolver.resolveTextSource(
                      simpleShape.text().orElseThrow(), bindings, true, "shape text"));
        }
        yield sameOptionalReference(resolvedText, simpleShape.text())
            ? shape
            : new ShapeInput.SimpleShape(
                simpleShape.name(),
                simpleShape.anchor(),
                simpleShape.presetGeometryToken(),
                resolvedText);
      }
      case ShapeInput.Connector _ -> shape;
    };
  }

  static DataValidationInput resolveDataValidation(
      DataValidationInput validation, ExecutionInputBindings bindings) throws IOException {
    Optional<DataValidationPromptInput> resolvedPrompt = validation.prompt();
    if (validation.prompt().isPresent()) {
      resolvedPrompt = Optional.of(resolvePrompt(validation.prompt().orElseThrow(), bindings));
    }
    Optional<DataValidationErrorAlertInput> resolvedAlert = validation.errorAlert();
    if (validation.errorAlert().isPresent()) {
      resolvedAlert =
          Optional.of(resolveErrorAlert(validation.errorAlert().orElseThrow(), bindings));
    }
    return sameOptionalReference(resolvedPrompt, validation.prompt())
            && sameOptionalReference(resolvedAlert, validation.errorAlert())
        ? validation
        : new DataValidationInput(
            validation.rule(),
            validation.allowBlank(),
            validation.suppressDropDownArrow(),
            resolvedPrompt,
            resolvedAlert);
  }

  private static DataValidationPromptInput resolvePrompt(
      DataValidationPromptInput prompt, ExecutionInputBindings bindings) throws IOException {
    TextSourceInput resolvedTitle =
        SourceBackedPlanResolver.resolveTextSource(
            prompt.title(), bindings, true, "validation prompt title");
    TextSourceInput resolvedText =
        SourceBackedPlanResolver.resolveTextSource(
            prompt.text(), bindings, true, "validation prompt text");
    return sameReference(resolvedTitle, prompt.title())
            && sameReference(resolvedText, prompt.text())
        ? prompt
        : new DataValidationPromptInput(resolvedTitle, resolvedText, prompt.showPromptBox());
  }

  private static DataValidationErrorAlertInput resolveErrorAlert(
      DataValidationErrorAlertInput alert, ExecutionInputBindings bindings) throws IOException {
    TextSourceInput resolvedTitle =
        SourceBackedPlanResolver.resolveTextSource(
            alert.title(), bindings, true, "validation error title");
    TextSourceInput resolvedText =
        SourceBackedPlanResolver.resolveTextSource(
            alert.text(), bindings, true, "validation error text");
    return sameReference(resolvedTitle, alert.title()) && sameReference(resolvedText, alert.text())
        ? alert
        : new DataValidationErrorAlertInput(
            alert.style(), resolvedTitle, resolvedText, alert.showErrorBox());
  }

  static TableInput resolveTable(TableInput table, ExecutionInputBindings bindings)
      throws IOException {
    TextSourceInput resolvedComment =
        SourceBackedPlanResolver.resolveTextSource(
            table.comment(), bindings, false, "table comment");
    return sameReference(resolvedComment, table.comment())
        ? table
        : new TableInput(
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

  static PrintLayoutInput resolvePrintLayout(
      PrintLayoutInput printLayout, ExecutionInputBindings bindings) throws IOException {
    HeaderFooterTextInput resolvedHeader = resolveHeaderFooter(printLayout.header(), bindings);
    HeaderFooterTextInput resolvedFooter = resolveHeaderFooter(printLayout.footer(), bindings);
    return sameReference(resolvedHeader, printLayout.header())
            && sameReference(resolvedFooter, printLayout.footer())
        ? printLayout
        : new PrintLayoutInput(
            printLayout.printArea(),
            printLayout.orientation(),
            printLayout.scaling(),
            printLayout.repeatingRows(),
            printLayout.repeatingColumns(),
            resolvedHeader,
            resolvedFooter,
            printLayout.setup());
  }

  static HeaderFooterTextInput resolveHeaderFooter(
      HeaderFooterTextInput text, ExecutionInputBindings bindings) throws IOException {
    TextSourceInput resolvedLeft =
        SourceBackedPlanResolver.resolveTextSource(
            text.left(), bindings, false, "header/footer left");
    TextSourceInput resolvedCenter =
        SourceBackedPlanResolver.resolveTextSource(
            text.center(), bindings, false, "header/footer center");
    TextSourceInput resolvedRight =
        SourceBackedPlanResolver.resolveTextSource(
            text.right(), bindings, false, "header/footer right");
    return sameReference(resolvedLeft, text.left())
            && sameReference(resolvedCenter, text.center())
            && sameReference(resolvedRight, text.right())
        ? text
        : new HeaderFooterTextInput(resolvedLeft, resolvedCenter, resolvedRight);
  }
}
