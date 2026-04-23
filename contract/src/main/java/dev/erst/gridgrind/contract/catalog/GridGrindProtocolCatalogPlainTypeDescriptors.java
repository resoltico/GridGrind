package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.ArrayFormulaInput;
import dev.erst.gridgrind.contract.dto.ArrayFormulaReport;
import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnInput;
import dev.erst.gridgrind.contract.dto.AutofilterFilterCriterionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortConditionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateInput;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.CellAlignmentInput;
import dev.erst.gridgrind.contract.dto.CellAlignmentReport;
import dev.erst.gridgrind.contract.dto.CellBorderInput;
import dev.erst.gridgrind.contract.dto.CellBorderReport;
import dev.erst.gridgrind.contract.dto.CellBorderSideInput;
import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellColorReport;
import dev.erst.gridgrind.contract.dto.CellFillInput;
import dev.erst.gridgrind.contract.dto.CellFillReport;
import dev.erst.gridgrind.contract.dto.CellFontInput;
import dev.erst.gridgrind.contract.dto.CellFontReport;
import dev.erst.gridgrind.contract.dto.CellGradientFillInput;
import dev.erst.gridgrind.contract.dto.CellGradientFillReport;
import dev.erst.gridgrind.contract.dto.CellGradientStopInput;
import dev.erst.gridgrind.contract.dto.CellGradientStopReport;
import dev.erst.gridgrind.contract.dto.CellProtectionInput;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.CellStyleInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.ColorInput;
import dev.erst.gridgrind.contract.dto.CommentAnchorInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingThresholdInput;
import dev.erst.gridgrind.contract.dto.CustomXmlDataBindingReport;
import dev.erst.gridgrind.contract.dto.CustomXmlExportReport;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.CustomXmlLinkedCellReport;
import dev.erst.gridgrind.contract.dto.CustomXmlLinkedTableReport;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingReport;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderSideInput;
import dev.erst.gridgrind.contract.dto.DifferentialStyleInput;
import dev.erst.gridgrind.contract.dto.DrawingMarkerInput;
import dev.erst.gridgrind.contract.dto.DrawingMarkerReport;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.ExecutionJournalInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.FormulaExternalWorkbookInput;
import dev.erst.gridgrind.contract.dto.FormulaUdfFunctionInput;
import dev.erst.gridgrind.contract.dto.FormulaUdfToolpackInput;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.HeaderFooterTextInput;
import dev.erst.gridgrind.contract.dto.IgnoredErrorInput;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionInput;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionReport;
import dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureInput;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureReport;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.PrintMarginsInput;
import dev.erst.gridgrind.contract.dto.PrintSetupInput;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.SheetDefaultsInput;
import dev.erst.gridgrind.contract.dto.SheetDisplayInput;
import dev.erst.gridgrind.contract.dto.SheetOutlineSummaryInput;
import dev.erst.gridgrind.contract.dto.SheetPresentationInput;
import dev.erst.gridgrind.contract.dto.SheetProtectionSettings;
import dev.erst.gridgrind.contract.dto.SignatureLineInput;
import dev.erst.gridgrind.contract.dto.TableColumnInput;
import dev.erst.gridgrind.contract.dto.TableColumnReport;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import java.util.List;

/** Owns the plain field-shape descriptor registry used by the protocol catalog. */
@SuppressWarnings("PMD.ExcessiveImports")
final class GridGrindProtocolCatalogPlainTypeDescriptors {
  private GridGrindProtocolCatalogPlainTypeDescriptors() {}

  static final List<CatalogPlainTypeDescriptor> PLAIN_TYPE_DESCRIPTORS =
      List.of(
          plainTypeDescriptor(
              "executionJournalType",
              ExecutionJournal.class,
              "ExecutionJournal",
              "Structured execution telemetry returned on every success and failure response,"
                  + " including validation, open, calculation, step, persistence, and close"
                  + " phases.",
              List.of(
                  "planId",
                  "level",
                  "source",
                  "persistence",
                  "validation",
                  "open",
                  "calculation",
                  "persistencePhase",
                  "close",
                  "steps",
                  "warnings",
                  "outcome",
                  "events")),
          plainTypeDescriptor(
              "executionJournalSourceSummaryType",
              ExecutionJournal.SourceSummary.class,
              "ExecutionJournalSourceSummary",
              "Journal summary of the authored workbook source.",
              List.of("path")),
          plainTypeDescriptor(
              "executionJournalPersistenceSummaryType",
              ExecutionJournal.PersistenceSummary.class,
              "ExecutionJournalPersistenceSummary",
              "Journal summary of the authored persistence policy.",
              List.of("path")),
          plainTypeDescriptor(
              "executionJournalPhaseType",
              ExecutionJournal.Phase.class,
              "ExecutionJournalPhase",
              "One timed execution phase with status, timestamps, and duration.",
              List.of("startedAt", "finishedAt")),
          plainTypeDescriptor(
              "executionJournalStepType",
              ExecutionJournal.Step.class,
              "ExecutionJournalStep",
              "Per-step execution telemetry with resolved targets, timing, outcome,"
                  + " and optional failure detail.",
              List.of("failure")),
          plainTypeDescriptor(
              "executionJournalTargetType",
              ExecutionJournal.Target.class,
              "ExecutionJournalTarget",
              "One canonical target label recorded inside a step journal.",
              List.of()),
          plainTypeDescriptor(
              "executionJournalFailureClassificationType",
              ExecutionJournal.FailureClassification.class,
              "ExecutionJournalFailureClassification",
              "Structured problem-code classification for one failed step.",
              List.of()),
          plainTypeDescriptor(
              "executionJournalCalculationType",
              ExecutionJournal.Calculation.class,
              "ExecutionJournalCalculation",
              "Top-level calculation preflight and execution timings for one request.",
              List.of()),
          plainTypeDescriptor(
              "executionJournalOutcomeType",
              ExecutionJournal.Outcome.class,
              "ExecutionJournalOutcome",
              "Final outcome summary for one execution journal.",
              List.of("failedStepIndex", "failedStepId", "failureCode")),
          plainTypeDescriptor(
              "executionJournalEventType",
              ExecutionJournal.Event.class,
              "ExecutionJournalEvent",
              "Fine-grained verbose execution event emitted for live CLI rendering.",
              List.of("stepIndex", "stepId")),
          plainTypeDescriptor(
              "requestWarningType",
              RequestWarning.class,
              "RequestWarning",
              "Non-fatal authored-plan warning surfaced on success and echoed inside the execution journal.",
              List.of()),
          plainTypeDescriptor(
              "executionPolicyInputType",
              ExecutionPolicyInput.class,
              "ExecutionPolicyInput",
              GridGrindContractText.executionPolicyInputSummary(),
              List.of("mode", "journal", "calculation")),
          plainTypeDescriptor(
              "calculationPolicyInputType",
              CalculationPolicyInput.class,
              "CalculationPolicyInput",
              GridGrindContractText.calculationPolicyInputSummary(),
              List.of("strategy")),
          plainTypeDescriptor(
              "executionModeInputType",
              ExecutionModeInput.class,
              "ExecutionModeInput",
              GridGrindContractText.executionModeInputSummary(),
              List.of("readMode", "writeMode")),
          plainTypeDescriptor(
              "executionJournalInputType",
              ExecutionJournalInput.class,
              "ExecutionJournalInput",
              GridGrindContractText.executionJournalInputSummary(),
              List.of("level")),
          plainTypeDescriptor(
              "calculationReportType",
              CalculationReport.class,
              "CalculationReport",
              "Structured calculation policy, preflight classification, and execution outcome"
                  + " returned on every success and failure response.",
              List.of("preflight")),
          plainTypeDescriptor(
              "calculationPreflightType",
              CalculationReport.Preflight.class,
              "CalculationPreflightReport",
              "Formula capability classification captured before server-side evaluation begins.",
              List.of()),
          plainTypeDescriptor(
              "calculationSummaryType",
              CalculationReport.Summary.class,
              "CalculationPreflightSummary",
              "Aggregate counts for evaluable, unevaluable, and unparseable formulas.",
              List.of()),
          plainTypeDescriptor(
              "formulaCapabilityType",
              CalculationReport.FormulaCapability.class,
              "FormulaCapabilityReport",
              "One classified formula-cell capability entry from calculation preflight.",
              List.of("problemCode", "message")),
          plainTypeDescriptor(
              "calculationExecutionType",
              CalculationReport.Execution.class,
              "CalculationExecutionReport",
              "Post-execution outcome for the authored calculation policy.",
              List.of("message")),
          plainTypeDescriptor(
              "formulaEnvironmentInputType",
              FormulaEnvironmentInput.class,
              "FormulaEnvironmentInput",
              "Request-scoped formula-evaluation environment covering external workbook bindings,"
                  + " missing-workbook policy, and template-backed UDF toolpacks.",
              List.of("externalWorkbooks", "missingWorkbookPolicy", "udfToolpacks")),
          plainTypeDescriptor(
              "ooxmlOpenSecurityInputType",
              OoxmlOpenSecurityInput.class,
              "OoxmlOpenSecurityInput",
              "Optional OOXML package-open settings for encrypted existing workbook sources."
                  + " password unlocks the encrypted OOXML package before GridGrind opens the"
                  + " inner .xlsx workbook.",
              List.of("password")),
          plainTypeDescriptor(
              "ooxmlPersistenceSecurityInputType",
              OoxmlPersistenceSecurityInput.class,
              "OoxmlPersistenceSecurityInput",
              "Optional OOXML package-security settings applied during persistence."
                  + " Supply encryption, signature, or both.",
              List.of("encryption", "signature")),
          plainTypeDescriptor(
              "ooxmlEncryptionInputType",
              OoxmlEncryptionInput.class,
              "OoxmlEncryptionInput",
              "OOXML package-encryption settings for workbook persistence."
                  + " mode defaults to AGILE when omitted.",
              List.of("mode")),
          plainTypeDescriptor(
              "ooxmlSignatureInputType",
              OoxmlSignatureInput.class,
              "OoxmlSignatureInput",
              "OOXML package-signing settings for workbook persistence."
                  + " pkcs12Path follows the request-owned path rule."
                  + " "
                  + GridGrindContractText.requestOwnedPathResolutionSummary()
                  + " keyPassword defaults to keystorePassword and digestAlgorithm defaults to"
                  + " SHA256 when omitted."
                  + " alias may be omitted only when the keystore contains exactly one"
                  + " signable private-key entry.",
              List.of("keyPassword", "alias", "digestAlgorithm", "description")),
          plainTypeDescriptor(
              "ooxmlPackageSecurityReportType",
              OoxmlPackageSecurityReport.class,
              "OoxmlPackageSecurityReport",
              "Factual OOXML package-security report covering encryption and package signatures.",
              List.of()),
          plainTypeDescriptor(
              "ooxmlEncryptionReportType",
              OoxmlEncryptionReport.class,
              "OoxmlEncryptionReport",
              "Factual OOXML package-encryption report for one workbook package."
                  + " Detail fields are present only when encrypted=true.",
              List.of()),
          plainTypeDescriptor(
              "ooxmlSignatureReportType",
              OoxmlSignatureReport.class,
              "OoxmlSignatureReport",
              "Factual OOXML package-signature report for one signature part."
                  + " state reflects the currently loaded workbook package, including"
                  + " INVALIDATED_BY_MUTATION for source signatures after in-memory edits.",
              List.of()),
          plainTypeDescriptor(
              "formulaExternalWorkbookInputType",
              FormulaExternalWorkbookInput.class,
              "FormulaExternalWorkbookInput",
              "One external workbook binding keyed by the workbook name used inside formulas."
                  + " path follows the request-owned path rule."
                  + " "
                  + GridGrindContractText.requestOwnedPathResolutionSummary(),
              List.of()),
          plainTypeDescriptor(
              "formulaUdfToolpackInputType",
              FormulaUdfToolpackInput.class,
              "FormulaUdfToolpackInput",
              "One named collection of template-backed user-defined functions.",
              List.of()),
          plainTypeDescriptor(
              "formulaUdfFunctionInputType",
              FormulaUdfFunctionInput.class,
              "FormulaUdfFunctionInput",
              "One template-backed user-defined function."
                  + " formulaTemplate may reference ARG1, ARG2, and higher placeholders."
                  + " maximumArgumentCount defaults to minimumArgumentCount when omitted.",
              List.of("maximumArgumentCount")),
          plainTypeDescriptor(
              "qualifiedCellAddressType",
              dev.erst.gridgrind.contract.selector.CellSelector.QualifiedAddress.class,
              "CellSelector.QualifiedAddress",
              "One workbook-qualified cell address used by selector-based targeted cell workflows.",
              List.of()),
          plainTypeDescriptor(
              "drawingMarkerInputType",
              DrawingMarkerInput.class,
              "DrawingMarkerInput",
              "One zero-based drawing marker with explicit column, row, and in-cell offsets.",
              List.of()),
          plainTypeDescriptor(
              "arrayFormulaInputType",
              ArrayFormulaInput.class,
              "ArrayFormulaInput",
              "One authored array formula bound to a contiguous single-cell or multi-cell range."
                  + " Leading = or {=...} wrappers normalize away for inline sources.",
              List.of()),
          plainTypeDescriptor(
              "customXmlMappingLocatorType",
              CustomXmlMappingLocator.class,
              "CustomXmlMappingLocator",
              "One locator for an existing workbook custom-XML mapping."
                  + " Supply mapId, name, or both; the locator must resolve to exactly one"
                  + " existing mapping.",
              List.of("mapId", "name")),
          plainTypeDescriptor(
              "customXmlImportInputType",
              CustomXmlImportInput.class,
              "CustomXmlImportInput",
              "One custom-XML import payload targeting an existing workbook mapping plus the XML"
                  + " content to import.",
              List.of()),
          plainTypeDescriptor(
              "chartInputType",
              ChartInput.class,
              "ChartInput",
              "One authored chart with a drawing anchor, chart-level presentation state,"
                  + " and one or more plots.",
              List.of("title", "legend", "displayBlanksAs", "plotOnlyVisibleCells")),
          plainTypeDescriptor(
              "chartAxisInputType",
              ChartInput.Axis.class,
              "ChartAxisInput",
              "One authored chart axis used by a chart plot."
                  + " visible defaults to true when omitted.",
              List.of("visible")),
          plainTypeDescriptor(
              "chartSeriesInputType",
              ChartInput.Series.class,
              "ChartSeriesInput",
              "One authored chart series with a title plus category and value data sources."
                  + " smooth, marker, and explosion fields are optional by chart family.",
              List.of("title", "smooth", "markerStyle", "markerSize", "explosion")),
          plainTypeDescriptor(
              "pictureDataInputType",
              PictureDataInput.class,
              "PictureDataInput",
              "One picture payload with explicit format and base64-encoded binary data.",
              List.of()),
          plainTypeDescriptor(
              "pictureInputType",
              PictureInput.class,
              "PictureInput",
              "Named picture-authoring payload for SET_PICTURE.",
              List.of("description")),
          plainTypeDescriptor(
              "signatureLineInputType",
              SignatureLineInput.class,
              "SignatureLineInput",
              "Named signature-line authoring payload for SET_SIGNATURE_LINE."
                  + " allowComments defaults to true when omitted and plainSignature is optional,"
                  + " but caption or suggested signer metadata must still be present.",
              List.of(
                  "allowComments",
                  "signingInstructions",
                  "suggestedSigner",
                  "suggestedSigner2",
                  "suggestedSignerEmail",
                  "caption",
                  "invalidStamp",
                  "plainSignature")),
          plainTypeDescriptor(
              "shapeInputType",
              ShapeInput.class,
              "ShapeInput",
              "Named simple-shape or connector authoring payload for SET_SHAPE."
                  + " kind is limited to the authored drawing shape family."
                  + " presetGeometryToken defaults to rect for SIMPLE_SHAPE when omitted."
                  + " Invalid presetGeometryToken values are rejected non-mutatingly.",
              List.of("presetGeometryToken", "text")),
          plainTypeDescriptor(
              "embeddedObjectInputType",
              EmbeddedObjectInput.class,
              "EmbeddedObjectInput",
              "Named embedded-object authoring payload for SET_EMBEDDED_OBJECT."
                  + " base64Data holds the embedded package bytes and previewImage holds the"
                  + " visible preview raster.",
              List.of()),
          plainTypeDescriptor(
              "commentInputType",
              CommentInput.class,
              "CommentInput",
              "Comment payload attached to one cell."
                  + " Comments can carry ordered rich-text runs and an explicit anchor box.",
              List.of("visible", "runs", "anchor")),
          plainTypeDescriptor(
              "commentAnchorInputType",
              CommentAnchorInput.class,
              "CommentAnchorInput",
              "Explicit comment-anchor bounds measured in zero-based column and row indexes.",
              List.of()),
          plainTypeDescriptor(
              "namedRangeTargetType",
              NamedRangeTarget.class,
              "NamedRangeTarget",
              "Named-range target payload."
                  + " Supply either sheetName plus range, or formula by itself.",
              List.of("sheetName", "range", "formula")),
          plainTypeDescriptor(
              "sheetProtectionSettingsType",
              SheetProtectionSettings.class,
              "SheetProtectionSettings",
              "Supported sheet-protection lock flags authored and reported by GridGrind.",
              List.of()),
          plainTypeDescriptor(
              "cellStyleInputType",
              CellStyleInput.class,
              "CellStyleInput",
              "Style patch applied to a cell or range; at least one field must be set."
                  + " Colors use #RRGGBB hex and style subgroups are nested explicitly.",
              List.of("numberFormat", "alignment", "font", "fill", "border", "protection")),
          plainTypeDescriptor(
              "cellAlignmentInputType",
              CellAlignmentInput.class,
              "CellAlignmentInput",
              "Alignment patch for cell styling; at least one field must be set."
                  + " textRotation uses XSSF's explicit 0-180 degree scale and indentation uses"
                  + " Excel's 0-250 cell-indent range.",
              List.of(
                  "wrapText",
                  "horizontalAlignment",
                  "verticalAlignment",
                  "textRotation",
                  "indentation")),
          plainTypeDescriptor(
              "cellFontInputType",
              CellFontInput.class,
              "CellFontInput",
              "Font patch for cell styling; at least one field must be set."
                  + " Colors can use RGB, theme, indexed, and tint semantics.",
              List.of(
                  "bold",
                  "italic",
                  "fontName",
                  "fontHeight",
                  "fontColor",
                  "fontColorTheme",
                  "fontColorIndexed",
                  "fontColorTint",
                  "underline",
                  "strikeout")),
          plainTypeDescriptor(
              "colorInputType",
              ColorInput.class,
              "ColorInput",
              "Color payload preserving RGB, theme, indexed, and tint semantics."
                  + " At least one of rgb, theme, or indexed must be supplied.",
              List.of("rgb", "theme", "indexed", "tint")),
          plainTypeDescriptor(
              "richTextRunInputType",
              RichTextRunInput.class,
              "RichTextRunInput",
              "One ordered rich-text run for a string cell."
                  + " text must be non-empty; font is an optional override patch."
                  + " The ordered run texts concatenate to the stored plain string value.",
              List.of("font")),
          plainTypeDescriptor(
              "cellFillInputType",
              CellFillInput.class,
              "CellFillInput",
              "Fill patch for cell styling. pattern controls solid and patterned fills;"
                  + " colors can use RGB, theme, indexed, and tint semantics."
                  + " gradient is mutually exclusive with patterned fill fields.",
              List.of(
                  "pattern",
                  "foregroundColor",
                  "foregroundColorTheme",
                  "foregroundColorIndexed",
                  "foregroundColorTint",
                  "backgroundColor",
                  "backgroundColorTheme",
                  "backgroundColorIndexed",
                  "backgroundColorTint",
                  "gradient")),
          plainTypeDescriptor(
              "cellGradientFillInputType",
              CellGradientFillInput.class,
              "CellGradientFillInput",
              "Gradient fill payload for cell-style authoring."
                  + " LINEAR gradients use degree, PATH gradients use left/right/top/bottom,"
                  + " and the two geometry modes must not be mixed."
                  + " stops must contain at least two entries.",
              List.of("type", "degree", "left", "right", "top", "bottom")),
          plainTypeDescriptor(
              "cellGradientStopInputType",
              CellGradientStopInput.class,
              "CellGradientStopInput",
              "One gradient stop with a normalized position between 0.0 and 1.0.",
              List.of()),
          plainTypeDescriptor(
              "cellBorderInputType",
              CellBorderInput.class,
              "CellBorderInput",
              "Border patch for cell styling; at least one side must be set."
                  + " Use 'all' as shorthand for all four sides.",
              List.of("all", "top", "right", "bottom", "left")),
          plainTypeDescriptor(
              "cellBorderSideInputType",
              CellBorderSideInput.class,
              "CellBorderSideInput",
              "One border side defined by its border style and optional color semantics.",
              List.of("style", "color", "colorTheme", "colorIndexed", "colorTint")),
          plainTypeDescriptor(
              "cellProtectionInputType",
              CellProtectionInput.class,
              "CellProtectionInput",
              "Cell protection patch; at least one field must be set."
                  + " These flags matter when sheet protection is enabled.",
              List.of("locked", "hiddenFormula")),
          plainTypeDescriptor(
              "dataValidationInputType",
              DataValidationInput.class,
              "DataValidationInput",
              "Supported data-validation definition attached to one sheet range.",
              List.of("allowBlank", "suppressDropDownArrow", "prompt", "errorAlert")),
          plainTypeDescriptor(
              "dataValidationPromptInputType",
              DataValidationPromptInput.class,
              "DataValidationPromptInput",
              "Optional prompt-box configuration shown when a validated cell is selected.",
              List.of("showPromptBox")),
          plainTypeDescriptor(
              "dataValidationErrorAlertInputType",
              DataValidationErrorAlertInput.class,
              "DataValidationErrorAlertInput",
              "Optional error-box configuration shown when invalid data is entered.",
              List.of("showErrorBox")),
          plainTypeDescriptor(
              "autofilterCustomConditionInputType",
              AutofilterFilterCriterionInput.CustomConditionInput.class,
              "AutofilterCustomConditionInput",
              "One comparator-value pair nested inside a custom autofilter criterion.",
              List.of()),
          plainTypeDescriptor(
              "autofilterFilterColumnInputType",
              AutofilterFilterColumnInput.class,
              "AutofilterFilterColumnInput",
              "One authored autofilter filter-column payload with an explicit column criterion.",
              List.of("showButton")),
          plainTypeDescriptor(
              "autofilterSortConditionInputType",
              AutofilterSortConditionInput.class,
              "AutofilterSortConditionInput",
              "One authored sort condition nested inside an autofilter sort state.",
              List.of("sortBy", "color", "iconId")),
          plainTypeDescriptor(
              "autofilterSortStateInputType",
              AutofilterSortStateInput.class,
              "AutofilterSortStateInput",
              "Authored autofilter sort-state payload with one or more ordered sort conditions.",
              List.of("caseSensitive", "columnSort", "sortMethod")),
          plainTypeDescriptor(
              "conditionalFormattingBlockInputType",
              ConditionalFormattingBlockInput.class,
              "ConditionalFormattingBlockInput",
              "One authored conditional-formatting block with ordered target ranges and rules."
                  + " rules must not be empty; ranges must be unique.",
              List.of()),
          plainTypeDescriptor(
              "conditionalFormattingThresholdInputType",
              ConditionalFormattingThresholdInput.class,
              "ConditionalFormattingThresholdInput",
              "Threshold payload shared by authored advanced conditional-formatting rules.",
              List.of("formula", "value")),
          plainTypeDescriptor(
              "headerFooterTextInputType",
              HeaderFooterTextInput.class,
              "HeaderFooterTextInput",
              "Plain left, center, and right header or footer text segments."
                  + " Null fields default to empty string.",
              List.of("left", "center", "right")),
          plainTypeDescriptor(
              "differentialStyleInputType",
              DifferentialStyleInput.class,
              "DifferentialStyleInput",
              "Differential style payload used by authored conditional-formatting rules."
                  + " At least one field must be set. Colors use #RRGGBB hex.",
              List.of(
                  "numberFormat",
                  "bold",
                  "italic",
                  "fontHeight",
                  "fontColor",
                  "underline",
                  "strikeout",
                  "fillColor",
                  "border")),
          plainTypeDescriptor(
              "differentialBorderInputType",
              DifferentialBorderInput.class,
              "DifferentialBorderInput",
              "Conditional-formatting differential border patch; at least one side must be set."
                  + " Use 'all' as shorthand for all four sides.",
              List.of("all", "top", "right", "bottom", "left")),
          plainTypeDescriptor(
              "differentialBorderSideInputType",
              DifferentialBorderSideInput.class,
              "DifferentialBorderSideInput",
              "One conditional-formatting differential border side defined by style and optional"
                  + " color.",
              List.of()),
          plainTypeDescriptor(
              "ignoredErrorInputType",
              IgnoredErrorInput.class,
              "IgnoredErrorInput",
              "One ignored-error block anchored to one A1-style range plus one or more"
                  + " ignored-error families.",
              List.of()),
          plainTypeDescriptor(
              "printLayoutInputType",
              PrintLayoutInput.class,
              "PrintLayoutInput",
              "Authoritative supported print-layout payload for one SET_PRINT_LAYOUT request."
                  + " All fields are optional and normalize to defaults when omitted.",
              List.of(
                  "printArea",
                  "orientation",
                  "scaling",
                  "repeatingRows",
                  "repeatingColumns",
                  "header",
                  "footer",
                  "setup")),
          plainTypeDescriptor(
              "printMarginsInputType",
              PrintMarginsInput.class,
              "PrintMarginsInput",
              "Explicit print margins measured in the workbook's stored inch-based values.",
              List.of()),
          plainTypeDescriptor(
              "printSetupInputType",
              PrintSetupInput.class,
              "PrintSetupInput",
              "Advanced page-setup payload nested under print-layout authoring."
                  + " All fields are optional and normalize to defaults when omitted.",
              List.of(
                  "margins",
                  "printGridlines",
                  "horizontallyCentered",
                  "verticallyCentered",
                  "paperSize",
                  "draft",
                  "blackAndWhite",
                  "copies",
                  "useFirstPageNumber",
                  "firstPageNumber",
                  "rowBreaks",
                  "columnBreaks")),
          plainTypeDescriptor(
              "sheetDefaultsInputType",
              SheetDefaultsInput.class,
              "SheetDefaultsInput",
              "Default row and column sizing authored as part of sheet-presentation state."
                  + " All fields are optional and normalize to defaults when omitted."
                  + " defaultColumnWidth must be > 0 and <= 255;"
                  + " defaultRowHeightPoints must be > 0 and <= "
                  + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS
                  + ".",
              List.of("defaultColumnWidth", "defaultRowHeightPoints")),
          plainTypeDescriptor(
              "sheetDisplayInputType",
              SheetDisplayInput.class,
              "SheetDisplayInput",
              "Screen-facing sheet display flags authored as part of sheet-presentation state."
                  + " All fields are optional and normalize to defaults when omitted.",
              List.of(
                  "displayGridlines",
                  "displayZeros",
                  "displayRowColHeadings",
                  "displayFormulas",
                  "rightToLeft")),
          plainTypeDescriptor(
              "sheetOutlineSummaryInputType",
              SheetOutlineSummaryInput.class,
              "SheetOutlineSummaryInput",
              "Outline-summary placement authored as part of sheet-presentation state."
                  + " All fields are optional and normalize to defaults when omitted.",
              List.of("rowSumsBelow", "rowSumsRight")),
          plainTypeDescriptor(
              "sheetPresentationInputType",
              SheetPresentationInput.class,
              "SheetPresentationInput",
              "Authoritative sheet-presentation payload for one SET_SHEET_PRESENTATION request."
                  + " All fields are optional and normalize to defaults or clear state when"
                  + " omitted.",
              List.of("display", "tabColor", "outlineSummary", "sheetDefaults", "ignoredErrors")),
          plainTypeDescriptor(
              "pivotTableInputType",
              PivotTableInput.class,
              "PivotTableInput",
              "Workbook-global pivot-table definition for one SET_PIVOT_TABLE request."
                  + " Source-column assignments across rowLabels, columnLabels, reportFilters,"
                  + " and dataFields must be disjoint."
                  + " reportFilters require anchor.topLeftAddress on row 3 or lower.",
              List.of("rowLabels", "columnLabels", "reportFilters")),
          plainTypeDescriptor(
              "pivotTableAnchorInputType",
              PivotTableInput.Anchor.class,
              "PivotTableAnchorInput",
              "Top-left anchor for a pivot table rendered on its destination sheet."
                  + " The address must be a single-cell A1 reference.",
              List.of()),
          plainTypeDescriptor(
              "pivotTableDataFieldInputType",
              PivotTableInput.DataField.class,
              "PivotTableDataFieldInput",
              "One authored pivot data field bound to a source column and aggregation function."
                  + " displayName defaults to sourceColumnName when omitted.",
              List.of("displayName", "valueFormat")),
          plainTypeDescriptor(
              "tableColumnInputType",
              TableColumnInput.class,
              "TableColumnInput",
              "Advanced table-column metadata applied by zero-based ordinal column index.",
              List.of(
                  "uniqueName", "totalsRowLabel", "totalsRowFunction", "calculatedColumnFormula")),
          plainTypeDescriptor(
              "tableInputType",
              TableInput.class,
              "TableInput",
              "Workbook-global table definition for one SET_TABLE request.",
              List.of(
                  "showTotalsRow",
                  "hasAutofilter",
                  "comment",
                  "published",
                  "insertRow",
                  "insertRowShift",
                  "headerRowCellStyle",
                  "dataCellStyle",
                  "totalsRowCellStyle",
                  "columns")),
          plainTypeDescriptor(
              "workbookProtectionInputType",
              WorkbookProtectionInput.class,
              "WorkbookProtectionInput",
              "Workbook-protection payload covering workbook and revisions lock state plus"
                  + " optional passwords.",
              List.of(
                  "structureLocked",
                  "windowsLocked",
                  "revisionsLocked",
                  "workbookPassword",
                  "revisionsPassword")),
          plainTypeDescriptor(
              "workbookProtectionReportType",
              WorkbookProtectionReport.class,
              "WorkbookProtectionReport",
              "Exact workbook-protection report covering structure, windows, revisions,"
                  + " and password-hash presence flags.",
              List.of()),
          plainTypeDescriptor(
              "sheetSummaryReportType",
              GridGrindResponse.SheetSummaryReport.class,
              "SheetSummaryReport",
              "Exact sheet summary report including visibility, protection, and structural counts.",
              List.of()),
          plainTypeDescriptor(
              "cellStyleReportType",
              GridGrindResponse.CellStyleReport.class,
              "CellStyleReport",
              "Exact effective cell-style report used by style assertions.",
              List.of()),
          plainTypeDescriptor(
              "cellAlignmentReportType",
              CellAlignmentReport.class,
              "CellAlignmentReport",
              "Exact cell-alignment report.",
              List.of()),
          plainTypeDescriptor(
              "cellFontReportType",
              CellFontReport.class,
              "CellFontReport",
              "Exact cell-font report.",
              List.of("fontColor")),
          plainTypeDescriptor(
              "cellFillReportType",
              CellFillReport.class,
              "CellFillReport",
              "Exact cell-fill report including pattern, colors, or gradient payload.",
              List.of("foregroundColor", "backgroundColor", "gradient")),
          plainTypeDescriptor(
              "cellBorderReportType",
              CellBorderReport.class,
              "CellBorderReport",
              "Exact four-sided cell-border report.",
              List.of()),
          plainTypeDescriptor(
              "cellBorderSideReportType",
              CellBorderSideReport.class,
              "CellBorderSideReport",
              "Exact one-sided cell-border report.",
              List.of("color")),
          plainTypeDescriptor(
              "cellProtectionReportType",
              CellProtectionReport.class,
              "CellProtectionReport",
              "Exact cell-protection report.",
              List.of()),
          plainTypeDescriptor(
              "cellColorReportType",
              CellColorReport.class,
              "CellColorReport",
              "Exact workbook color report preserving RGB, theme, indexed, and tint semantics.",
              List.of("rgb", "theme", "indexed", "tint")),
          plainTypeDescriptor(
              "fontHeightReportType",
              FontHeightReport.class,
              "FontHeightReport",
              "Exact font-height report expressed in twips and points.",
              List.of()),
          plainTypeDescriptor(
              "cellGradientFillReportType",
              CellGradientFillReport.class,
              "CellGradientFillReport",
              "Exact gradient-fill report with geometry and stops.",
              List.of("degree", "left", "right", "top", "bottom")),
          plainTypeDescriptor(
              "cellGradientStopReportType",
              CellGradientStopReport.class,
              "CellGradientStopReport",
              "Exact gradient stop report.",
              List.of()),
          plainTypeDescriptor(
              "tableEntryReportType",
              TableEntryReport.class,
              "TableEntryReport",
              "Exact workbook table report used by table-facts assertions.",
              List.of("comment", "headerRowCellStyle", "dataCellStyle", "totalsRowCellStyle")),
          plainTypeDescriptor(
              "tableColumnReportType",
              TableColumnReport.class,
              "TableColumnReport",
              "Exact table-column report.",
              List.of(
                  "uniqueName", "totalsRowLabel", "totalsRowFunction", "calculatedColumnFormula")),
          plainTypeDescriptor(
              "drawingMarkerReportType",
              DrawingMarkerReport.class,
              "DrawingMarkerReport",
              "Exact cell-relative drawing marker report.",
              List.of()),
          plainTypeDescriptor(
              "pivotTableAnchorReportType",
              PivotTableReport.Anchor.class,
              "PivotTableAnchorReport",
              "Exact pivot-table anchor report.",
              List.of()),
          plainTypeDescriptor(
              "pivotTableFieldReportType",
              PivotTableReport.Field.class,
              "PivotTableFieldReport",
              "Exact pivot field report bound to one source column.",
              List.of()),
          plainTypeDescriptor(
              "pivotTableDataFieldReportType",
              PivotTableReport.DataField.class,
              "PivotTableDataFieldReport",
              "Exact pivot data-field report.",
              List.of("valueFormat")),
          plainTypeDescriptor(
              "arrayFormulaReportType",
              ArrayFormulaReport.class,
              "ArrayFormulaReport",
              "One factual array-formula group report returned by GET_ARRAY_FORMULAS.",
              List.of()),
          plainTypeDescriptor(
              "customXmlMappingReportType",
              CustomXmlMappingReport.class,
              "CustomXmlMappingReport",
              "One factual workbook custom-XML mapping report.",
              List.of(
                  "schemaNamespace",
                  "schemaLanguage",
                  "schemaReference",
                  "schemaXml",
                  "dataBinding")),
          plainTypeDescriptor(
              "customXmlDataBindingReportType",
              CustomXmlDataBindingReport.class,
              "CustomXmlDataBindingReport",
              "Optional custom-XML data-binding metadata attached to one workbook mapping.",
              List.of("dataBindingName", "fileBinding", "connectionId", "fileBindingName")),
          plainTypeDescriptor(
              "customXmlLinkedCellReportType",
              CustomXmlLinkedCellReport.class,
              "CustomXmlLinkedCellReport",
              "One single-cell binding linked to a custom-XML mapping.",
              List.of()),
          plainTypeDescriptor(
              "customXmlLinkedTableReportType",
              CustomXmlLinkedTableReport.class,
              "CustomXmlLinkedTableReport",
              "One XML-mapped table linked to a custom-XML mapping.",
              List.of()),
          plainTypeDescriptor(
              "customXmlExportReportType",
              CustomXmlExportReport.class,
              "CustomXmlExportReport",
              "One exported custom-XML mapping payload plus the factual mapping metadata used to"
                  + " produce it.",
              List.of()),
          plainTypeDescriptor(
              "chartReportType",
              ChartReport.class,
              "ChartReport",
              "One factual chart report with chart-level presentation state and one or more"
                  + " plots.",
              List.of()),
          plainTypeDescriptor(
              "chartAxisReportType",
              ChartReport.Axis.class,
              "ChartAxisReport",
              "Exact chart-axis report.",
              List.of()),
          plainTypeDescriptor(
              "chartSeriesReportType",
              ChartReport.Series.class,
              "ChartSeriesReport",
              "Exact chart-series report."
                  + " smooth, marker, and explosion fields are populated only when the"
                  + " stored plot family supports them.",
              List.of("smooth", "markerStyle", "markerSize", "explosion")));

  private static CatalogPlainTypeDescriptor plainTypeDescriptor(
      String group,
      Class<? extends Record> recordType,
      String id,
      String summary,
      List<String> optionalFields) {
    return GridGrindProtocolCatalog.plainTypeDescriptor(
        group, recordType, id, summary, optionalFields);
  }
}
