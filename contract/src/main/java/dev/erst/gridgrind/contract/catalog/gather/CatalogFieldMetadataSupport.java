package dev.erst.gridgrind.contract.catalog.gather;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.catalog.FieldEntry;
import dev.erst.gridgrind.contract.catalog.FieldRequirement;
import dev.erst.gridgrind.contract.catalog.FieldShape;
import dev.erst.gridgrind.contract.catalog.ScalarType;
import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnInput;
import dev.erst.gridgrind.contract.dto.AutofilterFilterCriterionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortConditionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateInput;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
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
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellProtectionInput;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.CellStyleInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.ColorInput;
import dev.erst.gridgrind.contract.dto.CommentAnchorInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingThresholdInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderSideInput;
import dev.erst.gridgrind.contract.dto.DifferentialStyleInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorReport;
import dev.erst.gridgrind.contract.dto.DrawingMarkerInput;
import dev.erst.gridgrind.contract.dto.DrawingMarkerReport;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.ExecutionJournalInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FontHeightInput;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.FormulaExternalWorkbookInput;
import dev.erst.gridgrind.contract.dto.FormulaUdfFunctionInput;
import dev.erst.gridgrind.contract.dto.FormulaUdfToolpackInput;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.HeaderFooterTextInput;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.IgnoredErrorInput;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionInput;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionReport;
import dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureInput;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureReport;
import dev.erst.gridgrind.contract.dto.PaneInput;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.PrintAreaInput;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.PrintMarginsInput;
import dev.erst.gridgrind.contract.dto.PrintScalingInput;
import dev.erst.gridgrind.contract.dto.PrintSetupInput;
import dev.erst.gridgrind.contract.dto.PrintTitleColumnsInput;
import dev.erst.gridgrind.contract.dto.PrintTitleRowsInput;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.SheetCopyPosition;
import dev.erst.gridgrind.contract.dto.SheetDefaultsInput;
import dev.erst.gridgrind.contract.dto.SheetDisplayInput;
import dev.erst.gridgrind.contract.dto.SheetOutlineSummaryInput;
import dev.erst.gridgrind.contract.dto.SheetPresentationInput;
import dev.erst.gridgrind.contract.dto.SheetProtectionSettings;
import dev.erst.gridgrind.contract.dto.TableColumnInput;
import dev.erst.gridgrind.contract.dto.TableColumnReport;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.TableStyleReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.RecordComponent;
import java.lang.reflect.Type;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

/**
 * Resolves machine-readable field metadata for protocol-catalog record components.
 *
 * <p>This registry maps the full public DTO surface into catalog field-shape groups, so PMD's
 * import-count heuristic is intentionally suppressed here.
 */
@SuppressWarnings("PMD.ExcessiveImports")
public final class CatalogFieldMetadataSupport {
  private static final Set<Class<?>> STRING_FIELD_TYPES =
      Set.of(String.class, java.time.LocalDate.class, java.time.LocalDateTime.class);
  private static final Set<Class<?>> BOOLEAN_FIELD_TYPES = Set.of(boolean.class, Boolean.class);
  private static final Set<Class<?>> NUMERIC_FIELD_TYPES =
      Set.of(
          byte.class,
          short.class,
          int.class,
          long.class,
          float.class,
          double.class,
          Byte.class,
          Short.class,
          Integer.class,
          Long.class,
          Float.class,
          Double.class,
          java.math.BigDecimal.class,
          java.math.BigInteger.class);
  private static final Map<Class<?>, String> NESTED_FIELD_SHAPE_GROUPS =
      Map.ofEntries(
          Map.entry(ExpectedCellValue.class, "expectedCellValueTypes"),
          Map.entry(CellInput.class, "cellInputTypes"),
          Map.entry(HyperlinkTarget.class, "hyperlinkTargetTypes"),
          Map.entry(PaneInput.class, "paneTypes"),
          Map.entry(SheetCopyPosition.class, "sheetCopyPositionTypes"),
          Map.entry(
              dev.erst.gridgrind.contract.selector.WorkbookSelector.class, "workbookSelectorTypes"),
          Map.entry(dev.erst.gridgrind.contract.selector.SheetSelector.class, "sheetSelectorTypes"),
          Map.entry(dev.erst.gridgrind.contract.selector.CellSelector.class, "cellSelectorTypes"),
          Map.entry(dev.erst.gridgrind.contract.selector.RangeSelector.class, "rangeSelectorTypes"),
          Map.entry(
              dev.erst.gridgrind.contract.selector.RowBandSelector.class, "rowBandSelectorTypes"),
          Map.entry(
              dev.erst.gridgrind.contract.selector.ColumnBandSelector.class,
              "columnBandSelectorTypes"),
          Map.entry(dev.erst.gridgrind.contract.selector.TableSelector.class, "tableSelectorTypes"),
          Map.entry(
              dev.erst.gridgrind.contract.selector.TableRowSelector.class, "tableRowSelectorTypes"),
          Map.entry(
              dev.erst.gridgrind.contract.selector.TableCellSelector.class,
              "tableCellSelectorTypes"),
          Map.entry(
              dev.erst.gridgrind.contract.selector.NamedRangeSelector.Ref.class,
              "namedRangeRefSelectorTypes"),
          Map.entry(
              dev.erst.gridgrind.contract.selector.DrawingObjectSelector.class,
              "drawingObjectSelectorTypes"),
          Map.entry(dev.erst.gridgrind.contract.selector.ChartSelector.class, "chartSelectorTypes"),
          Map.entry(
              dev.erst.gridgrind.contract.selector.PivotTableSelector.class,
              "pivotTableSelectorTypes"),
          Map.entry(
              dev.erst.gridgrind.contract.selector.NamedRangeSelector.class,
              "namedRangeSelectorTypes"),
          Map.entry(NamedRangeScope.class, "namedRangeScopeTypes"),
          Map.entry(DrawingAnchorInput.class, "drawingAnchorInputTypes"),
          Map.entry(ChartInput.class, "chartInputTypes"),
          Map.entry(ChartInput.Title.class, "chartTitleInputTypes"),
          Map.entry(ChartInput.Legend.class, "chartLegendInputTypes"),
          Map.entry(PivotTableInput.Source.class, "pivotTableSourceTypes"),
          Map.entry(DataValidationRuleInput.class, "dataValidationRuleTypes"),
          Map.entry(AutofilterFilterCriterionInput.class, "autofilterFilterCriterionTypes"),
          Map.entry(ConditionalFormattingRuleInput.class, "conditionalFormattingRuleTypes"),
          Map.entry(PrintAreaInput.class, "printAreaTypes"),
          Map.entry(PrintScalingInput.class, "printScalingTypes"),
          Map.entry(PrintTitleRowsInput.class, "printTitleRowsTypes"),
          Map.entry(PrintTitleColumnsInput.class, "printTitleColumnsTypes"),
          Map.entry(TableStyleInput.class, "tableStyleTypes"),
          Map.entry(CalculationStrategyInput.class, "calculationStrategyTypes"),
          Map.entry(TextSourceInput.class, "textSourceTypes"),
          Map.entry(BinarySourceInput.class, "binarySourceTypes"),
          Map.entry(FontHeightInput.class, "fontHeightTypes"),
          Map.entry(GridGrindResponse.NamedRangeReport.class, "namedRangeReportTypes"),
          Map.entry(GridGrindResponse.SheetProtectionReport.class, "sheetProtectionReportTypes"),
          Map.entry(TableStyleReport.class, "tableStyleReportTypes"),
          Map.entry(PivotTableReport.class, "pivotTableReportTypes"),
          Map.entry(PivotTableReport.Source.class, "pivotTableReportSourceTypes"),
          Map.entry(ChartReport.class, "chartReportTypes"),
          Map.entry(ChartReport.Title.class, "chartTitleReportTypes"),
          Map.entry(ChartReport.Legend.class, "chartLegendReportTypes"),
          Map.entry(ChartReport.DataSource.class, "chartDataSourceReportTypes"),
          Map.entry(DrawingAnchorReport.class, "drawingAnchorReportTypes"));
  private static final Map<Class<?>, List<String>> NESTED_FIELD_SHAPE_UNIONS =
      Map.of(
          dev.erst.gridgrind.contract.selector.Selector.class,
          List.of(
              "workbookSelectorTypes",
              "sheetSelectorTypes",
              "cellSelectorTypes",
              "rangeSelectorTypes",
              "rowBandSelectorTypes",
              "columnBandSelectorTypes",
              "tableSelectorTypes",
              "tableRowSelectorTypes",
              "tableCellSelectorTypes",
              "namedRangeSelectorTypes",
              "drawingObjectSelectorTypes",
              "chartSelectorTypes",
              "pivotTableSelectorTypes"));
  private static final Map<Class<?>, String> TOP_LEVEL_FIELD_SHAPE_TYPE_SETS =
      Map.ofEntries(
          Map.entry(WorkbookPlan.WorkbookSource.class, "sourceTypes"),
          Map.entry(WorkbookPlan.WorkbookPersistence.class, "persistenceTypes"),
          Map.entry(WorkbookStep.class, "stepTypes"),
          Map.entry(MutationAction.class, "mutationActionTypes"),
          Map.entry(Assertion.class, "assertionTypes"),
          Map.entry(InspectionQuery.class, "inspectionQueryTypes"));
  private static final Map<Class<?>, String> PLAIN_FIELD_SHAPE_GROUPS =
      Map.ofEntries(
          Map.entry(CommentInput.class, "commentInputType"),
          Map.entry(ExecutionPolicyInput.class, "executionPolicyInputType"),
          Map.entry(CalculationPolicyInput.class, "calculationPolicyInputType"),
          Map.entry(ExecutionModeInput.class, "executionModeInputType"),
          Map.entry(ExecutionJournalInput.class, "executionJournalInputType"),
          Map.entry(ExecutionJournal.class, "executionJournalType"),
          Map.entry(ExecutionJournal.SourceSummary.class, "executionJournalSourceSummaryType"),
          Map.entry(
              ExecutionJournal.PersistenceSummary.class, "executionJournalPersistenceSummaryType"),
          Map.entry(ExecutionJournal.Phase.class, "executionJournalPhaseType"),
          Map.entry(ExecutionJournal.Step.class, "executionJournalStepType"),
          Map.entry(ExecutionJournal.Target.class, "executionJournalTargetType"),
          Map.entry(
              ExecutionJournal.FailureClassification.class,
              "executionJournalFailureClassificationType"),
          Map.entry(ExecutionJournal.Calculation.class, "executionJournalCalculationType"),
          Map.entry(ExecutionJournal.Outcome.class, "executionJournalOutcomeType"),
          Map.entry(ExecutionJournal.Event.class, "executionJournalEventType"),
          Map.entry(CalculationReport.class, "calculationReportType"),
          Map.entry(CalculationReport.Preflight.class, "calculationPreflightType"),
          Map.entry(CalculationReport.Summary.class, "calculationSummaryType"),
          Map.entry(CalculationReport.FormulaCapability.class, "formulaCapabilityType"),
          Map.entry(CalculationReport.Execution.class, "calculationExecutionType"),
          Map.entry(RequestWarning.class, "requestWarningType"),
          Map.entry(FormulaEnvironmentInput.class, "formulaEnvironmentInputType"),
          Map.entry(OoxmlOpenSecurityInput.class, "ooxmlOpenSecurityInputType"),
          Map.entry(OoxmlPersistenceSecurityInput.class, "ooxmlPersistenceSecurityInputType"),
          Map.entry(OoxmlEncryptionInput.class, "ooxmlEncryptionInputType"),
          Map.entry(OoxmlSignatureInput.class, "ooxmlSignatureInputType"),
          Map.entry(OoxmlPackageSecurityReport.class, "ooxmlPackageSecurityReportType"),
          Map.entry(OoxmlEncryptionReport.class, "ooxmlEncryptionReportType"),
          Map.entry(OoxmlSignatureReport.class, "ooxmlSignatureReportType"),
          Map.entry(FormulaExternalWorkbookInput.class, "formulaExternalWorkbookInputType"),
          Map.entry(FormulaUdfToolpackInput.class, "formulaUdfToolpackInputType"),
          Map.entry(FormulaUdfFunctionInput.class, "formulaUdfFunctionInputType"),
          Map.entry(
              dev.erst.gridgrind.contract.selector.CellSelector.QualifiedAddress.class,
              "qualifiedCellAddressType"),
          Map.entry(DrawingMarkerInput.class, "drawingMarkerInputType"),
          Map.entry(ChartInput.Series.class, "chartSeriesInputType"),
          Map.entry(ChartInput.DataSource.class, "chartDataSourceInputType"),
          Map.entry(PictureDataInput.class, "pictureDataInputType"),
          Map.entry(PictureInput.class, "pictureInputType"),
          Map.entry(ShapeInput.class, "shapeInputType"),
          Map.entry(EmbeddedObjectInput.class, "embeddedObjectInputType"),
          Map.entry(CommentAnchorInput.class, "commentAnchorInputType"),
          Map.entry(NamedRangeTarget.class, "namedRangeTargetType"),
          Map.entry(SheetProtectionSettings.class, "sheetProtectionSettingsType"),
          Map.entry(CellStyleInput.class, "cellStyleInputType"),
          Map.entry(CellAlignmentInput.class, "cellAlignmentInputType"),
          Map.entry(CellFontInput.class, "cellFontInputType"),
          Map.entry(ColorInput.class, "colorInputType"),
          Map.entry(RichTextRunInput.class, "richTextRunInputType"),
          Map.entry(CellFillInput.class, "cellFillInputType"),
          Map.entry(CellGradientFillInput.class, "cellGradientFillInputType"),
          Map.entry(CellGradientStopInput.class, "cellGradientStopInputType"),
          Map.entry(CellBorderInput.class, "cellBorderInputType"),
          Map.entry(CellBorderSideInput.class, "cellBorderSideInputType"),
          Map.entry(CellProtectionInput.class, "cellProtectionInputType"),
          Map.entry(DataValidationInput.class, "dataValidationInputType"),
          Map.entry(DataValidationPromptInput.class, "dataValidationPromptInputType"),
          Map.entry(DataValidationErrorAlertInput.class, "dataValidationErrorAlertInputType"),
          Map.entry(
              AutofilterFilterCriterionInput.CustomConditionInput.class,
              "autofilterCustomConditionInputType"),
          Map.entry(AutofilterFilterColumnInput.class, "autofilterFilterColumnInputType"),
          Map.entry(AutofilterSortConditionInput.class, "autofilterSortConditionInputType"),
          Map.entry(AutofilterSortStateInput.class, "autofilterSortStateInputType"),
          Map.entry(ConditionalFormattingBlockInput.class, "conditionalFormattingBlockInputType"),
          Map.entry(
              ConditionalFormattingThresholdInput.class, "conditionalFormattingThresholdInputType"),
          Map.entry(HeaderFooterTextInput.class, "headerFooterTextInputType"),
          Map.entry(DifferentialStyleInput.class, "differentialStyleInputType"),
          Map.entry(DifferentialBorderInput.class, "differentialBorderInputType"),
          Map.entry(DifferentialBorderSideInput.class, "differentialBorderSideInputType"),
          Map.entry(IgnoredErrorInput.class, "ignoredErrorInputType"),
          Map.entry(PrintLayoutInput.class, "printLayoutInputType"),
          Map.entry(PrintMarginsInput.class, "printMarginsInputType"),
          Map.entry(PrintSetupInput.class, "printSetupInputType"),
          Map.entry(SheetDefaultsInput.class, "sheetDefaultsInputType"),
          Map.entry(SheetDisplayInput.class, "sheetDisplayInputType"),
          Map.entry(SheetOutlineSummaryInput.class, "sheetOutlineSummaryInputType"),
          Map.entry(SheetPresentationInput.class, "sheetPresentationInputType"),
          Map.entry(PivotTableInput.class, "pivotTableInputType"),
          Map.entry(PivotTableInput.Anchor.class, "pivotTableAnchorInputType"),
          Map.entry(PivotTableInput.DataField.class, "pivotTableDataFieldInputType"),
          Map.entry(TableColumnInput.class, "tableColumnInputType"),
          Map.entry(TableInput.class, "tableInputType"),
          Map.entry(WorkbookProtectionInput.class, "workbookProtectionInputType"),
          Map.entry(WorkbookProtectionReport.class, "workbookProtectionReportType"),
          Map.entry(GridGrindResponse.SheetSummaryReport.class, "sheetSummaryReportType"),
          Map.entry(GridGrindResponse.CellStyleReport.class, "cellStyleReportType"),
          Map.entry(CellAlignmentReport.class, "cellAlignmentReportType"),
          Map.entry(CellFontReport.class, "cellFontReportType"),
          Map.entry(CellFillReport.class, "cellFillReportType"),
          Map.entry(CellBorderReport.class, "cellBorderReportType"),
          Map.entry(CellBorderSideReport.class, "cellBorderSideReportType"),
          Map.entry(CellProtectionReport.class, "cellProtectionReportType"),
          Map.entry(CellColorReport.class, "cellColorReportType"),
          Map.entry(FontHeightReport.class, "fontHeightReportType"),
          Map.entry(CellGradientFillReport.class, "cellGradientFillReportType"),
          Map.entry(CellGradientStopReport.class, "cellGradientStopReportType"),
          Map.entry(TableEntryReport.class, "tableEntryReportType"),
          Map.entry(TableColumnReport.class, "tableColumnReportType"),
          Map.entry(DrawingMarkerReport.class, "drawingMarkerReportType"),
          Map.entry(PivotTableReport.Anchor.class, "pivotTableAnchorReportType"),
          Map.entry(PivotTableReport.Field.class, "pivotTableFieldReportType"),
          Map.entry(PivotTableReport.DataField.class, "pivotTableDataFieldReportType"),
          Map.entry(ChartReport.Axis.class, "chartAxisReportType"),
          Map.entry(ChartReport.Series.class, "chartSeriesReportType"));

  private CatalogFieldMetadataSupport() {}

  /** Returns the catalog field entry derived from one reflected record component. */
  public static FieldEntry fieldEntry(RecordComponent component, Set<String> optionalFields) {
    Objects.requireNonNull(component, "component must not be null");
    Objects.requireNonNull(optionalFields, "optionalFields must not be null");
    return new FieldEntry(
        component.getName(),
        optionalFields.contains(component.getName())
            ? FieldRequirement.OPTIONAL
            : FieldRequirement.REQUIRED,
        fieldShape(component.getGenericType()),
        enumValues(component.getGenericType()));
  }

  /** Returns the machine-readable field shape for one record component type. */
  public static FieldShape fieldShape(Type type) {
    Objects.requireNonNull(type, "type must not be null");
    if (type instanceof ParameterizedType parameterizedType) {
      return fieldShape(parameterizedType);
    }
    if (type instanceof Class<?> classType) {
      return fieldShape(classType);
    }
    throw new IllegalStateException("Unsupported catalog field type: " + type);
  }

  /** Returns the machine-readable field shape for one parameterized record component type. */
  public static FieldShape fieldShape(ParameterizedType parameterizedType) {
    Objects.requireNonNull(parameterizedType, "parameterizedType must not be null");
    Type rawType = parameterizedType.getRawType();
    if (rawType == java.util.List.class) {
      Type[] typeArguments = parameterizedType.getActualTypeArguments();
      if (typeArguments.length != 1) {
        throw new IllegalStateException(
            "List field must declare exactly one type argument: " + parameterizedType);
      }
      return new FieldShape.ListShape(fieldShape(typeArguments[0]));
    }
    throw new IllegalStateException(
        "Unsupported parameterized catalog field type: " + parameterizedType);
  }

  /** Returns the machine-readable field shape for one non-parameterized record component type. */
  public static FieldShape fieldShape(Class<?> classType) {
    Objects.requireNonNull(classType, "classType must not be null");
    if (STRING_FIELD_TYPES.contains(classType)) {
      return new FieldShape.Scalar(ScalarType.STRING);
    }
    if (BOOLEAN_FIELD_TYPES.contains(classType)) {
      return new FieldShape.Scalar(ScalarType.BOOLEAN);
    }
    if (isNumericType(classType)) {
      return new FieldShape.Scalar(ScalarType.NUMBER);
    }
    if (classType.isEnum()) {
      return new FieldShape.Scalar(ScalarType.STRING);
    }
    List<String> nestedGroupUnion = NESTED_FIELD_SHAPE_UNIONS.get(classType);
    if (nestedGroupUnion == null) {
      nestedGroupUnion = lookupAssignableGroupList(NESTED_FIELD_SHAPE_UNIONS, classType);
    }
    if (nestedGroupUnion != null) {
      return new FieldShape.NestedTypeGroupUnionRef(nestedGroupUnion);
    }
    String topLevelTypeSet = TOP_LEVEL_FIELD_SHAPE_TYPE_SETS.get(classType);
    if (topLevelTypeSet == null) {
      topLevelTypeSet = lookupAssignableGroup(TOP_LEVEL_FIELD_SHAPE_TYPE_SETS, classType);
    }
    if (topLevelTypeSet != null) {
      return new FieldShape.TopLevelTypeSetRef(topLevelTypeSet);
    }
    String nestedGroup = NESTED_FIELD_SHAPE_GROUPS.get(classType);
    if (nestedGroup == null) {
      nestedGroup = lookupAssignableGroup(NESTED_FIELD_SHAPE_GROUPS, classType);
    }
    if (nestedGroup != null) {
      return new FieldShape.NestedTypeGroupRef(nestedGroup);
    }
    String plainGroup = PLAIN_FIELD_SHAPE_GROUPS.get(classType);
    if (plainGroup == null) {
      plainGroup = lookupAssignableGroup(PLAIN_FIELD_SHAPE_GROUPS, classType);
    }
    if (plainGroup != null) {
      return new FieldShape.PlainTypeGroupRef(plainGroup);
    }
    throw new IllegalStateException("Unsupported catalog field type: " + classType.getName());
  }

  static String lookupAssignableGroup(Map<Class<?>, String> groups, Class<?> classType) {
    String resolvedGroup = null;
    Class<?> resolvedBase = null;
    for (Map.Entry<Class<?>, String> entry : groups.entrySet()) {
      Class<?> candidateBase = entry.getKey();
      if (!candidateBase.isAssignableFrom(classType) || candidateBase.equals(classType)) {
        continue;
      }
      if (resolvedBase == null || resolvedBase.isAssignableFrom(candidateBase)) {
        resolvedBase = candidateBase;
        resolvedGroup = entry.getValue();
        continue;
      }
      if (!candidateBase.isAssignableFrom(resolvedBase)) {
        throw new IllegalStateException(
            "Ambiguous catalog field group mapping for "
                + classType.getName()
                + ": "
                + resolvedBase.getName()
                + " and "
                + candidateBase.getName());
      }
    }
    return resolvedGroup;
  }

  static List<String> lookupAssignableGroupList(
      Map<Class<?>, List<String>> groups, Class<?> classType) {
    List<String> resolvedGroup = null;
    Class<?> resolvedBase = null;
    for (Map.Entry<Class<?>, List<String>> entry : groups.entrySet()) {
      Class<?> candidateBase = entry.getKey();
      if (!candidateBase.isAssignableFrom(classType) || candidateBase.equals(classType)) {
        continue;
      }
      if (resolvedBase == null || resolvedBase.isAssignableFrom(candidateBase)) {
        resolvedBase = candidateBase;
        resolvedGroup = entry.getValue();
      }
    }
    return resolvedGroup;
  }

  /** Returns whether one non-parameterized record component type is represented as JSON NUMBER. */
  public static boolean isNumericType(Class<?> classType) {
    Objects.requireNonNull(classType, "classType must not be null");
    return NUMERIC_FIELD_TYPES.contains(classType);
  }

  /** Returns the set of all sealed types registered in the nested field-shape group map. */
  public static Set<Class<?>> registeredNestedTypes() {
    return NESTED_FIELD_SHAPE_GROUPS.keySet();
  }

  /** Returns the set of all record types registered in the plain field-shape group map. */
  public static Set<Class<?>> registeredPlainTypes() {
    return PLAIN_FIELD_SHAPE_GROUPS.keySet();
  }

  /** Validates that one nested sealed input type maps to the published field-shape group. */
  public static void validateNestedTypeGroupMapping(Class<?> sealedType, String expectedGroup) {
    Objects.requireNonNull(sealedType, "sealedType must not be null");
    String mappedGroup = NESTED_FIELD_SHAPE_GROUPS.get(sealedType);
    if (!expectedGroup.equals(mappedGroup)) {
      throw new IllegalStateException(
          "Field-shape nested group mapping mismatch for "
              + sealedType.getName()
              + ": expected="
              + expectedGroup
              + ", mapped="
              + mappedGroup);
    }
  }

  /** Validates that one plain record input type maps to the published field-shape group. */
  public static void validatePlainTypeGroupMapping(Class<?> recordType, String expectedGroup) {
    Objects.requireNonNull(recordType, "recordType must not be null");
    String mappedGroup = PLAIN_FIELD_SHAPE_GROUPS.get(recordType);
    if (!expectedGroup.equals(mappedGroup)) {
      throw new IllegalStateException(
          "Field-shape plain group mapping mismatch for "
              + recordType.getName()
              + ": expected="
              + expectedGroup
              + ", mapped="
              + mappedGroup);
    }
  }

  private static java.util.List<String> enumValues(Type type) {
    if (type instanceof Class<?> classType && classType.isEnum()) {
      return Arrays.stream(classType.getEnumConstants())
          .map(value -> ((Enum<?>) value).name())
          .toList();
    }
    return java.util.List.of();
  }
}
