package dev.erst.gridgrind.contract.action;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.GridGrindRequestProblemSupport;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.json.FieldValidationNamingRule;
import dev.erst.gridgrind.contract.json.FieldValidationProblem;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import java.util.Arrays;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Edge-path coverage for mutation-action selector metadata lookup. */
class MutationActionCoverageTest {
  @Test
  void discoversTargetsFromTheCanonicalOperationContract() {
    CellMutationAction.SetCell action =
        new CellMutationAction.SetCell(new CellInput.NumberValue(1.0));

    assertEquals(
        List.of(CellSelector.ByAddress.class, TableCellSelector.ByColumnName.class),
        List.of(MutationAction.allowedTargetTypes(action)));
    assertEquals(
        List.of(CellSelector.ByAddress.class, TableCellSelector.ByColumnName.class),
        List.of(MutationAction.allowedTargetTypesForType(CellMutationAction.SetCell.class)));
  }

  @Test
  void rejectsUnmappedAndMetadataFreeMutationActionTypes() {
    @SuppressWarnings("unchecked")
    Class<? extends MutationAction> nonRecordActionType =
        (Class<? extends MutationAction>) MutationAction.class;

    IllegalArgumentException nonRecordFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> MutationAction.allowedTargetTypesForType(nonRecordActionType));
    assertEquals(
        "No target-type mapping configured for action class dev.erst.gridgrind.contract.action.MutationAction",
        nonRecordFailure.getMessage());

    @SuppressWarnings("unchecked")
    Class<? extends MutationAction> missingMetadataActionType =
        (Class<? extends MutationAction>) (Class<?>) MissingMetadataActionRecord.class;

    IllegalArgumentException missingMetadataFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> MutationAction.allowedTargetTypesForType(missingMetadataActionType));
    assertEquals(
        "No target-type mapping configured for action class "
            + MissingMetadataActionRecord.class.getName(),
        missingMetadataFailure.getMessage());
  }

  @Test
  void rectangularRowValidationRejectsEmptyRowsNullRowsAndNullCells() {
    MutationAction.Validation.requireRectangularRows(
        List.of(List.of(new CellInput.Blank()), List.of(new CellInput.Blank())));

    assertEquals(
        "rows must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> MutationAction.Validation.requireRectangularRows(List.of()))
            .getMessage());
    assertEquals(
        "rows must not contain null rows",
        assertThrows(
                NullPointerException.class,
                () ->
                    MutationAction.Validation.requireRectangularRows(
                        Arrays.asList((List<CellInput>) null)))
            .getMessage());
    assertEquals(
        "rows must not contain null cell values",
        assertThrows(
                NullPointerException.class,
                () ->
                    MutationAction.Validation.requireRectangularRows(
                        List.of(Arrays.asList(new CellInput.Blank(), null))))
            .getMessage());
    assertEquals(
        "rows must not contain empty rows",
        assertThrows(
                IllegalArgumentException.class,
                () -> MutationAction.Validation.requireRectangularRows(List.of(List.of())))
            .getMessage());
    assertEquals(
        "rows must describe a rectangular matrix",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    MutationAction.Validation.requireRectangularRows(
                        List.of(
                            List.of(new CellInput.Blank()),
                            List.of(new CellInput.Blank(), new CellInput.Blank()))))
            .getMessage());
  }

  @Test
  void typedNameValidationCarriesCauseSpecificResolution() {
    var failure =
        MutationActionNameValidation.invalidField(
            FieldValidationProblem.atField("name", FieldValidationNamingRule.DEFINED_NAME_SYNTAX));
    FieldValidationProblem requestProblem =
        assertInstanceOf(FieldValidationProblem.class, failure.requestProblem());

    assertEquals(
        "name must start with a letter or underscore and contain only letters, digits, underscore, or period",
        failure.getMessage());
    assertEquals(FieldValidationNamingRule.DEFINED_NAME_SYNTAX, requestProblem.rule());
    assertEquals(
        "Provide a valid Excel defined name for field 'name'.",
        GridGrindRequestProblemSupport.resolution(
            requestProblem,
            new ProblemContext.ReadRequest(
                ProblemContextRequestSurfaces.RequestInput.standardInput(),
                ProblemContextRequestSurfaces.JsonLocation.unavailable())));
  }

  private record MissingMetadataActionRecord() {}
}
