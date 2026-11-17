package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.value.CellGridSpec;
import dev.erst.gridgrind.core.model.value.CellRowSpec;
import dev.erst.gridgrind.core.model.value.ExcelDateSystem;
import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;
import dev.erst.gridgrind.core.model.value.TemporalSemantics;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.PreparedMutationStep;
import dev.erst.gridgrind.core.plan.ResolvedMutationAction;
import dev.erst.gridgrind.core.plan.ResolvedPlan;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.InputReference;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.core.problem.ValueInvariantException;
import java.util.List;
import java.util.Optional;

/** Validates intrinsic authored temporal operands after actual source calendar discovery. */
final class TemporalInputValidation {
  private TemporalInputValidation() {}

  static void validate(
      ResolvedPlan plan,
      ExcelDateSystem system,
      Optional<RequestIntake> intake,
      List<GridGrindProblemDetail.Problem> problems) {
    for (int step = 0; step < plan.steps().size(); step++) {
      if (!(plan.steps().get(step) instanceof PreparedMutationStep mutation)) continue;
      String path = "steps[" + step + "].action";
      switch (mutation.action()) {
        case ResolvedMutationAction.SetCell cell ->
            check(cell.value(), system, path + ".value.value", plan.request(), intake, problems);
        case ResolvedMutationAction.SetRange range -> {
          boolean typed =
              ((CellMutationAction.SetRange) mutation.authoredStep().action()).rows()
                  instanceof CellGridSpec.Typed;
          for (int row = 0; row < range.rows().size(); row++)
            for (int column = 0; column < range.rows().get(row).size(); column++)
              check(
                  range.rows().get(row).get(column),
                  system,
                  path + ".rows.cells[" + row + "][" + column + "]" + (typed ? ".value" : ""),
                  plan.request(),
                  intake,
                  problems);
        }
        case ResolvedMutationAction.AppendRow row -> {
          boolean typed =
              ((CellMutationAction.AppendRow) mutation.authoredStep().action()).values()
                  instanceof CellRowSpec.Typed;
          for (int column = 0; column < row.values().size(); column++)
            check(
                row.values().get(column),
                system,
                path + ".values.cells[" + column + "]" + (typed ? ".value" : ""),
                plan.request(),
                intake,
                problems);
        }
        default -> {}
      }
    }
  }

  private static void check(
      ResolvedCellWrite value,
      ExcelDateSystem system,
      String path,
      WorkbookPlan request,
      Optional<RequestIntake> intake,
      List<GridGrindProblemDetail.Problem> problems) {
    try {
      TemporalSemantics.requireWritable(value, system);
    } catch (ValueInvariantException invalid) {
      JsonLocation location =
          intake
              .map(analysis -> analysis.jsonLocationAt(path))
              .orElseGet(() -> JsonLocation.pathOnly(path));
      var failure =
          new InputResolutionFailure(
              invalid, Optional.of(location), InputReference.kind("authored temporal value"));
      problems.add(
          GridGrindProblems.fromException(
              invalid, RequestPreflight.resolveInputsContext(request, failure)));
    }
  }
}
