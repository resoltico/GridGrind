package dev.erst.gridgrind.core.plan;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.vocabulary.StepId;

/** Ordered result payload returned for one requested inspection result. */
@Union(discriminator = "type")
public sealed interface InspectionResult
    permits InspectionIntrospectionResult, InspectionSurfaceResult, InspectionAnalysisResult {

  /** Stable caller-provided identifier copied from the matching read operation. */
  StepId stepId();
}
