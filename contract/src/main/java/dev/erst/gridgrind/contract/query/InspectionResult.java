package dev.erst.gridgrind.contract.query;

import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** Ordered result payload returned for one requested inspection result. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
public sealed interface InspectionResult
    permits InspectionIntrospectionResult, InspectionSurfaceResult, InspectionAnalysisResult {

  /** Stable caller-provided identifier copied from the matching read operation. */
  String stepId();
}
