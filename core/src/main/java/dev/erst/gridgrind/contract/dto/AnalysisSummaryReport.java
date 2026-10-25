package dev.erst.gridgrind.contract.dto;

/** Summary counts for one finding-bearing analysis result. */
public record AnalysisSummaryReport(
    int totalCount, int errorCount, int warningCount, int infoCount) {
  public AnalysisSummaryReport {
    if (totalCount < 0) {
      throw new IllegalArgumentException("totalCount must not be negative");
    }
    if (errorCount < 0) {
      throw new IllegalArgumentException("errorCount must not be negative");
    }
    if (warningCount < 0) {
      throw new IllegalArgumentException("warningCount must not be negative");
    }
    if (infoCount < 0) {
      throw new IllegalArgumentException("infoCount must not be negative");
    }
    if (totalCount != errorCount + warningCount + infoCount) {
      throw new IllegalArgumentException(
          "totalCount must equal errorCount + warningCount + infoCount");
    }
  }
}
