package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;

/** Shared top-level example descriptors for the shipped example registry. */
final class ExampleDefinitions {
  private ExampleDefinitions() {}

  static GridGrindShippedExamples.ShippedExample example(
      String id, String fileName, String summary, WorkbookPlan plan) {
    return new GridGrindShippedExamples.ShippedExample(id, fileName, summary, plan);
  }
}
