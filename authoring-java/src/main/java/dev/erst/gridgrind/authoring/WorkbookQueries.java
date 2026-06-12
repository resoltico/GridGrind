package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;

/** Canonical workbook-query factories kept internal to the Java authoring surface. */
final class WorkbookQueries {
  private WorkbookQueries() {}

  static WorkbookIntrospectionQuery.GetWorkbookSummary workbookSummary() {
    return new WorkbookIntrospectionQuery.GetWorkbookSummary();
  }

  static WorkbookIntrospectionQuery.GetPackageSecurity packageSecurity() {
    return new WorkbookIntrospectionQuery.GetPackageSecurity();
  }

  static WorkbookIntrospectionQuery.GetWorkbookProtection workbookProtection() {
    return new WorkbookIntrospectionQuery.GetWorkbookProtection();
  }

  static WorkbookIntrospectionQuery.GetNamedRanges namedRanges() {
    return new WorkbookIntrospectionQuery.GetNamedRanges();
  }
}
