package dev.erst.gridgrind.protocol.parity;

import dev.erst.gridgrind.protocol.catalog.Catalog;
import dev.erst.gridgrind.protocol.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.protocol.catalog.TypeEntry;
import dev.erst.gridgrind.protocol.dto.GridGrindRequest;
import dev.erst.gridgrind.protocol.dto.GridGrindResponse;
import dev.erst.gridgrind.protocol.exec.DefaultGridGrindRequestExecutor;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadResult;
import java.nio.file.Path;
import java.util.List;

/** Small parity-test helper for executing GridGrind requests and reading the protocol catalog. */
final class XlsxParityGridGrind {
  private static final DefaultGridGrindRequestExecutor EXECUTOR =
      new DefaultGridGrindRequestExecutor();

  private XlsxParityGridGrind() {}

  static Catalog catalog() {
    return GridGrindProtocolCatalog.catalog();
  }

  static boolean hasSourceType(String id) {
    return hasType(catalog().sourceTypes(), id);
  }

  static boolean hasPersistenceType(String id) {
    return hasType(catalog().persistenceTypes(), id);
  }

  static boolean hasOperationType(String id) {
    return hasType(catalog().operationTypes(), id);
  }

  static boolean hasReadType(String id) {
    return hasType(catalog().readTypes(), id);
  }

  static GridGrindResponse executeReadWorkbook(Path workbookPath, WorkbookReadOperation... reads) {
    return EXECUTOR.execute(
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.ExistingFile(workbookPath.toString()),
            new GridGrindRequest.WorkbookPersistence.None(),
            List.of(),
            List.of(reads)));
  }

  static GridGrindResponse.Success readWorkbook(Path workbookPath, WorkbookReadOperation... reads) {
    return success(executeReadWorkbook(workbookPath, reads));
  }

  static GridGrindResponse.Success mutateWorkbook(
      Path workbookPath,
      Path saveAsPath,
      List<WorkbookOperation> operations,
      WorkbookReadOperation... reads) {
    return success(
        EXECUTOR.execute(
            new GridGrindRequest(
                new GridGrindRequest.WorkbookSource.ExistingFile(workbookPath.toString()),
                new GridGrindRequest.WorkbookPersistence.SaveAs(saveAsPath.toString()),
                List.copyOf(operations),
                List.of(reads))));
  }

  static GridGrindResponse.Success overwriteWorkbook(
      Path workbookPath, List<WorkbookOperation> operations, WorkbookReadOperation... reads) {
    return success(
        EXECUTOR.execute(
            new GridGrindRequest(
                new GridGrindRequest.WorkbookSource.ExistingFile(workbookPath.toString()),
                new GridGrindRequest.WorkbookPersistence.OverwriteSource(),
                List.copyOf(operations),
                List.of(reads))));
  }

  static GridGrindResponse.Failure mutateWorkbookExpectingFailure(
      Path workbookPath, List<WorkbookOperation> operations, WorkbookReadOperation... reads) {
    GridGrindResponse response =
        EXECUTOR.execute(
            new GridGrindRequest(
                new GridGrindRequest.WorkbookSource.ExistingFile(workbookPath.toString()),
                new GridGrindRequest.WorkbookPersistence.None(),
                List.copyOf(operations),
                List.of(reads)));
    if (response instanceof GridGrindResponse.Failure failure) {
      return failure;
    }
    throw new AssertionError("Expected GridGrind failure but request succeeded: " + response);
  }

  static String savedPath(GridGrindResponse.Success success) {
    return switch (success.persistence()) {
      case GridGrindResponse.PersistenceOutcome.SavedAs savedAs -> savedAs.executionPath();
      case GridGrindResponse.PersistenceOutcome.Overwritten overwritten ->
          overwritten.executionPath();
      case GridGrindResponse.PersistenceOutcome.NotSaved _ ->
          throw new AssertionError("Expected the workbook to be persisted");
    };
  }

  static <T extends WorkbookReadResult> T read(
      GridGrindResponse.Success success, String requestId, Class<T> type) {
    return type.cast(
        success.reads().stream()
            .filter(result -> result.requestId().equals(requestId))
            .findFirst()
            .filter(type::isInstance)
            .orElseThrow(
                () -> new AssertionError("Missing read result " + requestId + " as " + type)));
  }

  private static GridGrindResponse.Success success(GridGrindResponse response) {
    if (response instanceof GridGrindResponse.Success success) {
      return success;
    }
    throw new AssertionError("Expected GridGrind success but got " + response);
  }

  private static boolean hasType(List<TypeEntry> entries, String id) {
    return entries.stream().map(TypeEntry::id).anyMatch(id::equals);
  }
}
