package dev.erst.gridgrind.executor.parity;

import static dev.erst.gridgrind.executor.parity.ParityPlanSupport.inspect;
import static dev.erst.gridgrind.executor.parity.ParityPlanSupport.mutate;
import static dev.erst.gridgrind.executor.parity.XlsxParityProbeRegistry.*;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.executor.parity.ParityPlanSupport.PendingMutation;
import dev.erst.gridgrind.executor.parity.XlsxParityProbeRegistry.ProbeContext;
import dev.erst.gridgrind.executor.parity.XlsxParityProbeRegistry.ProbeResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

/** Security and platform-gap parity probes. */
final class XlsxParitySecurityProbeGroup {
  private XlsxParitySecurityProbeGroup() {}

  static ProbeResult probeEventModelGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario largeSheet =
        context.scenario(XlsxParityScenarios.LARGE_SHEET);
    int poiRows = XlsxParityOracle.eventModelRowCount(largeSheet.workbookPath());
    GridGrindResponse.Success full =
        XlsxParityGridGrind.readWorkbook(
            largeSheet.workbookPath(),
            inspect(
                "workbook",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetWorkbookSummary()),
            inspect(
                "sheet", new SheetSelector.ByName("Large"), new InspectionQuery.GetSheetSummary()));
    GridGrindResponse response =
        XlsxParityGridGrind.executeReadWorkbook(
            largeSheet.workbookPath(),
            new ExecutionModeInput(
                ExecutionModeInput.ReadMode.EVENT_READ, ExecutionModeInput.WriteMode.FULL_XSSF),
            inspect(
                "workbook",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetWorkbookSummary()),
            inspect(
                "sheet", new SheetSelector.ByName("Large"), new InspectionQuery.GetSheetSummary()));
    if (response instanceof GridGrindResponse.Failure failure) {
      return fail("GridGrind event-read request failed: " + failure.problem().message());
    }
    GridGrindResponse.Success event = (GridGrindResponse.Success) response;
    InspectionResult.SheetSummaryResult sheetSummary =
        XlsxParityGridGrind.read(event, "sheet", InspectionResult.SheetSummaryResult.class);
    if (sheetSummary.sheet().physicalRowCount() != poiRows
        || sheetSummary.sheet().lastColumnIndex() != 19) {
      return fail(
          "GridGrind event-read sheet summary diverged from the large-sheet oracle: rows="
              + sheetSummary.sheet().physicalRowCount()
              + ", lastColumnIndex="
              + sheetSummary.sheet().lastColumnIndex());
    }
    return full.inspections().equals(event.inspections())
        ? pass("Event-model parity is present and matches full-XSSF summary reads.")
        : fail("GridGrind event-read summaries diverged from the full-XSSF summaries.");
  }

  static ProbeResult probeSxssfGap(ProbeContext context) {
    Path streamedWorkbook = XlsxParityOracle.sxssfWriteWorkbook(context.derivedDirectory("sxssf"));
    boolean poiSucceeded =
        XlsxParitySupport.call(
            "inspect streamed parity workbook size",
            () -> Files.exists(streamedWorkbook) && Files.size(streamedWorkbook) > 0);
    Path gridGrindWorkbook = context.derivedWorkbook("streaming-gridgrind");
    List<PendingMutation> operations = new ArrayList<>();
    operations.add(mutate(new SheetSelector.ByName("Streamed"), new MutationAction.EnsureSheet()));
    for (int rowIndex = 0; rowIndex < 1_500; rowIndex++) {
      operations.add(
          mutate(
              new SheetSelector.ByName("Streamed"),
              new MutationAction.AppendRow(
                  List.of(text("R" + rowIndex), new CellInput.Numeric((double) rowIndex)))));
    }
    GridGrindResponse.Success streamed =
        XlsxParityGridGrind.writeNewWorkbook(
            gridGrindWorkbook,
            new ExecutionModeInput(
                ExecutionModeInput.ReadMode.FULL_XSSF,
                ExecutionModeInput.WriteMode.STREAMING_WRITE),
            operations,
            inspect(
                "workbook",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetWorkbookSummary()),
            inspect(
                "sheet",
                new SheetSelector.ByName("Streamed"),
                new InspectionQuery.GetSheetSummary()));
    GridGrindResponse.Success reopened =
        XlsxParityGridGrind.readWorkbook(
            gridGrindWorkbook,
            inspect(
                "workbook",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetWorkbookSummary()),
            inspect(
                "sheet",
                new SheetSelector.ByName("Streamed"),
                new InspectionQuery.GetSheetSummary()));
    InspectionResult.SheetSummaryResult sheetSummary =
        XlsxParityGridGrind.read(streamed, "sheet", InspectionResult.SheetSummaryResult.class);
    return poiSucceeded
            && Files.exists(gridGrindWorkbook)
            && sheetSummary.sheet().physicalRowCount() == 1_500
            && sheetSummary.sheet().lastColumnIndex() == 1
            && streamed.inspections().equals(reopened.inspections())
        ? pass("SXSSF streaming-write parity is present.")
        : fail("GridGrind streaming-write parity did not match the expected low-memory workbook.");
  }

  static ProbeResult probeEncryptionGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario encrypted =
        context.scenario(XlsxParityScenarios.ENCRYPTED_WORKBOOK);
    OoxmlOpenSecurityInput sourceSecurity = encryptedOpenSecurity();
    boolean poiSourceOpens = XlsxParityOracle.encryptedWorkbookOpens(encrypted.workbookPath());

    GridGrindResponse.Success sourceRead =
        XlsxParityGridGrind.readWorkbook(
            encrypted.workbookPath(),
            sourceSecurity,
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Encrypted", List.of("A1")),
                new InspectionQuery.GetCells()));
    InspectionResult.PackageSecurityResult sourceSecurityResult =
        XlsxParityGridGrind.read(
            sourceRead, "security", InspectionResult.PackageSecurityResult.class);
    InspectionResult.CellsResult sourceCells =
        XlsxParityGridGrind.read(sourceRead, "cells", InspectionResult.CellsResult.class);

    Path preservedOutput = context.derivedWorkbook("encrypted-preserved");
    GridGrindResponse.Success preservedSave =
        XlsxParityGridGrind.mutateWorkbook(
            encrypted.workbookPath(),
            sourceSecurity,
            preservedOutput,
            null,
            null,
            List.of(
                mutate(
                    new CellSelector.ByAddress("Encrypted", "A2"),
                    new MutationAction.SetCell(text("Preserved encryption")))));
    boolean poiPreservedOpens = XlsxParityOracle.encryptedWorkbookOpens(preservedOutput);
    String poiPreservedText =
        XlsxParityOracle.encryptedStringCell(preservedOutput, "Encrypted", "A2");
    GridGrindResponse.Success preservedRead =
        XlsxParityGridGrind.readWorkbook(
            preservedOutput,
            sourceSecurity,
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Encrypted", List.of("A1", "A2")),
                new InspectionQuery.GetCells()));
    InspectionResult.PackageSecurityResult preservedSecurityResult =
        XlsxParityGridGrind.read(
            preservedRead, "security", InspectionResult.PackageSecurityResult.class);
    InspectionResult.CellsResult preservedCells =
        XlsxParityGridGrind.read(preservedRead, "cells", InspectionResult.CellsResult.class);

    Path authoredOutput = context.derivedWorkbook("encrypted-authored");
    GridGrindResponse.Success authoredSave =
        XlsxParityGridGrind.writeNewWorkbook(
            authoredOutput,
            new OoxmlPersistenceSecurityInput(
                new OoxmlEncryptionInput(XlsxParityScenarios.ENCRYPTION_PASSWORD, null), null),
            (ExecutionModeInput) null,
            List.of(
                mutate(new SheetSelector.ByName("Secure"), new MutationAction.EnsureSheet()),
                mutate(
                    new CellSelector.ByAddress("Secure", "A1"),
                    new MutationAction.SetCell(text("Authored encrypted")))));
    String poiAuthoredText = XlsxParityOracle.encryptedStringCell(authoredOutput, "Secure", "A1");
    GridGrindResponse.Success authoredRead =
        XlsxParityGridGrind.readWorkbook(
            authoredOutput,
            sourceSecurity,
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Secure", List.of("A1")),
                new InspectionQuery.GetCells()));
    InspectionResult.PackageSecurityResult authoredSecurityResult =
        XlsxParityGridGrind.read(
            authoredRead, "security", InspectionResult.PackageSecurityResult.class);
    InspectionResult.CellsResult authoredCells =
        XlsxParityGridGrind.read(authoredRead, "cells", InspectionResult.CellsResult.class);

    boolean preservedSavedToExpectedPath =
        XlsxParityGridGrind.savedPath(preservedSave)
            .equals(preservedOutput.toAbsolutePath().toString());
    boolean authoredSavedToExpectedPath =
        XlsxParityGridGrind.savedPath(authoredSave)
            .equals(authoredOutput.toAbsolutePath().toString());

    return poiSourceOpens
            && hasEncryptedAgilePackage(sourceSecurityResult.security())
            && "Encrypted workbook".equals(textCell(sourceCells, "A1"))
            && preservedSavedToExpectedPath
            && poiPreservedOpens
            && "Preserved encryption".equals(poiPreservedText)
            && hasEncryptedAgilePackage(preservedSecurityResult.security())
            && "Encrypted workbook".equals(textCell(preservedCells, "A1"))
            && "Preserved encryption".equals(textCell(preservedCells, "A2"))
            && authoredSavedToExpectedPath
            && "Authored encrypted".equals(poiAuthoredText)
            && hasEncryptedAgilePackage(authoredSecurityResult.security())
            && "Authored encrypted".equals(textCell(authoredCells, "A1"))
        ? pass(
            "OOXML encryption parity is present for encrypted open, preserved encrypted save-as, and authored encrypted save-as.")
        : fail(
            "OOXML encryption parity mismatch."
                + " poiSourceOpens="
                + poiSourceOpens
                + " sourceEncrypted="
                + hasEncryptedAgilePackage(sourceSecurityResult.security())
                + " preservedSaved="
                + preservedSavedToExpectedPath
                + " poiPreservedOpens="
                + poiPreservedOpens
                + " poiPreservedText="
                + poiPreservedText
                + " preservedEncrypted="
                + hasEncryptedAgilePackage(preservedSecurityResult.security())
                + " authoredSaved="
                + authoredSavedToExpectedPath
                + " poiAuthoredText="
                + poiAuthoredText
                + " authoredEncrypted="
                + hasEncryptedAgilePackage(authoredSecurityResult.security()));
  }

  static ProbeResult probeEncryptionPasswordErrors(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario encrypted =
        context.scenario(XlsxParityScenarios.ENCRYPTED_WORKBOOK);

    GridGrindResponse missingPasswordResponse =
        XlsxParityGridGrind.executeReadWorkbook(
            encrypted.workbookPath(),
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()));
    GridGrindResponse wrongPasswordResponse =
        XlsxParityGridGrind.executeReadWorkbook(
            encrypted.workbookPath(),
            new OoxmlOpenSecurityInput(java.util.Optional.of("gridgrind-phase9-wrong-password")),
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()));

    if (!(missingPasswordResponse instanceof GridGrindResponse.Failure missingPassword)) {
      return fail("Encrypted workbook open without a password unexpectedly succeeded.");
    }
    if (!(wrongPasswordResponse instanceof GridGrindResponse.Failure wrongPassword)) {
      return fail("Encrypted workbook open with the wrong password unexpectedly succeeded.");
    }

    return missingPassword.problem().code() == GridGrindProblemCode.WORKBOOK_PASSWORD_REQUIRED
            && missingPassword.problem().category() == GridGrindProblemCategory.SECURITY
            && wrongPassword.problem().code() == GridGrindProblemCode.INVALID_WORKBOOK_PASSWORD
            && wrongPassword.problem().category() == GridGrindProblemCategory.SECURITY
        ? pass(
            "Encrypted workbook password failures map to stable security problem codes for missing and invalid passwords.")
        : fail(
            "Encrypted workbook password failure parity mismatch."
                + " missingCode="
                + missingPassword.problem().code()
                + " missingCategory="
                + missingPassword.problem().category()
                + " wrongCode="
                + wrongPassword.problem().code()
                + " wrongCategory="
                + wrongPassword.problem().category());
  }

  static ProbeResult probeSigningGap(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario signed =
        context.scenario(XlsxParityScenarios.SIGNED_WORKBOOK);
    boolean poiSourceValid = XlsxParityOracle.signatureValid(signed.workbookPath());
    GridGrindResponse.Success sourceRead =
        XlsxParityGridGrind.readWorkbook(
            signed.workbookPath(),
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Signed", List.of("A1")),
                new InspectionQuery.GetCells()));
    InspectionResult.PackageSecurityResult sourceSecurityResult =
        XlsxParityGridGrind.read(
            sourceRead, "security", InspectionResult.PackageSecurityResult.class);
    InspectionResult.CellsResult sourceCells =
        XlsxParityGridGrind.read(sourceRead, "cells", InspectionResult.CellsResult.class);

    Path preservedOutput = context.derivedWorkbook("signed-preserved");
    GridGrindResponse.Success preservedSave =
        XlsxParityGridGrind.mutateWorkbook(signed.workbookPath(), preservedOutput, List.of());
    boolean poiPreservedValid = XlsxParityOracle.signatureValid(preservedOutput);
    GridGrindResponse.Success preservedRead =
        XlsxParityGridGrind.readWorkbook(
            preservedOutput,
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()));
    InspectionResult.PackageSecurityResult preservedSecurityResult =
        XlsxParityGridGrind.read(
            preservedRead, "security", InspectionResult.PackageSecurityResult.class);

    XlsxParityScenarios.MaterializedScenario unsignedMutationSource =
        context.copiedScenario(XlsxParityScenarios.SIGNED_WORKBOOK, "signed-unsigned-mutation");
    Path unsignedOutput = context.derivedWorkbook("signed-unsigned-mutation");
    GridGrindResponse unsignedMutationResponse =
        XlsxParityGridGrind.executeMutateWorkbook(
            unsignedMutationSource.workbookPath(),
            unsignedOutput,
            List.of(
                mutate(
                    new CellSelector.ByAddress("Signed", "C1"),
                    new MutationAction.SetCell(text("Touch")))));
    if (!(unsignedMutationResponse instanceof GridGrindResponse.Failure unsignedMutationFailure)) {
      return fail("Mutating a signed workbook without explicit re-sign unexpectedly succeeded.");
    }

    XlsxParityScenarios.MaterializedScenario resignSource =
        context.copiedScenario(XlsxParityScenarios.SIGNED_WORKBOOK, "signed-resigned-source");
    Path resignedOutput = context.derivedWorkbook("signed-resigned");
    GridGrindResponse.Success resignedSave =
        XlsxParityGridGrind.mutateWorkbook(
            resignSource.workbookPath(),
            null,
            resignedOutput,
            new OoxmlPersistenceSecurityInput(
                null,
                signingInput(
                    resignSource.attachment(XlsxParityScenarios.SIGNING_PKCS12_ATTACHMENT),
                    "GridGrind parity re-sign")),
            null,
            List.of(
                mutate(
                    new CellSelector.ByAddress("Signed", "C1"),
                    new MutationAction.SetCell(text("Re-signed")))));
    boolean poiResignedValid = XlsxParityOracle.signatureValid(resignedOutput);
    GridGrindResponse.Success resignedRead =
        XlsxParityGridGrind.readWorkbook(
            resignedOutput,
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Signed", List.of("C1")),
                new InspectionQuery.GetCells()));
    InspectionResult.PackageSecurityResult resignedSecurityResult =
        XlsxParityGridGrind.read(
            resignedRead, "security", InspectionResult.PackageSecurityResult.class);
    InspectionResult.CellsResult resignedCells =
        XlsxParityGridGrind.read(resignedRead, "cells", InspectionResult.CellsResult.class);

    Path authoredOutput = context.derivedWorkbook("signed-authored");
    GridGrindResponse.Success authoredSave =
        XlsxParityGridGrind.writeNewWorkbook(
            authoredOutput,
            new OoxmlPersistenceSecurityInput(
                null,
                signingInput(
                    signed.attachment(XlsxParityScenarios.SIGNING_PKCS12_ATTACHMENT),
                    "GridGrind parity authored signature")),
            (ExecutionModeInput) null,
            List.of(
                mutate(new SheetSelector.ByName("Signed"), new MutationAction.EnsureSheet()),
                mutate(
                    new CellSelector.ByAddress("Signed", "A1"),
                    new MutationAction.SetCell(text("Authored signed")))));
    boolean poiAuthoredValid = XlsxParityOracle.signatureValid(authoredOutput);
    GridGrindResponse.Success authoredRead =
        XlsxParityGridGrind.readWorkbook(
            authoredOutput,
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Signed", List.of("A1")),
                new InspectionQuery.GetCells()));
    InspectionResult.PackageSecurityResult authoredSecurityResult =
        XlsxParityGridGrind.read(
            authoredRead, "security", InspectionResult.PackageSecurityResult.class);
    InspectionResult.CellsResult authoredCells =
        XlsxParityGridGrind.read(authoredRead, "cells", InspectionResult.CellsResult.class);

    return poiSourceValid
            && hasSingleSignatureState(
                sourceSecurityResult.security(), ExcelOoxmlSignatureState.VALID)
            && "Signed workbook".equals(textCell(sourceCells, "A1"))
            && XlsxParityGridGrind.savedPath(preservedSave)
                .equals(preservedOutput.toAbsolutePath().toString())
            && poiPreservedValid
            && hasSingleSignatureState(
                preservedSecurityResult.security(), ExcelOoxmlSignatureState.VALID)
            && unsignedMutationFailure.problem().code() == GridGrindProblemCode.INVALID_REQUEST
            && unsignedMutationFailure
                .problem()
                .message()
                .contains("persistence.security.signature")
            && !Files.exists(unsignedOutput)
            && XlsxParityGridGrind.savedPath(resignedSave)
                .equals(resignedOutput.toAbsolutePath().toString())
            && poiResignedValid
            && hasSingleSignatureState(
                resignedSecurityResult.security(), ExcelOoxmlSignatureState.VALID)
            && "Re-signed".equals(textCell(resignedCells, "C1"))
            && XlsxParityGridGrind.savedPath(authoredSave)
                .equals(authoredOutput.toAbsolutePath().toString())
            && poiAuthoredValid
            && hasSingleSignatureState(
                authoredSecurityResult.security(), ExcelOoxmlSignatureState.VALID)
            && "Authored signed".equals(textCell(authoredCells, "A1"))
        ? pass(
            "OOXML signing parity is present for signature validation, unchanged signature preservation, explicit re-signing, and authored signed save-as.")
        : fail(
            "OOXML signing parity mismatch."
                + " poiSourceValid="
                + poiSourceValid
                + " sourceValid="
                + hasSingleSignatureState(
                    sourceSecurityResult.security(), ExcelOoxmlSignatureState.VALID)
                + " poiPreservedValid="
                + poiPreservedValid
                + " preservedValid="
                + hasSingleSignatureState(
                    preservedSecurityResult.security(), ExcelOoxmlSignatureState.VALID)
                + " unsignedFailureCode="
                + unsignedMutationFailure.problem().code()
                + " unsignedOutputExists="
                + Files.exists(unsignedOutput)
                + " poiResignedValid="
                + poiResignedValid
                + " resignedValid="
                + hasSingleSignatureState(
                    resignedSecurityResult.security(), ExcelOoxmlSignatureState.VALID)
                + " poiAuthoredValid="
                + poiAuthoredValid
                + " authoredValid="
                + hasSingleSignatureState(
                    authoredSecurityResult.security(), ExcelOoxmlSignatureState.VALID));
  }

  static ProbeResult probeSigningInvalidSignature(ProbeContext context) {
    XlsxParityScenarios.MaterializedScenario invalid =
        context.scenario(XlsxParityScenarios.INVALID_SIGNATURE_WORKBOOK);
    boolean poiInvalid = !XlsxParityOracle.signatureValid(invalid.workbookPath());
    GridGrindResponse.Success read =
        XlsxParityGridGrind.readWorkbook(
            invalid.workbookPath(),
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Signed", List.of("A1")),
                new InspectionQuery.GetCells()));
    InspectionResult.PackageSecurityResult security =
        XlsxParityGridGrind.read(read, "security", InspectionResult.PackageSecurityResult.class);
    InspectionResult.CellsResult cells =
        XlsxParityGridGrind.read(read, "cells", InspectionResult.CellsResult.class);

    return poiInvalid
            && hasSingleSignatureState(security.security(), ExcelOoxmlSignatureState.INVALID)
            && "Signed workbook".equals(textCell(cells, "A1"))
        ? pass(
            "Invalid OOXML package signatures degrade into truthful INVALID signature facts instead of aborting the read.")
        : fail(
            "Invalid OOXML signature parity mismatch."
                + " poiInvalid="
                + poiInvalid
                + " gridState="
                + security.security().signatures());
  }
}
