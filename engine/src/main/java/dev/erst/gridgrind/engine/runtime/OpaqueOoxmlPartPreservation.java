package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.Preservation;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/** Establishes byte identity for explicitly host-required OOXML package parts. */
final class OpaqueOoxmlPartPreservation {
  private static final int BUFFER_SIZE = 16 * 1024;

  private OpaqueOoxmlPartPreservation() {}

  static Preservation.ByteEstablished establish(
      Path sourceArtifact, Path stagedArtifact, List<String> partNames)
      throws IOException, PreservationFailedException {
    Objects.requireNonNull(sourceArtifact, "sourceArtifact must not be null");
    Objects.requireNonNull(stagedArtifact, "stagedArtifact must not be null");
    Objects.requireNonNull(partNames, "partNames must not be null");
    try (ZipFile source = new ZipFile(sourceArtifact.toFile());
        ZipFile staged = new ZipFile(stagedArtifact.toFile())) {
      for (String partName : partNames) {
        String requiredPartName =
            Objects.requireNonNull(partName, "partNames must not contain null values");
        if (!identicalPart(source, staged, requiredPartName)) {
          throw new PreservationFailedException(Preservation.Kind.BYTE, List.of(requiredPartName));
        }
      }
    }
    return new Preservation.ByteEstablished(partNames);
  }

  private static boolean identicalPart(ZipFile source, ZipFile staged, String partName)
      throws IOException {
    ZipEntry sourcePart = source.getEntry(zipEntryName(partName));
    ZipEntry stagedPart = staged.getEntry(zipEntryName(partName));
    if (sourcePart == null || stagedPart == null || sourcePart.getSize() != stagedPart.getSize()) {
      return false;
    }
    try (InputStream sourceBytes = source.getInputStream(sourcePart);
        InputStream stagedBytes = staged.getInputStream(stagedPart)) {
      return identicalBytes(sourceBytes, stagedBytes);
    }
  }

  static boolean identicalBytes(InputStream source, InputStream staged) throws IOException {
    byte[] sourceBuffer = new byte[BUFFER_SIZE];
    byte[] stagedBuffer = new byte[BUFFER_SIZE];
    while (true) {
      int sourceRead = source.read(sourceBuffer);
      int stagedRead = staged.read(stagedBuffer);
      if (sourceRead != stagedRead) {
        return false;
      }
      if (sourceRead < 0) {
        return true;
      }
      for (int index = 0; index < sourceRead; index++) {
        if (sourceBuffer[index] != stagedBuffer[index]) {
          return false;
        }
      }
    }
  }

  private static String zipEntryName(String partName) {
    return partName.substring(1);
  }
}
