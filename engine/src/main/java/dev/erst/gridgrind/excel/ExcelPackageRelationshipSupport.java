package dev.erst.gridgrind.excel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.jspecify.annotations.Nullable;

/** Shared package-relationship traversal and cleanup helpers for OOXML package maintenance. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelPackageRelationshipSupport {
  private ExcelPackageRelationshipSupport() {}

  /** Supplies package parts while allowing callers to surface invalid-package failures. */
  @FunctionalInterface
  public interface PackagePartsSupplier {
    /** Returns the package parts to inspect. */
    Iterable<PackagePart> get() throws InvalidFormatException;
  }

  public static boolean partIsStillReferenced(
      Iterable<PackagePart> packageParts, @Nullable PackagePartName partName) {
    if (partName == null) {
      return false;
    }
    try {
      for (PackagePart part : packageParts) {
        if (part.isRelationshipPart()) {
          continue;
        }
        for (PackageRelationship relationship : part.getRelationships()) {
          if (relationship.getTargetMode() == TargetMode.EXTERNAL) {
            continue;
          }
          if (partName
              .getURI()
              .equals(
                  PackagingURIHelper.resolvePartUri(
                      part.getPartName().getURI(), relationship.getTargetURI()))) {
            return true;
          }
        }
      }
      return false;
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException("Failed to inspect package relationships", exception);
    }
  }

  public static Iterable<PackagePart> requireParts(PackagePartsSupplier packagePartsSupplier) {
    try {
      return packagePartsSupplier.get();
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException("Failed to inspect package relationships", exception);
    }
  }

  public static void cleanupPackagePartIfUnused(
      OPCPackage pkg, @Nullable PackagePartName partName) {
    if (partName == null) {
      return;
    }
    if (partIsStillReferenced(requireParts(pkg::getParts), partName)) {
      return;
    }
    if (pkg.containPart(partName)) {
      pkg.deletePartRecursive(partName);
    }
  }
}
