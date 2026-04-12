package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;

/** Package-owned seam for checked POI package and relationship inspection. */
final class PoiPackageInspection {
  private PoiPackageInspection() {}

  static List<PackagePart> packageParts(OPCPackage pkg, String failureMessage) {
    Objects.requireNonNull(pkg, "pkg must not be null");
    return packageParts(pkg::getParts, failureMessage);
  }

  static List<PackageRelationship> relationships(PackagePart part, String failureMessage) {
    Objects.requireNonNull(part, "part must not be null");
    return relationships(part::getRelationships, failureMessage);
  }

  static List<PackagePart> packageParts(
      CheckedPackagePartsSupplier supplier, String failureMessage) {
    Objects.requireNonNull(supplier, "supplier must not be null");
    Objects.requireNonNull(failureMessage, "failureMessage must not be null");
    try {
      List<PackagePart> parts = new ArrayList<>();
      for (PackagePart part : supplier.get()) {
        parts.add(part);
      }
      return List.copyOf(parts);
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException(failureMessage, exception);
    }
  }

  static List<PackageRelationship> relationships(
      CheckedPackageRelationshipsSupplier supplier, String failureMessage) {
    Objects.requireNonNull(supplier, "supplier must not be null");
    Objects.requireNonNull(failureMessage, "failureMessage must not be null");
    try {
      List<PackageRelationship> relationships = new ArrayList<>();
      for (PackageRelationship relationship : supplier.get()) {
        relationships.add(relationship);
      }
      return List.copyOf(relationships);
    } catch (InvalidFormatException exception) {
      throw new IllegalStateException(failureMessage, exception);
    }
  }

  /** Supplies package parts while preserving POI's checked InvalidFormatException contract. */
  @FunctionalInterface
  interface CheckedPackagePartsSupplier {
    /** Returns the current package-part set. */
    Iterable<PackagePart> get() throws InvalidFormatException;
  }

  /**
   * Supplies package relationships while preserving POI's checked InvalidFormatException contract.
   */
  @FunctionalInterface
  interface CheckedPackageRelationshipsSupplier {
    /** Returns the current relationship set. */
    Iterable<PackageRelationship> get() throws InvalidFormatException;
  }
}
