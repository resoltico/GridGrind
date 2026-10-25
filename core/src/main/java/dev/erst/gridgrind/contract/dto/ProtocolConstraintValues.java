package dev.erst.gridgrind.contract.dto;

/** Stable scalar bounds and patterns shared by protocol validation and catalog publication. */
public final class ProtocolConstraintValues {
  /** Pattern accepted for every authored workbook step identifier. */
  public static final String STEP_ID_PATTERN = "[A-Za-z0-9._-]+";

  /** Pattern accepted for one protocol RGB color literal. */
  public static final String RGB_HEX_PATTERN = "^#[0-9A-Fa-f]{6}$";

  /** Pattern documenting the Unicode identifier alphabet accepted for authored defined names. */
  public static final String DEFINED_NAME_PATTERN = "^[\\p{L}_\\\\][\\p{L}\\p{N}_.\\\\]*$";

  /** Maximum Unicode code-point count for one authored Excel defined name. */
  public static final int DEFINED_NAME_MAX_CODE_POINTS = 255;

  /** Minimum percentage for an Excel data-bar width. */
  public static final int DATA_BAR_WIDTH_MIN = 0;

  /** Maximum percentage for an Excel data-bar width. */
  public static final int DATA_BAR_WIDTH_MAX = 100;

  /** Required suffix for source and persisted workbook paths. */
  public static final String WORKBOOK_PATH_SUFFIX = ".xlsx";

  private ProtocolConstraintValues() {}
}
