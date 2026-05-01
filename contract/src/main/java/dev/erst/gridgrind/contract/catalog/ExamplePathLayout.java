package dev.erst.gridgrind.contract.catalog;

/** Path layout policy for generated example requests. */
enum ExamplePathLayout {
  BUILT_IN("generated-workbooks/", "examples/"),
  REPOSITORY("generated-workbooks/", "");

  private final String generatedWorkbookPrefix;
  private final String assetPrefix;

  ExamplePathLayout(String generatedWorkbookPrefix, String assetPrefix) {
    this.generatedWorkbookPrefix = generatedWorkbookPrefix;
    this.assetPrefix = assetPrefix;
  }

  String generatedWorkbook(String fileName) {
    return generatedWorkbookPrefix + fileName;
  }

  String asset(String relativePath) {
    return assetPrefix + relativePath;
  }
}
