open module dev.erst.gridgrind.architecture.tests {
  requires com.tngtech.archunit;
  requires com.tngtech.archunit.junit6.api;
  requires dev.erst.gridgrind.core;
  requires dev.erst.gridgrind.json;
  requires dev.erst.gridgrind.cli;
  requires dev.erst.gridgrind.engine;

  exports dev.erst.gridgrind.architecture;
}
