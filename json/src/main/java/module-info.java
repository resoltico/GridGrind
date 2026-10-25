module dev.erst.gridgrind.json {
  requires transitive dev.erst.gridgrind.core;
  requires com.fasterxml.jackson.annotation;
  requires tools.jackson.databind;
  requires static org.jspecify;

  exports dev.erst.gridgrind.json.codec;
}
