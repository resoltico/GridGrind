module dev.erst.gridgrind.engine {
  requires transitive dev.erst.gridgrind.contract;
  requires dev.erst.gridgrind.excel.foundation;
  requires java.desktop;
  requires java.xml;
  requires java.xml.crypto;
  requires static org.jspecify;
  requires org.apache.poi.poi;
  requires org.apache.poi.ooxml;
  requires org.apache.santuario.xmlsec;

  exports dev.erst.gridgrind.engine.api;

  opens dev.erst.gridgrind.excel.spi to
      java.xml.crypto;
}
