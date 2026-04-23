module dev.erst.gridgrind.engine {
  requires dev.erst.gridgrind.excel.foundation;
  requires java.desktop;
  requires java.xml;
  requires java.xml.crypto;
  requires org.apache.poi.poi;
  requires org.apache.poi.ooxml;
  requires org.apache.santuario.xmlsec;

  exports dev.erst.gridgrind.excel;
}
