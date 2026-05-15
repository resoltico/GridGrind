package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlDsigProviderSupport;
import dev.erst.gridgrind.excel.spi.ExcelOoxmlXmlSec4RelationshipTransformService;
import javax.xml.crypto.dsig.TransformService;
import org.junit.jupiter.api.Test;

/** Regression coverage for GridGrind's xmlsec 4 OOXML dsig provider override. */
class ExcelOoxmlDsigProviderSupportTest {
  @Test
  void relationshipTransformResolvesToGridGrindXmlSec4Service() throws Exception {
    TransformService transformService =
        TransformService.getInstance(
            ExcelOoxmlXmlSec4RelationshipTransformService.TRANSFORM_URI,
            "DOM",
            ExcelOoxmlDsigProviderSupport.provider());

    assertEquals(ExcelOoxmlXmlSec4RelationshipTransformService.class, transformService.getClass());
  }
}
