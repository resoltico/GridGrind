package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.engine.security.SecurityAdapter;
import dev.erst.gridgrind.engine.security.spi.OoxmlRelationshipTransformService;
import javax.xml.crypto.dsig.TransformService;
import org.junit.jupiter.api.Test;

/** Regression coverage for GridGrind's xmlsec 4 OOXML dsig provider override. */
class ExcelOoxmlDsigProviderSupportTest {
  @Test
  void relationshipTransformResolvesToGridGrindXmlSec4Service() throws Exception {
    TransformService transformService =
        TransformService.getInstance(
            OoxmlRelationshipTransformService.TRANSFORM_URI, "DOM", SecurityAdapter.provider());

    assertEquals(OoxmlRelationshipTransformService.class, transformService.getClass());
  }
}
