package dev.erst.gridgrind.engine.bootstrap;

import static org.junit.jupiter.api.Assertions.*;

import java.io.*;
import java.util.concurrent.atomic.AtomicInteger;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;

/**
 * Required repair resources cannot fall back to stock bytes; every class read transfers no handle.
 */
@Tag("native-public")
@SuppressWarnings("PMD.UseProperClassLoader")
// The fixture replaces resources on the artifact loader; host TCCL would test a different boundary.
class NativeProviderResourceTest {
  private static final String PREFIX = "META-INF/gridgrind/native-provider/";

  @Test
  void missingRepairRefusesInsteadOfLoadingTheAvailableStockClass() throws Exception {
    var loader =
        new NativeProviderLoader(
            new ClassLoader(getClass().getClassLoader()) {
              @Override
              public InputStream getResourceAsStream(String name) {
                return (PREFIX + "classes/org/apache/poi/ss/usermodel/DateUtil.class").equals(name)
                    ? null
                    : super.getResourceAsStream(name);
              }
            });
    assertNotNull(
        getClass().getClassLoader().getResource("org/apache/poi/ss/usermodel/DateUtil.class"));
    var failure =
        assertThrows(
            ClassNotFoundException.class,
            () -> loader.loadClass("org.apache.poi.ss.usermodel.DateUtil"));
    assertTrue(failure.getMessage().contains("Required native provider repair is missing"));
  }

  @Test
  void anEmptyRequiredListIsNotAnImplicitUnmodifiedProviderConfiguration() {
    assertThrows(
        IllegalStateException.class,
        () ->
            new NativeProviderLoader(
                new ClassLoader(getClass().getClassLoader()) {
                  @Override
                  public InputStream getResourceAsStream(String name) {
                    return (PREFIX + "required-classes.list").equals(name)
                        ? new ByteArrayInputStream(new byte[0])
                        : super.getResourceAsStream(name);
                  }
                }));
  }

  @Test
  void bothStockAndCapsuleResourceStreamsCloseBeforeClassDefinitionReturns() throws Exception {
    AtomicInteger opened = new AtomicInteger(), closed = new AtomicInteger();
    var loader =
        new NativeProviderLoader(
            new ClassLoader(getClass().getClassLoader()) {
              @Override
              public InputStream getResourceAsStream(String name) {
                InputStream input = super.getResourceAsStream(name);
                if (input == null) return null;
                opened.incrementAndGet();
                return new FilterInputStream(input) {
                  @Override
                  public void close() throws IOException {
                    super.close();
                    closed.incrementAndGet();
                  }
                };
              }
            });
    Class<?> repaired = loader.loadClass("org.apache.poi.ss.usermodel.DateUtil");
    assertSame(loader, repaired.getClassLoader());
    Class<?> stock = loader.loadClass("org.apache.poi.util.Units");
    assertSame(loader, stock.getClassLoader());
    assertTrue(opened.get() > 1);
    assertEquals(opened.get(), closed.get());
  }
}
