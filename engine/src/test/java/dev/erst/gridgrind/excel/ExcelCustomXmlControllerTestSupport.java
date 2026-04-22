package dev.erst.gridgrind.excel;

import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.VarHandle;
import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import javax.xml.transform.Result;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.URIResolver;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.model.MapInfo;
import org.apache.poi.xssf.usermodel.XSSFMap;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTMap;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSchema;
import org.w3c.dom.Node;

/** Shared helpers for focused custom-XML controller tests. */
final class ExcelCustomXmlControllerTestSupport {
  private static final MethodHandles.Lookup LOOKUP = MethodHandles.lookup();
  private static final VarHandle MAP_INFO_MAPS = mapInfoMapsHandle();
  private static final MethodHandle XSSF_MAP_CONSTRUCTOR = xssfMapConstructor();

  private ExcelCustomXmlControllerTestSupport() {}

  static void addDuplicateMapping(XSSFWorkbook workbook) {
    Collection<XSSFMap> mappings = workbook.getCustomXMLMappings();
    XSSFMap original = mappings.iterator().next();
    MapInfo mapInfo = workbook.getMapInfo();
    CTMap duplicate = mapInfo.getCTMapInfo().addNewMap();
    duplicate.set(original.getCtMap());
    duplicate.setID(original.getCtMap().getID() + 1);
    duplicate.setName(original.getCtMap().getName());
    mutableMapRegistry(mapInfo)
        .put(Math.toIntExact(duplicate.getID()), instantiateMap(duplicate, mapInfo));
  }

  static FakeMap fakeMap(
      CTMap ctMap,
      CTSchema schema,
      Node schemaNode,
      List<org.apache.poi.xssf.usermodel.helpers.XSSFSingleXmlCell> linkedCells,
      List<XSSFTable> linkedTables) {
    return new FakeMap(ctMap, schema, schemaNode, linkedCells, linkedTables);
  }

  static FakeTable fakeTable(XSSFSheet sheet) {
    return new FakeTable(sheet);
  }

  static Transformer failingTransformer() {
    return new FailingTransformer();
  }

  @SuppressWarnings("unchecked")
  private static Map<Integer, XSSFMap> mutableMapRegistry(MapInfo mapInfo) {
    return (Map<Integer, XSSFMap>) MAP_INFO_MAPS.get(mapInfo);
  }

  private static XSSFMap instantiateMap(CTMap ctMap, MapInfo mapInfo) {
    try {
      return (XSSFMap) XSSF_MAP_CONSTRUCTOR.invoke(ctMap, mapInfo);
    } catch (RuntimeException | Error exception) {
      throw exception;
    } catch (Throwable throwable) {
      throw new IllegalStateException(
          "Failed to instantiate duplicate XSSFMap for test coverage", throwable);
    }
  }

  private static VarHandle mapInfoMapsHandle() {
    try {
      return MethodHandles.privateLookupIn(MapInfo.class, LOOKUP)
          .findVarHandle(MapInfo.class, "maps", Map.class);
    } catch (NoSuchFieldException | IllegalAccessException exception) {
      throw new ExceptionInInitializerError(exception);
    }
  }

  private static MethodHandle xssfMapConstructor() {
    try {
      return MethodHandles.privateLookupIn(XSSFMap.class, LOOKUP)
          .findConstructor(
              XSSFMap.class,
              java.lang.invoke.MethodType.methodType(void.class, CTMap.class, MapInfo.class));
    } catch (NoSuchMethodException | IllegalAccessException exception) {
      throw new ExceptionInInitializerError(exception);
    }
  }

  /**
   * Minimal fake map that drives snapshot and projection branches without a full workbook package.
   */
  static final class FakeMap extends XSSFMap {
    private final CTMap ctMap;
    private final CTSchema schema;
    private final Node schemaNode;
    private final List<org.apache.poi.xssf.usermodel.helpers.XSSFSingleXmlCell> linkedCells;
    private final List<XSSFTable> linkedTables;

    private FakeMap(
        CTMap ctMap,
        CTSchema schema,
        Node schemaNode,
        List<org.apache.poi.xssf.usermodel.helpers.XSSFSingleXmlCell> linkedCells,
        List<XSSFTable> linkedTables) {
      super(ctMap, new MapInfo());
      this.ctMap = ctMap;
      this.schema = schema;
      this.schemaNode = schemaNode;
      this.linkedCells = linkedCells;
      this.linkedTables = linkedTables;
    }

    @Override
    public CTMap getCtMap() {
      return ctMap;
    }

    @Override
    public CTSchema getCTSchema() {
      return schema;
    }

    @Override
    public Node getSchema() {
      return schemaNode;
    }

    @Override
    public List<org.apache.poi.xssf.usermodel.helpers.XSSFSingleXmlCell> getRelatedSingleXMLCell() {
      return linkedCells;
    }

    @Override
    public List<XSSFTable> getRelatedTables() {
      return linkedTables;
    }
  }

  /** Minimal fake table that supplies linked-table metadata without table package plumbing. */
  static final class FakeTable extends XSSFTable {
    private final XSSFSheet sheet;

    private FakeTable(XSSFSheet sheet) {
      this.sheet = sheet;
    }

    @Override
    public XSSFSheet getXSSFSheet() {
      return sheet;
    }

    @Override
    public String getName() {
      return "OrdersTable";
    }

    @Override
    public String getDisplayName() {
      return "Orders Table";
    }

    @Override
    public AreaReference getArea() {
      return new AreaReference("Orders!A1:B2", SpreadsheetVersion.EXCEL2007);
    }

    @Override
    public String getCommonXpath() {
      return "/Orders/Order";
    }
  }

  /**
   * Transformer that deterministically fails so schema-serialization translation can be asserted.
   */
  static final class FailingTransformer extends Transformer {
    @Override
    public void transform(Source xmlSource, Result outputTarget) throws TransformerException {
      throw new TransformerException("boom");
    }

    @Override
    public void setParameter(String name, Object value) {}

    @Override
    public Object getParameter(String name) {
      return null;
    }

    @Override
    public void clearParameters() {}

    @Override
    public void setURIResolver(URIResolver resolver) {}

    @Override
    public URIResolver getURIResolver() {
      return null;
    }

    @Override
    public void setOutputProperties(Properties oformat) {}

    @Override
    public Properties getOutputProperties() {
      return new Properties();
    }

    @Override
    public void setOutputProperty(String name, String value) {}

    @Override
    public String getOutputProperty(String name) {
      return null;
    }

    @Override
    public void setErrorListener(javax.xml.transform.ErrorListener listener) {}

    @Override
    public javax.xml.transform.ErrorListener getErrorListener() {
      return null;
    }
  }
}
