package dev.erst.gridgrind.engine.api;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.grant.*;
import dev.erst.gridgrind.core.model.formula.*;
import dev.erst.gridgrind.core.model.security.*;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.value.*;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.WorkbookPlan.WorkbookPersistence.IfExists;
import dev.erst.gridgrind.core.problem.*;
import dev.erst.gridgrind.core.result.*;
import dev.erst.gridgrind.core.source.*;
import dev.erst.gridgrind.core.vocabulary.*;
import dev.erst.gridgrind.engine.backend.xssf.OoxmlSecurityTestSupport;
import java.lang.ref.WeakReference;
import java.nio.file.*;
import java.time.*;
import java.util.*;
import java.util.concurrent.atomic.*;
import java.util.zip.*;
import org.apache.poi.util.LocaleUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/** Actual published ports use stock parent providers and the delivered resource capsule. */
@Tag("native-public")
class NativePublicExecutionTest {
  @TempDir Path root;
  private static final SheetName DATA = new SheetName("Data");

  @ParameterizedTest
  @CsvSource({"de-DE,America/Los_Angeles", "ar-EG,Europe/Riga", "th-TH,Pacific/Auckland"})
  void fixedCivilAndTextPolicyPreservesHostDefaultsAndParentProviderBytes(String tag, String zone)
      throws Exception {
    Locale beforeLocale = Locale.getDefault();
    TimeZone beforeZone = TimeZone.getDefault();
    LocaleUtil.resetUserLocale();
    LocaleUtil.resetUserTimeZone();
    byte[] parent = classBytes(org.apache.poi.ss.usermodel.DataFormatter.class);
    assertThrows(
        NoSuchMethodException.class,
        () ->
            org.apache.poi.ss.usermodel.DateUtil.class.getDeclaredMethod(
                "parseDateTime", String.class, boolean.class),
        "Parent classpath must use the original vendor provider, not compiler patch directories");
    try {
      Locale.setDefault(Locale.forLanguageTag(tag));
      TimeZone.setDefault(TimeZone.getTimeZone(zone));
      List<WorkbookStep> steps =
          new ArrayList<>(
              List.of(
                  sheet(),
                  write(
                      "a",
                      "A1",
                      new CellValueSpec.DateTime(LocalDateTime.parse("2026-03-08T02:30:00"))),
                  write(
                      "b",
                      "A2",
                      new CellValueSpec.DateTime(LocalDateTime.parse("2026-03-29T03:30:00"))),
                  write("n", "B1", new CellValueSpec.NumberValue(1234.5)),
                  write(
                      "text",
                      "B2",
                      new CellValueSpec.Formula(TextSourceSpec.inline("TEXT(B1,\"0.00\")"))),
                  new MutationStep(
                      new StepId("format"),
                      new RangeSelector.ByRange(DATA, new RangeRef("A1:A2")),
                      new CellMutationAction.ApplyStyle(
                          new CellStylePatch(
                              Optional.of("yyyy-mm-dd hh:mm:ss"),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty(),
                              Optional.empty()))),
                  new AssertionStep(
                      new StepId("la"),
                      cell("A1"),
                      new CellAssertion.DisplayValue("2026-03-08 02:30:00")),
                  new AssertionStep(
                      new StepId("riga"),
                      cell("A2"),
                      new CellAssertion.DisplayValue("2026-03-29 03:30:00")),
                  new AssertionStep(
                      new StepId("computed"),
                      cell("B2"),
                      new CellAssertion.Value(new CellValue.Text("1234.50")))));
      var result =
          GridGrindEngine.requestExecutor()
              .execute(
                  plan(
                      steps,
                      new WorkbookPlan.WorkbookSource.New(),
                      new WorkbookPlan.WorkbookPersistence.None()),
                  inputs(List.of(), null));
      assertInstanceOf(WorkbookResult.Success.class, result, result.toString());
      assertEquals(Locale.forLanguageTag(tag), Locale.getDefault());
      assertEquals(zone, TimeZone.getDefault().getID());
      assertArrayEquals(parent, classBytes(org.apache.poi.ss.usermodel.DataFormatter.class));
      assertEquals(Locale.getDefault(), LocaleUtil.getUserLocale());
      assertEquals(TimeZone.getDefault(), LocaleUtil.getUserTimeZone());
      Locale.setDefault(Locale.ITALY);
      TimeZone.setDefault(TimeZone.getTimeZone("Pacific/Auckland"));
      assertEquals(
          Locale.ITALY,
          LocaleUtil.getUserLocale(),
          "An absent host policy must not become a sticky explicit value");
      assertEquals("Pacific/Auckland", LocaleUtil.getUserTimeZone().getID());
    } finally {
      Locale.setDefault(beforeLocale);
      TimeZone.setDefault(beforeZone);
      LocaleUtil.resetUserLocale();
      LocaleUtil.resetUserTimeZone();
    }
  }

  @Test
  void actual1904DateProducerAndConsumerUseTheWorkbookEpoch() throws Exception {
    Path source = root.resolve("mac.xlsx");
    try (var workbook = new XSSFWorkbook()) {
      workbook.getCTWorkbook().getWorkbookPr().setDate1904(true);
      workbook.createSheet("Data");
      try (var output = Files.newOutputStream(source)) {
        workbook.write(output);
      }
    }
    var steps =
        List.<WorkbookStep>of(
            write("date", "A1", new CellValueSpec.Date(LocalDate.of(1904, 1, 1))),
            write(
                "date-fn",
                "B1",
                new CellValueSpec.Formula(TextSourceSpec.inline("DATE(2026,3,8)"))),
            write(
                "text-fn",
                "C1",
                new CellValueSpec.Formula(
                    TextSourceSpec.inline("TEXT(DATE(2026,3,8),\"yyyy-mm-dd\")"))),
            new AssertionStep(
                new StepId("serial"),
                cell("B1"),
                new CellAssertion.Value(new CellValue.NumberValue(44627))),
            new AssertionStep(
                new StepId("civil"),
                cell("C1"),
                new CellAssertion.Value(new CellValue.Text("2026-03-08"))));
    var result =
        GridGrindEngine.requestExecutor()
            .execute(
                plan(
                    steps,
                    new WorkbookPlan.WorkbookSource.ExistingFile(source.toString()),
                    new WorkbookPlan.WorkbookPersistence.None()),
                inputs(List.of(new ExecutionGrant.ReadAuthority.File(source)), null));
    assertInstanceOf(WorkbookResult.Success.class, result, result.toString());
    var invalid =
        List.<WorkbookStep>of(
            write("first", "D1", new CellValueSpec.NumberValue(9)),
            write("too-early", "D2", new CellValueSpec.Date(LocalDate.of(1903, 12, 31))));
    var failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            GridGrindEngine.requestExecutor()
                .execute(
                    plan(
                        invalid,
                        new WorkbookPlan.WorkbookSource.ExistingFile(source.toString()),
                        new WorkbookPlan.WorkbookPersistence.None()),
                    inputs(List.of(new ExecutionGrant.ReadAuthority.File(source)), null)));
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals(
        0,
        assertInstanceOf(ExecutionJournal.Outcome.Failed.class, failure.journal().outcome())
            .completedStepCount());
    assertTrue(failure.problem().message().contains("MAC_1904"));
  }

  @Test
  void mixedEpochExternalEvaluationRestoresTheActualCallingWorkbookContext() throws Exception {
    Path main = root.resolve("main.xlsx"), external = root.resolve("external.xlsx");
    try (var workbook = new XSSFWorkbook()) {
      workbook.getCTWorkbook().getWorkbookPr().setDate1904(true);
      workbook.createSheet("Data");
      try (var out = Files.newOutputStream(main)) {
        workbook.write(out);
      }
    }
    try (var workbook = new XSSFWorkbook()) {
      var row = workbook.createSheet("Other").createRow(0);
      row.createCell(0).setCellFormula("TEXT(DATE(2026,3,8),\"yyyy-mm-dd\")");
      try (var out = Files.newOutputStream(external)) {
        workbook.write(out);
      }
    }
    var steps =
        List.<WorkbookStep>of(
            write(
                "external",
                "A1",
                new CellValueSpec.Formula(TextSourceSpec.inline("'[Source.xlsx]Other'!A1"))),
            write(
                "main-date",
                "B1",
                new CellValueSpec.Formula(TextSourceSpec.inline("DATE(2026,3,8)"))),
            new AssertionStep(
                new StepId("external-text"),
                cell("A1"),
                new CellAssertion.Value(new CellValue.Text("2026-03-08"))),
            new AssertionStep(
                new StepId("main-serial"),
                cell("B1"),
                new CellAssertion.Value(new CellValue.NumberValue(44627))));
    var basic =
        plan(
            steps,
            new WorkbookPlan.WorkbookSource.ExistingFile(main.toString()),
            new WorkbookPlan.WorkbookPersistence.None());
    var request =
        WorkbookPlan.standard(
            basic.source(),
            basic.persistence(),
            basic.execution(),
            new FormulaEnvironmentSpec(
                List.of(new FormulaExternalWorkbookSpec("Source.xlsx", external.toString())),
                FormulaMissingWorkbookPolicy.ERROR,
                List.of()),
            steps);
    var result =
        GridGrindEngine.requestExecutor()
            .execute(
                request,
                inputs(
                    List.of(
                        new ExecutionGrant.ReadAuthority.File(main),
                        new ExecutionGrant.ReadAuthority.File(external)),
                    null));
    assertInstanceOf(WorkbookResult.Success.class, result, result.toString());
  }

  @Test
  void ordinaryAndCachedFormulaWidthsUseTheSameFixedMonthRendering() throws Exception {
    Path source = root.resolve("months.xlsx"), target = root.resolve("widths.xlsx");
    try (var workbook = new XSSFWorkbook()) {
      var sheet = workbook.createSheet("Data");
      var row = sheet.createRow(0);
      row.createCell(0).setCellValue(46089);
      row.createCell(1).setCellFormula("A1");
      var style = workbook.createCellStyle();
      style.setDataFormat(workbook.createDataFormat().getFormat("mmmm"));
      row.getCell(0).setCellStyle(style);
      row.getCell(1).setCellStyle(style);
      workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
      try (var out = Files.newOutputStream(source)) {
        workbook.write(out);
      }
    }
    Locale before = Locale.getDefault();
    try {
      Locale.setDefault(Locale.GERMANY);
      var steps =
          List.<WorkbookStep>of(
              new MutationStep(
                  new StepId("width"),
                  new SheetSelector.ByName(DATA),
                  new WorkbookMutationAction.AutoSizeColumns()));
      var request =
          WorkbookPlan.standard(
              new WorkbookPlan.WorkbookSource.ExistingFile(source.toString()),
              new WorkbookPlan.WorkbookPersistence.SaveAs(
                  target.toString(), IfExists.REJECT, OoxmlPersistenceSecuritySpec.none()),
              ExecutionPolicySpec.defaults(),
              FormulaEnvironmentSpec.empty(),
              steps);
      var result =
          GridGrindEngine.requestExecutor()
              .execute(
                  request, inputs(List.of(new ExecutionGrant.ReadAuthority.File(source)), target));
      assertInstanceOf(WorkbookResult.Success.class, result, result.toString());
    } finally {
      Locale.setDefault(before);
    }
    try (var zip = new ZipFile(target.toFile());
        var input =
            zip.getInputStream(Objects.requireNonNull(zip.getEntry("xl/worksheets/sheet1.xml")))) {
      var document =
          javax.xml.parsers.DocumentBuilderFactory.newDefaultNSInstance()
              .newDocumentBuilder()
              .parse(input);
      var widths =
          document.getElementsByTagNameNS(
              "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "col");
      var covered = new java.util.HashSet<Integer>();
      for (int index = 0; index < widths.getLength(); index++) {
        var column = (org.w3c.dom.Element) widths.item(index);
        int first = Integer.parseInt(column.getAttribute("min")),
            last = Integer.parseInt(column.getAttribute("max"));
        for (int value = first; value <= last; value++)
          if (value <= 2) {
            assertEquals(
                7d,
                Double.parseDouble(column.getAttribute("width")),
                "March has five codepoints plus the declared two-character padding");
            covered.add(value);
          }
      }
      assertEquals(
          java.util.Set.of(1, 2),
          covered,
          "Both ordinary and cached-formula widths must be serialized");
      var cells =
          document.getElementsByTagNameNS(
              "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "c");
      boolean cached = false;
      for (int index = 0; index < cells.getLength(); index++) {
        var cell = (org.w3c.dom.Element) cells.item(index);
        if ("B1".equals(cell.getAttribute("r"))) {
          assertEquals(
              "A1",
              cell.getElementsByTagNameNS(
                      "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "f")
                  .item(0)
                  .getTextContent());
          assertEquals(
              46089d,
              Double.parseDouble(
                  cell.getElementsByTagNameNS(
                          "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "v")
                      .item(0)
                      .getTextContent()));
          cached = true;
        }
      }
      assertTrue(
          cached, "The stored formula cache remains available to the declared width operation");
    }
  }

  @Test
  void callbacksStayOnCallerWithExplicitHostPolicyThroughNestingFailureAndInterruption()
      throws Exception {
    LocaleUtil.setUserLocale(Locale.FRANCE);
    LocaleUtil.setUserTimeZone(TimeZone.getTimeZone("Asia/Tokyo"));
    Thread caller = Thread.currentThread();
    ClassLoader tccl = caller.getContextClassLoader();
    var steps =
        List.<WorkbookStep>of(
            sheet(), write("read", "A1", new CellValueSpec.Text(TextSourceSpec.standardInput())));
    var request =
        plan(
            steps,
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None());
    AtomicInteger opens = new AtomicInteger(),
        closes = new AtomicInteger(),
        events = new AtomicInteger();
    AtomicBoolean nested = new AtomicBoolean();
    Runnable host =
        () -> {
          assertSame(caller, Thread.currentThread());
          assertSame(tccl, caller.getContextClassLoader());
          assertEquals(Locale.FRANCE, LocaleUtil.getUserLocale());
          assertEquals("Asia/Tokyo", LocaleUtil.getUserTimeZone().getID());
        };
    try {
      var grant = grant(List.of(new ExecutionGrant.ReadAuthority.Stdin()), null, List.of());
      var in =
          new GridGrindRequestInputs(
              root,
              root.resolve("scratch"),
              () -> {
                host.run();
                opens.incrementAndGet();
                return new java.io.ByteArrayInputStream(
                    "owned payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)) {
                  @Override
                  public void close() {
                    host.run();
                    closes.incrementAndGet();
                  }
                };
              },
              grant,
              Optional.empty());
      var result =
          GridGrindEngine.requestExecutor()
              .execute(
                  request,
                  in,
                  event -> {
                    host.run();
                    events.incrementAndGet();
                    if (nested.compareAndSet(false, true)) {
                      var nestedSteps = List.<WorkbookStep>of(sheet());
                      assertInstanceOf(
                          WorkbookResult.Success.class,
                          GridGrindEngine.requestExecutor()
                              .execute(
                                  plan(
                                      nestedSteps,
                                      new WorkbookPlan.WorkbookSource.New(),
                                      new WorkbookPlan.WorkbookPersistence.None()),
                                  inputs(List.of(), null)));
                    }
                  });
      assertInstanceOf(WorkbookResult.Success.class, result, result.toString());
      assertEquals(1, opens.get());
      assertEquals(1, closes.get());
      assertTrue(events.get() > 0);
      host.run();
      var primary = new AssertionError("owned fatal callback");
      var simple =
          List.<WorkbookStep>of(sheet(), write("write", "A1", new CellValueSpec.NumberValue(1)));
      var error =
          assertThrows(
              AssertionError.class,
              () ->
                  GridGrindEngine.requestExecutor()
                      .execute(
                          plan(
                              simple,
                              new WorkbookPlan.WorkbookSource.New(),
                              new WorkbookPlan.WorkbookPersistence.None()),
                          inputs(List.of(), null),
                          event -> {
                            if (event.stepId().map(id -> "write".equals(id.value())).orElse(false)
                                && event.status() == ExecutionProgressEvent.Status.SUCCEEDED) {
                              Thread.currentThread().interrupt();
                              throw primary;
                            }
                          }));
      assertSame(primary, error);
      assertTrue(Thread.currentThread().isInterrupted());
      Thread.interrupted();
      host.run();
      assertNoScratchFiles();
    } finally {
      Thread.interrupted();
      LocaleUtil.resetUserLocale();
      LocaleUtil.resetUserTimeZone();
    }
  }

  @Test
  void actualPrivateSantuarioProviderSignsEncryptsDecryptsAndVerifies() throws Exception {
    var material = OoxmlSecurityTestSupport.createSignedWorkbook(root.resolve("fixture"));
    var store = new SecretReference("store");
    var key = new SecretReference("key");
    var password = new SecretReference("password");
    var steps =
        List.<WorkbookStep>of(
            sheet(),
            write(
                "write",
                "A1",
                new CellValueSpec.Text(TextSourceSpec.inline("Synthetic protected value"))));
    Path target = root.resolve("protected.xlsx");
    var security =
        new OoxmlPersistenceSecuritySpec(
            new OoxmlPersistenceEncryptionSpec.Encrypt(OoxmlEncryptionSpec.strong(password)),
            new OoxmlPersistenceSignatureSpec.Sign(
                new OoxmlSignatureSpec(
                    material.pkcs12Path().toString(),
                    store,
                    Optional.of(key),
                    Optional.of(material.alias()),
                    ExcelOoxmlSignatureDigestAlgorithm.SHA256,
                    Optional.empty())));
    var request =
        plan(
            steps,
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                target.toString(), IfExists.REJECT, security));
    var caller = Thread.currentThread();
    AtomicInteger callbacks = new AtomicInteger();
    GridGrindSecretResolver resolver =
        ref -> {
          assertSame(caller, Thread.currentThread());
          callbacks.incrementAndGet();
          return switch (ref.id()) {
            case "store" -> material.keystorePassword().toCharArray();
            case "key" -> material.keyPassword().toCharArray();
            case "password" -> OoxmlSecurityTestSupport.ENCRYPTION_PASSWORD.toCharArray();
            default -> throw new AssertionError("unknown synthetic reference");
          };
        };
    var input =
        new GridGrindRequestInputs(
            root,
            root.resolve("scratch"),
            grant(
                List.of(new ExecutionGrant.ReadAuthority.File(material.pkcs12Path())),
                target,
                List.of(store, key, password)),
            resolver);
    var result = GridGrindEngine.requestExecutor().execute(request, input);
    assertInstanceOf(WorkbookResult.Success.class, result, result.toString());
    var reads =
        List.<WorkbookStep>of(
            new InspectionStep(
                new StepId("security"),
                new WorkbookSelector.Current(),
                new WorkbookIntrospectionQuery.GetPackageSecurity()));
    var opened =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile(
                target.toString(), new OoxmlOpenSecuritySpec(Optional.of(password))),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicySpec.defaults(),
            FormulaEnvironmentSpec.empty(),
            reads);
    var reopened =
        assertInstanceOf(
            WorkbookResult.Success.class,
            GridGrindEngine.requestExecutor()
                .execute(
                    opened,
                    new GridGrindRequestInputs(
                        root,
                        root.resolve("scratch"),
                        grant(
                            List.of(new ExecutionGrant.ReadAuthority.File(target)),
                            null,
                            List.of(password)),
                        resolver)));
    var report =
        assertInstanceOf(
            WorkbookInspectionResult.PackageSecurityResult.class,
            reopened.inspections().getFirst());
    assertTrue(
        report.security().signatures().stream()
            .anyMatch(signature -> signature.state() == ExcelOoxmlSignatureState.VALID),
        report.security().toString());
    assertTrue(callbacks.get() > 0);
    assertNoScratchFiles();
  }

  @Test
  void realmsAreNotStronglyCachedByTheSharedPublishedFacade() throws Exception {
    assertRealmRetention(1);
  }

  @Test
  @Tag("native-lifetime")
  void realmsAreReclaimedWhenJvmSoftAdaptationCachesAreCleared() throws Exception {
    assertRealmRetention(0);
  }

  @SuppressWarnings("PMD.DoNotCallGarbageCollectionExplicitly")
  // Bounded GC requests are the observable trigger for this reclamation test, never production
  // policy.
  private void assertRealmRetention(int maximumRetained)
      throws java.io.IOException, InterruptedException {
    List<WeakReference<ClassLoader>> realms = new ArrayList<>();
    var steps = List.<WorkbookStep>of(sheet());
    var request =
        plan(
            steps,
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None());
    for (int run = 0; run < 8; run++) realms.add(executeAndObserveRealm(request));
    for (int attempt = 0; attempt < 8; attempt++) {
      System.gc();
      Thread.sleep(25);
      if (realms.stream().allMatch(ref -> ref.get() == null)) break;
    }
    long retained = realms.stream().filter(ref -> ref.get() != null).count();
    assertTrue(
        retained <= maximumRetained,
        "Retained " + retained + " realms; permitted " + maximumRetained);
    assertNoScratchFiles();
  }

  @SuppressWarnings("PMD.UseProperClassLoader")
  // This test observes the executing implementation's loader identity, never the host TCCL.
  private WeakReference<ClassLoader> executeAndObserveRealm(WorkbookPlan request) {
    AtomicReference<WeakReference<ClassLoader>> observed = new AtomicReference<>();
    assertInstanceOf(
        WorkbookResult.Success.class,
        GridGrindEngine.requestExecutor()
            .execute(
                request,
                inputs(List.of(), null),
                event -> {
                  if (observed.get() == null) {
                    ClassLoader realm =
                        StackWalker.getInstance(StackWalker.Option.RETAIN_CLASS_REFERENCE)
                            .walk(frames ->
                                frames
                                    .filter(frame ->
                                        "dev.erst.gridgrind.engine.runtime.NativeEnginePort"
                                            .equals(frame.getClassName()))
                                    .map(frame -> frame.getDeclaringClass().getClassLoader())
                                    .findFirst()
                                    .orElseThrow());
                    assertNotSame(GridGrindEngine.class.getClassLoader(), realm);
                    observed.set(new WeakReference<>(realm));
                  }
                }));
    return observed.get();
  }

  private static MutationStep sheet() {
    return new MutationStep(
        new StepId("sheet"),
        new SheetSelector.ByName(DATA),
        new WorkbookMutationAction.EnsureSheet());
  }

  private static MutationStep write(String id, String address, CellValueSpec value) {
    return new MutationStep(new StepId(id), cell(address), new CellMutationAction.SetCell(value));
  }

  private static CellSelector.ByAddress cell(String address) {
    return new CellSelector.ByAddress(DATA, new CellAddress(address));
  }

  private static WorkbookPlan plan(
      List<WorkbookStep> steps,
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence) {
    return WorkbookPlan.standard(
        source,
        persistence,
        new ExecutionPolicySpec(
            ExecutionModeSpec.defaults(),
            new ExecutionJournalSpec(ExecutionJournalLevel.VERBOSE),
            new CalculationPolicySpec(new CalculationStrategySpec.EvaluateAll(), false),
            AssertionModeSpec.FAIL_FAST),
        FormulaEnvironmentSpec.empty(),
        steps);
  }

  private GridGrindRequestInputs inputs(List<ExecutionGrant.ReadAuthority> reads, Path target) {
    return new GridGrindRequestInputs(
        root, root.resolve("scratch"), grant(reads, target, List.of()));
  }

  private static ExecutionGrant grant(
      List<ExecutionGrant.ReadAuthority> reads, Path target, List<SecretReference> secrets) {
    return new ExecutionGrant(
        64 * 1024 * 1024,
        reads,
        List.of(
            "ENSURE_SHEET",
            "SET_CELL",
            "APPLY_STYLE",
            "EXPECT_DISPLAY_VALUE",
            "EXPECT_CELL_VALUE",
            "GET_PACKAGE_SECURITY",
            "IMPORT_CUSTOM_XML_MAPPING",
            "AUTO_SIZE_COLUMNS"),
        new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
        target == null
            ? new ExecutionGrant.PublicationAuthority.None()
            : new ExecutionGrant.PublicationAuthority.SaveAs(target, IfExists.REJECT),
        secrets,
        AcceptancePolicy.minimum());
  }

  private static byte[] classBytes(Class<?> type) throws java.io.IOException {
    try (var in = type.getResourceAsStream(type.getSimpleName() + ".class")) {
      return Objects.requireNonNull(in).readAllBytes();
    }
  }

  private void assertNoScratchFiles() throws java.io.IOException {
    Path scratch = root.resolve("scratch");
    if (Files.exists(scratch))
      try (var files = Files.walk(scratch)) {
        assertTrue(files.noneMatch(Files::isRegularFile));
      }
  }
}
