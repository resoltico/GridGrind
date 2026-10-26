package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.lang.foreign.FunctionDescriptor;
import java.lang.foreign.MemorySegment;
import java.lang.foreign.SymbolLookup;
import java.lang.foreign.ValueLayout;
import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.MethodType;
import java.util.List;
import java.util.Locale;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicBoolean;
import org.junit.jupiter.api.Test;

/** Focused coverage tests for platform-native stdin probes. */
class NativeStdinProbeTest {
  @Test
  void currentProcessFactoryReturnsAnInstance() {
    NativeStdinProbe probe = NativeStdinProbe.currentProcess();

    assertNotNull(probe);
    probe.getAsBoolean();
  }

  @Test
  void libraryLookupReturnsLookupOnSuccess() {
    SymbolLookup lookup = name -> Optional.empty();

    assertEquals(Optional.of(lookup), NativeStdinProbe.libraryLookup(() -> lookup));
  }

  @Test
  void libraryLookupReturnsEmptyWhenLookupFails() {
    Optional<SymbolLookup> lookup =
        NativeStdinProbe.libraryLookup(
            () -> {
              throw new UnsatisfiedLinkError("library not available");
            });

    assertTrue(lookup.isEmpty());
  }

  @Test
  void libraryLookupSupplierReturnsLookupOnSuccess() {
    SymbolLookup lookup = name -> Optional.empty();

    assertEquals(Optional.of(lookup), NativeStdinProbe.libraryLookupSupplier(() -> lookup).get());
  }

  @Test
  void libraryLookupSupplierReturnsEmptyWhenLookupFails() {
    assertTrue(
        NativeStdinProbe.libraryLookupSupplier(
                () -> {
                  throw new UnsatisfiedLinkError("library not available");
                })
            .get()
            .isEmpty());
  }

  @Test
  void kernel32LibraryLookupReturnsAnOptionalWithoutThrowing() {
    assertNotNull(NativeStdinProbe.kernel32LibraryLookup());
  }

  @Test
  void kernel32LibraryLookupUsesInjectedFactoryForSuccessAndFailure() {
    assertTrue(
        NativeStdinProbe.kernel32LibraryLookup("Windows 11", "definitely-missing").isEmpty());
    assertTrue(
        NativeStdinProbe.kernel32LibraryLookup("Windows 11", "definitely-missing").isEmpty());
    assertTrue(NativeStdinProbe.kernel32LibraryLookup("Linux", "must-not-be-used").isEmpty());
  }

  @Test
  void namedLibraryLookupReturnsEmptyForMissingLibraries() {
    assertTrue(NativeStdinProbe.namedLibraryLookup("definitely-missing").isEmpty());
  }

  @Test
  void namedLibraryLookupCanResolveOneKnownRuntimeLibraryForCurrentHost() {
    String osName = System.getProperty("os.name", "");
    Optional<SymbolLookup> lookup =
        knownRuntimeLibraries(osName).stream()
            .map(NativeStdinProbe::namedLibraryLookup)
            .filter(Optional::isPresent)
            .map(Optional::get)
            .findFirst();

    assertTrue(
        lookup.isPresent(), "expected one known runtime library lookup to succeed for " + osName);
  }

  @Test
  void downcallHandleProviderReturnsAProviderForTheSuppliedLinker() {
    java.lang.foreign.Linker linker = java.lang.foreign.Linker.nativeLinker();
    MemorySegment isattySymbol = linker.defaultLookup().find("isatty").orElseThrow();
    MethodHandle handle =
        NativeStdinProbe.downcallHandleProvider(linker)
            .downcallHandle(
                isattySymbol, FunctionDescriptor.of(ValueLayout.JAVA_INT, ValueLayout.JAVA_INT));

    assertNotNull(handle);
  }

  @Test
  void unixCurrentProcessFactoryReturnsAnInstanceForCurrentEnvironment() {
    NativeStdinProbe probe = NativeStdinProbe.unixCurrentProcess();

    assertNotNull(probe);
    probe.getAsBoolean();
  }

  @Test
  void windowsCurrentProcessFactoryReturnsAnInstanceForCurrentEnvironment() {
    NativeStdinProbe probe = NativeStdinProbe.windowsCurrentProcess();

    assertNotNull(probe);
    probe.getAsBoolean();
  }

  @Test
  void windowsCurrentProcessUsesLibraryLookupWhenAvailable() throws Throwable {
    NativeStdinProbe probe =
        NativeStdinProbe.windowsCurrentProcess(
            () -> Optional.of(symbolLookup()),
            (symbol, descriptor) ->
                symbol.address() == 11L
                    ? getStdHandleHandle(MemorySegment.ofAddress(99L))
                    : getConsoleModeHandle(1));

    assertTrue(probe.getAsBoolean());
  }

  @Test
  void currentProcessUsesWindowsFactoryForWindowsNames() {
    AtomicBoolean windowsFactoryCalled = new AtomicBoolean();
    AtomicBoolean unixFactoryCalled = new AtomicBoolean();

    NativeStdinProbe probe =
        NativeStdinProbe.currentProcess(
            "Windows 11",
            () -> {
              windowsFactoryCalled.set(true);
              return NativeStdinProbe.of(() -> true);
            },
            () -> {
              unixFactoryCalled.set(true);
              return NativeStdinProbe.of(() -> false);
            });

    assertTrue(probe.getAsBoolean());
    assertTrue(windowsFactoryCalled.get());
    assertFalse(unixFactoryCalled.get());
  }

  @Test
  void currentProcessUsesUnixFactoryForNonWindowsNames() {
    AtomicBoolean windowsFactoryCalled = new AtomicBoolean();
    AtomicBoolean unixFactoryCalled = new AtomicBoolean();

    NativeStdinProbe probe =
        NativeStdinProbe.currentProcess(
            "Mac OS X",
            () -> {
              windowsFactoryCalled.set(true);
              return NativeStdinProbe.of(() -> false);
            },
            () -> {
              unixFactoryCalled.set(true);
              return NativeStdinProbe.of(() -> true);
            });

    assertTrue(probe.getAsBoolean());
    assertFalse(windowsFactoryCalled.get());
    assertTrue(unixFactoryCalled.get());
  }

  @Test
  void reportsFalseWhenNativeProbeThrows() {
    NativeStdinProbe probe =
        NativeStdinProbe.of(
            () -> {
              throw new UnsatisfiedLinkError("missing symbol");
            });

    assertFalse(probe.getAsBoolean());
  }

  @Test
  void unsupportedFactoryReturnsSingletonInstance() {
    assertSame(NativeStdinProbe.unsupported(), NativeStdinProbe.unsupported());
    assertFalse(NativeStdinProbe.unsupported().getAsBoolean());
  }

  @Test
  void identifiesWindowsFromOsName() {
    assertTrue(NativeStdinProbe.isWindows("Windows Server 2025"));
    assertFalse(NativeStdinProbe.isWindows("Linux"));
  }

  @Test
  void unixCurrentProcessUsesInjectedIsattyHandle() throws Throwable {
    NativeStdinProbe probe =
        NativeStdinProbe.unixCurrentProcess(
            name -> MemorySegment.ofAddress(1L),
            (symbol, descriptor) -> {
              assertEquals(
                  FunctionDescriptor.of(ValueLayout.JAVA_INT, ValueLayout.JAVA_INT), descriptor);
              return isattyHandle(1);
            });

    assertTrue(probe.getAsBoolean());
  }

  @Test
  void unixCurrentProcessReturnsUnsupportedWhenLookupFails() {
    NativeStdinProbe probe =
        NativeStdinProbe.currentProcessProbe(
            () ->
                NativeStdinProbe.unixCurrentProcess(
                    name -> {
                      throw new NativeStdinProbe.NativeProbeUnavailableException("missing " + name);
                    },
                    (symbol, descriptor) -> {
                      throw new AssertionError("must not request a downcall handle");
                    }));

    assertFalse(probe.getAsBoolean());
  }

  @Test
  void currentProcessProbeReturnsUnsupportedWhenFactoryThrows() {
    NativeStdinProbe probe =
        NativeStdinProbe.currentProcessProbe(
            () -> {
              throw new NativeStdinProbe.NativeProbeUnavailableException("missing native symbol");
            });

    assertFalse(probe.getAsBoolean());
  }

  @Test
  void currentProcessProbeReturnsUnsupportedWhenFactoryThrowsLinkageError() {
    NativeStdinProbe probe =
        NativeStdinProbe.currentProcessProbe(
            () -> {
              throw new UnsatisfiedLinkError("native runtime unavailable");
            });

    assertFalse(probe.getAsBoolean());
  }

  @Test
  void windowsCurrentProcessReportsInteractiveWhenConsoleModeSucceeds() throws Throwable {
    NativeStdinProbe probe =
        NativeStdinProbe.windowsCurrentProcessFromLookup(
            Optional.of(symbolLookup()),
            (symbol, descriptor) ->
                symbol.address() == 11L
                    ? getStdHandleHandle(MemorySegment.ofAddress(99L))
                    : getConsoleModeHandle(1));

    assertTrue(probe.getAsBoolean());
  }

  @Test
  void windowsCurrentProcessUsesInjectedSymbolLookup() throws Throwable {
    NativeStdinProbe probe =
        NativeStdinProbe.windowsCurrentProcess(
            symbolResolver(),
            (symbol, descriptor) ->
                symbol.address() == 11L
                    ? getStdHandleHandle(MemorySegment.ofAddress(99L))
                    : getConsoleModeHandle(1));

    assertTrue(probe.getAsBoolean());
  }

  @Test
  void windowsCurrentProcessReturnsUnsupportedWhenLookupIsUnavailable() {
    NativeStdinProbe probe =
        NativeStdinProbe.windowsCurrentProcessFromLookup(
            Optional.empty(),
            (symbol, descriptor) -> {
              throw new AssertionError("must not request a downcall handle");
            });

    assertFalse(probe.getAsBoolean());
  }

  @Test
  void windowsCurrentProcessReturnsUnsupportedWhenLookupOmitsRequiredSymbols() {
    NativeStdinProbe probe =
        NativeStdinProbe.windowsCurrentProcessFromLookup(
            Optional.of(name -> Optional.empty()),
            (symbol, descriptor) -> {
              throw new AssertionError("must not request a downcall handle");
            });

    assertFalse(probe.getAsBoolean());
  }

  @Test
  void windowsCurrentProcessReturnsFalseWhenStandardInputHandleIsNull() throws Throwable {
    NativeStdinProbe probe =
        NativeStdinProbe.windowsCurrentProcess(
            symbolResolver(),
            (symbol, descriptor) ->
                symbol.address() == 11L
                    ? getStdHandleHandle(MemorySegment.NULL)
                    : getConsoleModeHandle(1));

    assertFalse(probe.getAsBoolean());
  }

  @Test
  void windowsCurrentProcessReturnsFalseWhenConsoleModeProbeFails() throws Throwable {
    NativeStdinProbe probe =
        NativeStdinProbe.windowsCurrentProcess(
            symbolResolver(),
            (symbol, descriptor) ->
                symbol.address() == 11L
                    ? getStdHandleHandle(MemorySegment.ofAddress(199L))
                    : getConsoleModeHandle(0));

    assertFalse(probe.getAsBoolean());
  }

  @Test
  void windowsCurrentProcessReturnsUnsupportedWhenLookupFails() {
    NativeStdinProbe probe =
        NativeStdinProbe.currentProcessProbe(
            () ->
                NativeStdinProbe.windowsCurrentProcess(
                    name -> {
                      throw new NativeStdinProbe.NativeProbeUnavailableException("missing " + name);
                    },
                    (symbol, descriptor) -> {
                      throw new AssertionError("must not request a downcall handle");
                    }));

    assertFalse(probe.getAsBoolean());
  }

  @Test
  void propagatesUnexpectedProbeFailuresInsteadOfMaskingThemAsNonInteractive() {
    NativeStdinProbe probe =
        NativeStdinProbe.of(
            () -> {
              throw new IllegalStateException("unexpected failure");
            });

    IllegalStateException failure = assertThrows(IllegalStateException.class, probe::getAsBoolean);
    assertEquals("unexpected failure", failure.getMessage());
  }

  @Test
  void propagatesErrorsInsteadOfNormalizingThemToNonInteractive() {
    NativeStdinProbe probe =
        NativeStdinProbe.of(
            () -> {
              throw new AssertionError("boom");
            });

    AssertionError failure = assertThrows(AssertionError.class, probe::getAsBoolean);
    assertEquals("boom", failure.getMessage());
  }

  @Test
  void wrapsCheckedFailuresInIllegalStateException() {
    NativeStdinProbe probe =
        NativeStdinProbe.of(
            () -> {
              throw new java.io.IOException("checked boom");
            });

    IllegalStateException failure = assertThrows(IllegalStateException.class, probe::getAsBoolean);
    assertEquals("Native standard-input probe threw a checked failure", failure.getMessage());
    assertInstanceOf(java.io.IOException.class, failure.getCause());
    assertEquals("checked boom", failure.getCause().getMessage());
  }

  private static NativeStdinProbe.SymbolResolver symbolResolver() {
    return name ->
        switch (name) {
          case "GetStdHandle" -> MemorySegment.ofAddress(11L);
          case "GetConsoleMode" -> MemorySegment.ofAddress(12L);
          default -> throw new IllegalArgumentException("unexpected symbol " + name);
        };
  }

  private static SymbolLookup symbolLookup() {
    return name ->
        switch (name) {
          case "GetStdHandle" -> Optional.of(MemorySegment.ofAddress(11L));
          case "GetConsoleMode" -> Optional.of(MemorySegment.ofAddress(12L));
          default -> Optional.empty();
        };
  }

  private static List<String> knownRuntimeLibraries(String osName) {
    if (NativeStdinProbe.isWindows(osName)) {
      return List.of("Kernel32", "kernel32");
    }
    if (osName.toLowerCase(Locale.ROOT).contains("mac")) {
      return List.of("libSystem.B.dylib", "System", "c");
    }
    return List.of("libc.so.6", "c");
  }

  private static MethodHandle isattyHandle(int result) {
    return bindStaticHandle(
        "isattyResult", MethodType.methodType(int.class, int.class, int.class), result);
  }

  private static MethodHandle getStdHandleHandle(MemorySegment result) {
    return bindStaticHandle(
        "stdHandleResult",
        MethodType.methodType(MemorySegment.class, MemorySegment.class, int.class),
        result);
  }

  private static MethodHandle getConsoleModeHandle(int result) {
    return bindStaticHandle(
        "consoleModeResult",
        MethodType.methodType(int.class, int.class, MemorySegment.class, MemorySegment.class),
        result);
  }

  private static MethodHandle bindStaticHandle(
      String methodName, MethodType methodType, Object boundArgument) {
    try {
      return MethodHandles.insertArguments(
          MethodHandles.lookup().findStatic(NativeStdinProbeTest.class, methodName, methodType),
          0,
          boundArgument);
    } catch (ReflectiveOperationException exception) {
      throw new LinkageError("failed to build synthetic native handle for test", exception);
    }
  }

  @SuppressWarnings("unused")
  private static int isattyResult(int result, int ignoredDescriptor) {
    return result;
  }

  @SuppressWarnings("unused")
  private static MemorySegment stdHandleResult(MemorySegment result, int ignoredStdHandle) {
    return result;
  }

  @SuppressWarnings("unused")
  private static int consoleModeResult(
      int result, MemorySegment ignoredHandle, MemorySegment ignoredModePointer) {
    return result;
  }
}
