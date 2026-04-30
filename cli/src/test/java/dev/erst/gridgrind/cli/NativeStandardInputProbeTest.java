package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertSame;
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
class NativeStandardInputProbeTest {
  @Test
  void currentProcessFactoryReturnsAnInstance() {
    NativeStandardInputProbe probe = NativeStandardInputProbe.currentProcess();

    assertNotNull(probe);
    probe.getAsBoolean();
  }

  @Test
  void libraryLookupReturnsLookupOnSuccess() {
    SymbolLookup lookup = name -> Optional.empty();

    assertEquals(Optional.of(lookup), NativeStandardInputProbe.libraryLookup(() -> lookup));
  }

  @Test
  void libraryLookupReturnsEmptyWhenLookupFails() {
    Optional<SymbolLookup> lookup =
        NativeStandardInputProbe.libraryLookup(
            () -> {
              throw new UnsatisfiedLinkError("library not available");
            });

    assertTrue(lookup.isEmpty());
  }

  @Test
  void libraryLookupSupplierReturnsLookupOnSuccess() {
    SymbolLookup lookup = name -> Optional.empty();

    assertEquals(
        Optional.of(lookup), NativeStandardInputProbe.libraryLookupSupplier(() -> lookup).get());
  }

  @Test
  void libraryLookupSupplierReturnsEmptyWhenLookupFails() {
    assertTrue(
        NativeStandardInputProbe.libraryLookupSupplier(
                () -> {
                  throw new UnsatisfiedLinkError("library not available");
                })
            .get()
            .isEmpty());
  }

  @Test
  void kernel32LibraryLookupReturnsAnOptionalWithoutThrowing() {
    assertNotNull(NativeStandardInputProbe.kernel32LibraryLookup());
  }

  @Test
  void kernel32LibraryLookupUsesInjectedFactoryForSuccessAndFailure() {
    assertTrue(
        NativeStandardInputProbe.kernel32LibraryLookup("Windows 11", "definitely-missing")
            .isEmpty());
    assertTrue(
        NativeStandardInputProbe.kernel32LibraryLookup("Windows 11", "definitely-missing")
            .isEmpty());
    assertTrue(
        NativeStandardInputProbe.kernel32LibraryLookup("Linux", "must-not-be-used").isEmpty());
  }

  @Test
  void namedLibraryLookupReturnsEmptyForMissingLibraries() {
    assertTrue(NativeStandardInputProbe.namedLibraryLookup("definitely-missing").isEmpty());
  }

  @Test
  void namedLibraryLookupCanResolveOneKnownRuntimeLibraryForCurrentHost() {
    String osName = System.getProperty("os.name", "");
    Optional<SymbolLookup> lookup =
        knownRuntimeLibraries(osName).stream()
            .map(NativeStandardInputProbe::namedLibraryLookup)
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
        NativeStandardInputProbe.downcallHandleProvider(linker)
            .downcallHandle(
                isattySymbol, FunctionDescriptor.of(ValueLayout.JAVA_INT, ValueLayout.JAVA_INT));

    assertNotNull(handle);
  }

  @Test
  void unixCurrentProcessFactoryReturnsAnInstanceForCurrentEnvironment() {
    NativeStandardInputProbe probe = NativeStandardInputProbe.unixCurrentProcess();

    assertNotNull(probe);
    probe.getAsBoolean();
  }

  @Test
  void windowsCurrentProcessFactoryReturnsAnInstanceForCurrentEnvironment() {
    NativeStandardInputProbe probe = NativeStandardInputProbe.windowsCurrentProcess();

    assertNotNull(probe);
    probe.getAsBoolean();
  }

  @Test
  void windowsCurrentProcessUsesLibraryLookupWhenAvailable() throws Throwable {
    NativeStandardInputProbe probe =
        NativeStandardInputProbe.windowsCurrentProcess(
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

    NativeStandardInputProbe probe =
        NativeStandardInputProbe.currentProcess(
            "Windows 11",
            () -> {
              windowsFactoryCalled.set(true);
              return NativeStandardInputProbe.of(() -> true);
            },
            () -> {
              unixFactoryCalled.set(true);
              return NativeStandardInputProbe.of(() -> false);
            });

    assertTrue(probe.getAsBoolean());
    assertTrue(windowsFactoryCalled.get());
    assertFalse(unixFactoryCalled.get());
  }

  @Test
  void currentProcessUsesUnixFactoryForNonWindowsNames() {
    AtomicBoolean windowsFactoryCalled = new AtomicBoolean();
    AtomicBoolean unixFactoryCalled = new AtomicBoolean();

    NativeStandardInputProbe probe =
        NativeStandardInputProbe.currentProcess(
            "Mac OS X",
            () -> {
              windowsFactoryCalled.set(true);
              return NativeStandardInputProbe.of(() -> false);
            },
            () -> {
              unixFactoryCalled.set(true);
              return NativeStandardInputProbe.of(() -> true);
            });

    assertTrue(probe.getAsBoolean());
    assertFalse(windowsFactoryCalled.get());
    assertTrue(unixFactoryCalled.get());
  }

  @Test
  void reportsFalseWhenNativeProbeThrows() {
    NativeStandardInputProbe probe =
        NativeStandardInputProbe.of(
            () -> {
              throw new UnsatisfiedLinkError("missing symbol");
            });

    assertFalse(probe.getAsBoolean());
  }

  @Test
  void unsupportedFactoryReturnsSingletonInstance() {
    assertSame(NativeStandardInputProbe.unsupported(), NativeStandardInputProbe.unsupported());
    assertFalse(NativeStandardInputProbe.unsupported().getAsBoolean());
  }

  @Test
  void identifiesWindowsFromOsName() {
    assertTrue(NativeStandardInputProbe.isWindows("Windows Server 2025"));
    assertFalse(NativeStandardInputProbe.isWindows("Linux"));
  }

  @Test
  void unixCurrentProcessUsesInjectedIsattyHandle() throws Throwable {
    NativeStandardInputProbe probe =
        NativeStandardInputProbe.unixCurrentProcess(
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
    NativeStandardInputProbe probe =
        NativeStandardInputProbe.currentProcessProbe(
            () ->
                NativeStandardInputProbe.unixCurrentProcess(
                    name -> {
                      throw new IllegalArgumentException("missing " + name);
                    },
                    (symbol, descriptor) -> {
                      throw new AssertionError("must not request a downcall handle");
                    }));

    assertFalse(probe.getAsBoolean());
  }

  @Test
  void currentProcessProbeReturnsUnsupportedWhenFactoryThrows() {
    NativeStandardInputProbe probe =
        NativeStandardInputProbe.currentProcessProbe(
            () -> {
              throw new IllegalArgumentException("missing symbol");
            });

    assertFalse(probe.getAsBoolean());
  }

  @Test
  void currentProcessProbeReturnsUnsupportedWhenFactoryThrowsLinkageError() {
    NativeStandardInputProbe probe =
        NativeStandardInputProbe.currentProcessProbe(
            () -> {
              throw new UnsatisfiedLinkError("native runtime unavailable");
            });

    assertFalse(probe.getAsBoolean());
  }

  @Test
  void windowsCurrentProcessReportsInteractiveWhenConsoleModeSucceeds() throws Throwable {
    NativeStandardInputProbe probe =
        NativeStandardInputProbe.windowsCurrentProcessFromLookup(
            symbolLookup(),
            (symbol, descriptor) ->
                symbol.address() == 11L
                    ? getStdHandleHandle(MemorySegment.ofAddress(99L))
                    : getConsoleModeHandle(1));

    assertTrue(probe.getAsBoolean());
  }

  @Test
  void windowsCurrentProcessUsesInjectedSymbolLookup() throws Throwable {
    NativeStandardInputProbe probe =
        NativeStandardInputProbe.windowsCurrentProcess(
            symbolResolver(),
            (symbol, descriptor) ->
                symbol.address() == 11L
                    ? getStdHandleHandle(MemorySegment.ofAddress(99L))
                    : getConsoleModeHandle(1));

    assertTrue(probe.getAsBoolean());
  }

  @Test
  void windowsCurrentProcessReturnsUnsupportedWhenLookupIsUnavailable() {
    NativeStandardInputProbe probe =
        NativeStandardInputProbe.windowsCurrentProcessFromLookup(
            (SymbolLookup) null,
            (symbol, descriptor) -> {
              throw new AssertionError("must not request a downcall handle");
            });

    assertFalse(probe.getAsBoolean());
  }

  @Test
  void windowsCurrentProcessReturnsUnsupportedWhenLookupOmitsRequiredSymbols() {
    NativeStandardInputProbe probe =
        NativeStandardInputProbe.windowsCurrentProcessFromLookup(
            name -> Optional.empty(),
            (symbol, descriptor) -> {
              throw new AssertionError("must not request a downcall handle");
            });

    assertFalse(probe.getAsBoolean());
  }

  @Test
  void windowsCurrentProcessReturnsFalseWhenStandardInputHandleIsNull() throws Throwable {
    NativeStandardInputProbe probe =
        NativeStandardInputProbe.windowsCurrentProcess(
            symbolResolver(),
            (symbol, descriptor) ->
                symbol.address() == 11L
                    ? getStdHandleHandle(MemorySegment.NULL)
                    : getConsoleModeHandle(1));

    assertFalse(probe.getAsBoolean());
  }

  @Test
  void windowsCurrentProcessReturnsFalseWhenConsoleModeProbeFails() throws Throwable {
    NativeStandardInputProbe probe =
        NativeStandardInputProbe.windowsCurrentProcess(
            symbolResolver(),
            (symbol, descriptor) ->
                symbol.address() == 11L
                    ? getStdHandleHandle(MemorySegment.ofAddress(199L))
                    : getConsoleModeHandle(0));

    assertFalse(probe.getAsBoolean());
  }

  @Test
  void windowsCurrentProcessReturnsUnsupportedWhenLookupFails() {
    NativeStandardInputProbe probe =
        NativeStandardInputProbe.currentProcessProbe(
            () ->
                NativeStandardInputProbe.windowsCurrentProcess(
                    name -> {
                      throw new IllegalArgumentException("missing " + name);
                    },
                    (symbol, descriptor) -> {
                      throw new AssertionError("must not request a downcall handle");
                    }));

    assertFalse(probe.getAsBoolean());
  }

  private static NativeStandardInputProbe.SymbolResolver symbolResolver() {
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
    if (NativeStandardInputProbe.isWindows(osName)) {
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
          MethodHandles.lookup()
              .findStatic(NativeStandardInputProbeTest.class, methodName, methodType),
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
