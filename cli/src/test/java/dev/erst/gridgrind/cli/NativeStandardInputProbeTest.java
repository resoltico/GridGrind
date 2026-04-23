package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.lang.foreign.FunctionDescriptor;
import java.lang.foreign.MemorySegment;
import java.lang.foreign.SymbolLookup;
import java.lang.foreign.ValueLayout;
import java.lang.invoke.MethodHandle;
import java.lang.invoke.MethodHandles;
import java.lang.invoke.MethodType;
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
  void libraryLookupOrNullReturnsLookupOnSuccess() {
    SymbolLookup lookup = name -> Optional.empty();

    assertSame(lookup, NativeStandardInputProbe.libraryLookupOrNull(() -> lookup));
  }

  @Test
  void libraryLookupOrNullReturnsNullWhenLookupFails() {
    SymbolLookup lookup =
        NativeStandardInputProbe.libraryLookupOrNull(
            () -> {
              throw new UnsatisfiedLinkError("library not available");
            });

    assertNull(lookup);
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
