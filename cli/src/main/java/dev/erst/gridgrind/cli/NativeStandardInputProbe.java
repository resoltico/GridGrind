package dev.erst.gridgrind.cli;

import java.lang.foreign.Arena;
import java.lang.foreign.FunctionDescriptor;
import java.lang.foreign.Linker;
import java.lang.foreign.MemorySegment;
import java.lang.foreign.SymbolLookup;
import java.lang.foreign.ValueLayout;
import java.lang.invoke.MethodHandle;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import java.util.function.BooleanSupplier;
import java.util.function.Supplier;

/** JVM-local stdin TTY detection backed by native platform APIs. */
final class NativeStandardInputProbe implements BooleanSupplier {
  /** Builds one OS-specific native probe, possibly failing before the probe is usable. */
  @FunctionalInterface
  interface ProbeFactory {
    /** Creates one probe instance from the current process environment. */
    NativeStandardInputProbe create();
  }

  /** Resolves one native symbol to a callable memory segment. */
  @FunctionalInterface
  interface SymbolResolver {
    /** Returns the native symbol for the supplied name or throws if it is unavailable. */
    MemorySegment findOrThrow(String name);
  }

  /** Builds one FFM downcall handle for a resolved native symbol. */
  @FunctionalInterface
  interface DowncallHandleProvider {
    /** Returns a callable handle for the supplied symbol and function descriptor. */
    MethodHandle downcallHandle(MemorySegment symbol, FunctionDescriptor descriptor);
  }

  /** Checked boolean supplier used to wrap FFM-backed native probes. */
  @FunctionalInterface
  interface CheckedBooleanSupplier {
    /** Returns whether stdin is interactive, possibly throwing while probing native state. */
    boolean getAsBoolean() throws Throwable;
  }

  private static final int STDIN_FILENO = 0;
  private static final int STD_INPUT_HANDLE = -10;
  private static final NativeStandardInputProbe UNSUPPORTED =
      new NativeStandardInputProbe(() -> false);

  private final CheckedBooleanSupplier probe;

  static NativeStandardInputProbe currentProcess() {
    return currentProcess(
        System.getProperty("os.name", ""),
        NativeStandardInputProbe::windowsCurrentProcess,
        NativeStandardInputProbe::unixCurrentProcess);
  }

  static NativeStandardInputProbe currentProcess(
      String osName,
      Supplier<NativeStandardInputProbe> windowsProbeFactory,
      Supplier<NativeStandardInputProbe> unixProbeFactory) {
    Objects.requireNonNull(osName, "osName must not be null");
    Objects.requireNonNull(windowsProbeFactory, "windowsProbeFactory must not be null");
    Objects.requireNonNull(unixProbeFactory, "unixProbeFactory must not be null");
    return isWindows(osName) ? windowsProbeFactory.get() : unixProbeFactory.get();
  }

  static NativeStandardInputProbe unsupported() {
    return UNSUPPORTED;
  }

  static NativeStandardInputProbe of(CheckedBooleanSupplier probe) {
    return new NativeStandardInputProbe(probe);
  }

  static boolean isWindows(String osName) {
    return osName.toLowerCase(Locale.ROOT).contains("win");
  }

  static NativeStandardInputProbe currentProcessProbe(ProbeFactory factory) {
    try {
      return factory.create();
    } catch (RuntimeException | LinkageError exception) {
      return unsupported();
    }
  }

  static NativeStandardInputProbe unixCurrentProcess() {
    Linker linker = Linker.nativeLinker();
    return currentProcessProbe(
        () ->
            unixCurrentProcess(
                linker.defaultLookup()::findOrThrow,
                (symbol, descriptor) -> linker.downcallHandle(symbol, descriptor)));
  }

  static NativeStandardInputProbe unixCurrentProcess(
      SymbolResolver symbolResolver, DowncallHandleProvider downcallHandleProvider) {
    MethodHandle isatty =
        downcallHandleProvider.downcallHandle(
            symbolResolver.findOrThrow("isatty"),
            FunctionDescriptor.of(ValueLayout.JAVA_INT, ValueLayout.JAVA_INT));
    return of(() -> ((int) isatty.invokeExact(STDIN_FILENO)) == 1);
  }

  static NativeStandardInputProbe windowsCurrentProcess() {
    Linker linker = Linker.nativeLinker();
    return windowsCurrentProcess(
        (Supplier<Optional<SymbolLookup>>) NativeStandardInputProbe::kernel32LibraryLookup,
        downcallHandleProvider(linker));
  }

  static NativeStandardInputProbe windowsCurrentProcess(
      Supplier<Optional<SymbolLookup>> libraryLookupSupplier,
      DowncallHandleProvider downcallHandleProvider) {
    SymbolLookup kernel32Lookup = libraryLookupSupplier.get().orElse(null);
    return windowsCurrentProcessFromLookup(kernel32Lookup, downcallHandleProvider);
  }

  static Optional<SymbolLookup> libraryLookup(Supplier<SymbolLookup> lookupFactory) {
    try {
      return Optional.of(lookupFactory.get());
    } catch (RuntimeException | LinkageError exception) {
      return Optional.empty();
    }
  }

  static Supplier<Optional<SymbolLookup>> libraryLookupSupplier(
      Supplier<SymbolLookup> lookupFactory) {
    Objects.requireNonNull(lookupFactory, "lookupFactory must not be null");
    return () -> libraryLookup(lookupFactory);
  }

  static Optional<SymbolLookup> namedLibraryLookup(String libraryName) {
    Objects.requireNonNull(libraryName, "libraryName must not be null");
    return libraryLookup(() -> SymbolLookup.libraryLookup(libraryName, Arena.global()));
  }

  static Optional<SymbolLookup> kernel32LibraryLookup() {
    return kernel32LibraryLookup(System.getProperty("os.name", ""), "Kernel32");
  }

  static Optional<SymbolLookup> kernel32LibraryLookup(String osName, String libraryName) {
    Objects.requireNonNull(osName, "osName must not be null");
    Objects.requireNonNull(libraryName, "libraryName must not be null");
    if (!isWindows(osName)) {
      return Optional.empty();
    }
    return namedLibraryLookup(libraryName);
  }

  static DowncallHandleProvider downcallHandleProvider(Linker linker) {
    Objects.requireNonNull(linker, "linker must not be null");
    return linker::downcallHandle;
  }

  static NativeStandardInputProbe windowsCurrentProcessFromLookup(
      SymbolLookup symbolLookup, DowncallHandleProvider downcallHandleProvider) {
    if (symbolLookup == null) {
      return unsupported();
    }
    return currentProcessProbe(
        () ->
            windowsCurrentProcess(
                (SymbolResolver) name -> findOrThrow(symbolLookup, name), downcallHandleProvider));
  }

  static NativeStandardInputProbe windowsCurrentProcess(
      SymbolResolver symbolResolver, DowncallHandleProvider downcallHandleProvider) {
    MethodHandle getStdHandle =
        downcallHandleProvider.downcallHandle(
            symbolResolver.findOrThrow("GetStdHandle"),
            FunctionDescriptor.of(ValueLayout.ADDRESS, ValueLayout.JAVA_INT));
    MethodHandle getConsoleMode =
        downcallHandleProvider.downcallHandle(
            symbolResolver.findOrThrow("GetConsoleMode"),
            FunctionDescriptor.of(ValueLayout.JAVA_INT, ValueLayout.ADDRESS, ValueLayout.ADDRESS));
    return of(
        () -> {
          MemorySegment handle = (MemorySegment) getStdHandle.invokeExact(STD_INPUT_HANDLE);
          if (handle.equals(MemorySegment.NULL)) {
            return false;
          }
          try (Arena arena = Arena.ofConfined()) {
            MemorySegment mode = arena.allocate(ValueLayout.JAVA_INT);
            return ((int) getConsoleMode.invokeExact(handle, mode)) != 0;
          }
        });
  }

  private NativeStandardInputProbe(CheckedBooleanSupplier probe) {
    this.probe = Objects.requireNonNull(probe, "probe must not be null");
  }

  @Override
  public boolean getAsBoolean() {
    try {
      return probe.getAsBoolean();
    } catch (Throwable throwable) {
      return false;
    }
  }

  private static MemorySegment findOrThrow(SymbolLookup symbolLookup, String name) {
    Optional<MemorySegment> symbol = symbolLookup.find(name);
    if (symbol.isPresent()) {
      return symbol.get();
    }
    throw new IllegalArgumentException("missing native symbol: " + name);
  }
}
