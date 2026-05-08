package dev.erst.gridgrind.cli;

import java.util.Objects;
import java.util.function.BooleanSupplier;
import java.util.function.Supplier;

/** Lazily resolves one BooleanSupplier delegate and reuses it for later calls. */
final class LazyBooleanSupplier implements BooleanSupplier {
  private BooleanSupplier delegate;

  LazyBooleanSupplier(Supplier<BooleanSupplier> delegateFactory) {
    Supplier<BooleanSupplier> requiredFactory =
        Objects.requireNonNull(delegateFactory, "delegateFactory must not be null");
    this.delegate =
        () -> {
          BooleanSupplier resolved =
              Objects.requireNonNull(
                  requiredFactory.get(), "delegateFactory must not return a null supplier");
          delegate = resolved;
          return resolved.getAsBoolean();
        };
  }

  @Override
  public boolean getAsBoolean() {
    return delegate.getAsBoolean();
  }
}
