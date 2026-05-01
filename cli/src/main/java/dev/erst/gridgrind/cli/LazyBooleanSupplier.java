package dev.erst.gridgrind.cli;

import java.util.Objects;
import java.util.function.BooleanSupplier;
import java.util.function.Supplier;

/** Lazily resolves one BooleanSupplier delegate and reuses it for later calls. */
final class LazyBooleanSupplier implements BooleanSupplier {
  private final Supplier<BooleanSupplier> delegateFactory;
  private BooleanSupplier delegate;

  LazyBooleanSupplier(Supplier<BooleanSupplier> delegateFactory) {
    this.delegateFactory =
        Objects.requireNonNull(delegateFactory, "delegateFactory must not be null");
  }

  @Override
  public boolean getAsBoolean() {
    if (delegate == null) {
      delegate =
          Objects.requireNonNull(
              delegateFactory.get(), "delegateFactory must not return a null supplier");
    }
    return delegate.getAsBoolean();
  }
}
