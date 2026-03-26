package com.code_intelligence.jazzer.utils;

import java.lang.reflect.Field;
import java.util.Arrays;
import sun.misc.Unsafe;

/** Provides access to {@link Unsafe} for Jazzer classes that expect the upstream helper type. */
public final class UnsafeProvider {
  private static final Unsafe UNSAFE = getUnsafeInternal();

  private UnsafeProvider() {}

  /** Returns the process-wide {@link Unsafe} instance. */
  public static Unsafe getUnsafe() {
    return UNSAFE;
  }

  private static Unsafe getUnsafeInternal() {
    try {
      return Unsafe.getUnsafe();
    } catch (Throwable unused) {
      for (Field field : Unsafe.class.getDeclaredFields()) {
        if (field.getType() == Unsafe.class) {
          field.setAccessible(true);
          try {
            return (Unsafe) field.get(null);
          } catch (IllegalAccessException exception) {
            throw new IllegalStateException(
                "Failed to access Unsafe via reflection for the local Jazzer replay layer.",
                exception);
          }
        }
      }
      throw new IllegalStateException(
          "Failed to locate an Unsafe field on sun.misc.Unsafe: "
              + Arrays.deepToString(Unsafe.class.getDeclaredFields()));
    }
  }
}
