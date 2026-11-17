/* SPDX-License-Identifier: AGPL-3.0-only */
package org.apache.poi.ss.usermodel;

import java.util.function.Supplier;

/** Operation-owned native date system; limits come from the explicit core context. */
public final class DateSystemContext {
  private static final ScopedValue<Boolean> EPOCH = ScopedValue.newInstance();

  private DateSystemContext() {}

  public static boolean is1904() {
    return EPOCH.isBound() && EPOCH.get();
  }

  private static double windowsMaximum, macMaximum;
  private static boolean configured;

  public static void configure(double windows, double mac) {
    if (!Double.isFinite(windows) || !Double.isFinite(mac) || windows < 0 || mac < 0)
      throw new IllegalArgumentException("Invalid native date bounds");
    if (configured && (windows != windowsMaximum || mac != macMaximum))
      throw new IllegalStateException("Native realm policy changed");
    windowsMaximum = windows;
    macMaximum = mac;
    configured = true;
  }

  public static double maximumSerial() {
    if (!configured) throw new IllegalStateException("Explicit core date bounds are required");
    return is1904() ? macMaximum : windowsMaximum;
  }

  public static <T> T call(boolean epoch, Supplier<T> action) {
    return ScopedValue.where(EPOCH, epoch).call(action::get);
  }
}
