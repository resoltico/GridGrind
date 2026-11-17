package dev.erst.gridgrind.engine.media;

import dev.erst.gridgrind.core.limits.Limit;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import java.util.Optional;

/** One invocation's monotonic media-use and structural-work counters, shared by nested readers. */
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  Limit.ID_AUTHORED_MEDIA_BYTES,
  Limit.ID_AUTHORED_MEDIA_TOTAL_BYTES,
  Limit.ID_MEDIA_INSPECTION_BYTES,
  Limit.ID_MEDIA_RASTER_PIXELS,
  Limit.ID_MEDIA_EXPANDED_BYTES,
  Limit.ID_MEDIA_STRUCTURE_RECORDS,
  Limit.ID_MEDIA_CONTAINER_DEPTH,
  Limit.ID_EPS_SYNTAX_TOKENS,
  Limit.ID_EPS_SYNTAX_DEPTH
})
public final class MediaAdmissionBudget {
  private long mediaUse;
  private long inspected;
  private long pixels;
  private long expanded;
  private long records;
  private long tokens;
  private Optional<String> blocked = Optional.empty();

  /** Remaining decoded-byte allowance for one media occurrence, before acquisition. */
  public long payloadAllowance() {
    available();
    long remaining = Limit.MAX_AUTHORED_MEDIA_TOTAL_BYTES - mediaUse;
    if (remaining == 0) throw block(Limit.ID_AUTHORED_MEDIA_TOTAL_BYTES);
    return Math.min(Limit.MAX_AUTHORED_MEDIA_BYTES, remaining);
  }

  /** Identifies the currently binding occurrence or aggregate payload limit. */
  public String payloadLimitId() {
    return Limit.MAX_AUTHORED_MEDIA_TOTAL_BYTES - mediaUse < Limit.MAX_AUTHORED_MEDIA_BYTES
        ? Limit.ID_AUTHORED_MEDIA_TOTAL_BYTES
        : Limit.ID_AUTHORED_MEDIA_BYTES;
  }

  /** Checks the occurrence size before copying materialized immutable bytes. */
  public void requirePayloadSize(long bytes) {
    available();
    if (bytes < 0) throw new IllegalStateException("Negative media size");
    if (bytes > Limit.MAX_AUTHORED_MEDIA_BYTES) throw block(Limit.ID_AUTHORED_MEDIA_BYTES);
  }

  /** Charges one authored occurrence, including repeated equal or identical sources. */
  public void reserveOccurrence(long bytes) {
    requirePayloadSize(bytes);
    mediaUse =
        charge(
            mediaUse,
            bytes,
            Limit.MAX_AUTHORED_MEDIA_TOTAL_BYTES,
            Limit.ID_AUTHORED_MEDIA_TOTAL_BYTES);
  }

  void inspect(long bytes) {
    inspected =
        charge(inspected, bytes, Limit.MAX_MEDIA_INSPECTION_BYTES, Limit.ID_MEDIA_INSPECTION_BYTES);
  }

  void record() {
    records =
        charge(records, 1, Limit.MAX_MEDIA_STRUCTURE_RECORDS, Limit.ID_MEDIA_STRUCTURE_RECORDS);
  }

  void raster(long width, long height) {
    if (width <= 0 || height <= 0)
      throw new IllegalStateException("Raster dimensions must be checked at their format boundary");
    if (width > Limit.MAX_MEDIA_RASTER_PIXELS / height) throw block(Limit.ID_MEDIA_RASTER_PIXELS);
    pixels =
        charge(pixels, width * height, Limit.MAX_MEDIA_RASTER_PIXELS, Limit.ID_MEDIA_RASTER_PIXELS);
  }

  long remainingExpansion() {
    available();
    return Limit.MAX_MEDIA_EXPANDED_BYTES - expanded;
  }

  void expand(long bytes) {
    expanded =
        charge(expanded, bytes, Limit.MAX_MEDIA_EXPANDED_BYTES, Limit.ID_MEDIA_EXPANDED_BYTES);
  }

  void token() {
    tokens = charge(tokens, 1, Limit.MAX_EPS_SYNTAX_TOKENS, Limit.ID_EPS_SYNTAX_TOKENS);
  }

  void depth(int depth) {
    available();
    if (depth > Limit.MAX_MEDIA_CONTAINER_DEPTH) throw block(Limit.ID_MEDIA_CONTAINER_DEPTH);
  }

  void syntaxDepth(int depth) {
    available();
    if (depth > Limit.MAX_EPS_SYNTAX_DEPTH) throw block(Limit.ID_EPS_SYNTAX_DEPTH);
  }

  /** Retains the binding resource refusal after failed partial acquisition. */
  public void poison(String id) {
    blocked = Optional.of(id);
  }

  private long charge(long previous, long amount, long maximum, String id) {
    available();
    if (amount < 0) throw new IllegalStateException("Negative media charge");
    if (amount > maximum - previous) throw block(id);
    return previous + amount;
  }

  private InvalidRequestException block(String id) {
    poison(id);
    return MediaIssues.limit(id);
  }

  private void available() {
    if (blocked.isPresent()) throw MediaIssues.limit(blocked.orElseThrow());
  }

  /** Returns media-use bytes charged even when payload identity repeats. */
  public long mediaUseBytes() {
    return mediaUse;
  }
}
