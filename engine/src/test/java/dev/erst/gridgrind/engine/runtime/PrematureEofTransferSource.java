package dev.erst.gridgrind.engine.runtime;

import java.nio.ByteBuffer;
import java.nio.channels.WritableByteChannel;

/** Declares one byte but reaches EOF before a fallback transfer can read it. */
final class PrematureEofTransferSource extends RequestPathPublicationFileSupport.TransferSource {
  @Override
  long size() {
    return 1;
  }

  @Override
  long transferTo(long position, long count, WritableByteChannel target) {
    return 0;
  }

  @Override
  void position(long position) {}

  @Override
  int read(ByteBuffer destination) {
    return -1;
  }
}
