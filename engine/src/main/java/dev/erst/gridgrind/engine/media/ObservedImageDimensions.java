package dev.erst.gridgrind.engine.media;

import dev.erst.gridgrind.core.problem.FieldValidationMediaRule;
import dev.erst.gridgrind.core.problem.FieldValidationProblem;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.problem.MediaStructureReason;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.engine.backend.xssf.UnexpectedNativeContentException;
import java.util.Optional;

/** Bounded factual header observation, independent of authored admission and raster allocation. */
public final class ObservedImageDimensions {
  private ObservedImageDimensions() {}

  /** Factual first-image dimensions when representable at the report boundary. */
  public record Dimensions(int width, int height) {}

  /** Observes supported raster headers; inability leaves dimensions absent and bytes untouched. */
  public static Optional<Dimensions> read(byte[] bytes) {
    if (bytes.length < 4) return Optional.empty();
    var input =
        new MediaInput(
            bytes, 0, bytes.length, ExcelPictureFormat.PNG, new MediaAdmissionBudget(), 0);
    try {
      var geometry = geometry(input);
      long width = geometry.width();
      long height = geometry.height();
      return width > 0 && height > 0 && width <= Integer.MAX_VALUE && height <= Integer.MAX_VALUE
          ? Optional.of(new Dimensions((int) width, (int) height))
          : Optional.empty();
    } catch (InvalidRequestException expected) {
      if (expected.requestProblem() instanceof FieldValidationProblem field
          && field.rule() instanceof FieldValidationMediaRule) return Optional.empty();
      throw new UnexpectedNativeContentException(expected);
    } catch (UnexpectedNativeContentException classified) {
      throw classified;
    } catch (RuntimeException unexpected) {
      throw new UnexpectedNativeContentException(unexpected);
    }
  }

  private static ImageGeometry geometry(MediaInput in) {
    int first = in.u8();
    int second = in.u8();
    if (first == 137 && second == 80) {
      in.magic(78, 71, 13, 10, 26, 10);
      in.require(in.u32() == 13, MediaStructureReason.HEADER);
      in.magic(73, 72, 68, 82);
      return new ImageGeometry(in.u32(), in.u32());
    }
    if (first == 71 && second == 73) {
      in.magic(70, 56);
      int version = in.u8();
      in.require((version == 55 || version == 57) && in.u8() == 97, MediaStructureReason.HEADER);
      in.little = true;
      return new ImageGeometry(in.u16(), in.u16());
    }
    if (first == 66 && second == 77) return bitmap(in);
    if ((first == 73 && second == 73) || (first == 77 && second == 77)) {
      in.little = first == 73;
      return tiff(in);
    }
    if (first == 255 && second == 216) return jpeg(in);
    return ImageGeometry.vector();
  }

  private static ImageGeometry bitmap(MediaInput in) {
    in.little = true;
    in.seek(14);
    long header = in.u32();
    if (header == 12) return new ImageGeometry(in.u16(), in.u16());
    if (header >= 40) return new ImageGeometry(in.s32(), Math.abs((long) in.s32()));
    return ImageGeometry.vector();
  }

  private static ImageGeometry tiff(MediaInput in) {
    boolean big = tiffVersion(in);
    long directory = unsigned(in, big);
    var entries = in.region(directory, in.size() - directory);
    long count = big ? entries.u64() : entries.u16();
    in.require(count <= entries.remaining() / (big ? 20 : 12), MediaStructureReason.COUNT);
    long width = 0;
    long height = 0;
    for (long i = 0; i < count && (width == 0 || height == 0); i++) {
      in.record();
      int tag = entries.u16();
      int type = entries.u16();
      long size = unsigned(entries, big);
      var slot = entries.take(big ? 8 : 4);
      if ((tag != 256 && tag != 257) || size != 1) continue;
      long value = scalar(slot, type, big);
      if (tag == 256) width = value;
      else height = value;
    }
    return new ImageGeometry(width, height);
  }

  private static boolean tiffVersion(MediaInput in) {
    long version = in.u16();
    boolean big = version == 43;
    in.require(version == 42 || big, MediaStructureReason.HEADER);
    if (big) in.require(in.u16() == 8 && in.u16() == 0, MediaStructureReason.HEADER);
    return big;
  }

  private static long unsigned(MediaInput in, boolean big) {
    return big ? in.u64() : in.u32();
  }

  private static long scalar(MediaInput slot, int type, boolean big) {
    return switch (type) {
      case 3 -> slot.u16();
      case 4 -> slot.u32();
      case 16 -> big ? slot.u64() : 0;
      default -> 0;
    };
  }

  private static ImageGeometry jpeg(MediaInput in) {
    while (in.remaining() > 0) {
      in.record();
      in.require(in.u8() == 255, MediaStructureReason.HEADER);
      int marker;
      do {
        marker = in.u8();
      } while (marker == 255);
      if (marker == 217 || marker == 218) break;
      if (marker == 1 || marker >= 208 && marker <= 215) continue;
      int length = in.u16();
      in.require(length >= 2, MediaStructureReason.LENGTH);
      var data = in.take(length - 2);
      if (marker >= 192 && marker <= 207 && marker != 196 && marker != 200 && marker != 204) {
        data.u8();
        long height = data.u16();
        return new ImageGeometry(data.u16(), height);
      }
    }
    return ImageGeometry.vector();
  }
}
