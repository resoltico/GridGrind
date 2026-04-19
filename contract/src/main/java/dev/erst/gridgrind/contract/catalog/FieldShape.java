package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** JSON-serializable recursive field-shape contract used by protocol discovery output. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
@JsonSubTypes({
  @JsonSubTypes.Type(value = FieldShape.Scalar.class, name = "SCALAR"),
  @JsonSubTypes.Type(value = FieldShape.ListShape.class, name = "LIST"),
  @JsonSubTypes.Type(value = FieldShape.TopLevelTypeSetRef.class, name = "TOP_LEVEL_TYPE_SET"),
  @JsonSubTypes.Type(value = FieldShape.NestedTypeGroupRef.class, name = "NESTED_TYPE_GROUP"),
  @JsonSubTypes.Type(
      value = FieldShape.NestedTypeGroupUnionRef.class,
      name = "NESTED_TYPE_GROUP_UNION"),
  @JsonSubTypes.Type(value = FieldShape.PlainTypeGroupRef.class, name = "PLAIN_TYPE_GROUP")
})
public sealed interface FieldShape
    permits FieldShape.Scalar,
        FieldShape.ListShape,
        FieldShape.TopLevelTypeSetRef,
        FieldShape.NestedTypeGroupRef,
        FieldShape.NestedTypeGroupUnionRef,
        FieldShape.PlainTypeGroupRef {

  /** Scalar JSON value type such as STRING, NUMBER, or BOOLEAN. */
  record Scalar(ScalarType scalarType) implements FieldShape {
    public Scalar {
      Objects.requireNonNull(scalarType, "scalarType must not be null");
    }
  }

  /** Recursive array shape whose elements all share the same shape. */
  record ListShape(FieldShape elementShape) implements FieldShape {
    public ListShape {
      Objects.requireNonNull(elementShape, "elementShape must not be null");
    }
  }

  /** Reference to one top-level discriminated family published elsewhere in the catalog. */
  record TopLevelTypeSetRef(String typeSet) implements FieldShape {
    public TopLevelTypeSetRef {
      Objects.requireNonNull(typeSet, "typeSet must not be null");
      if (typeSet.isBlank()) {
        throw new IllegalArgumentException("typeSet must not be blank");
      }
    }
  }

  /** Reference to one nested discriminated union group published elsewhere in the catalog. */
  record NestedTypeGroupRef(String group) implements FieldShape {
    public NestedTypeGroupRef {
      Objects.requireNonNull(group, "group must not be null");
      if (group.isBlank()) {
        throw new IllegalArgumentException("group must not be blank");
      }
    }
  }

  /** Reference to one union of nested discriminated groups published elsewhere in the catalog. */
  record NestedTypeGroupUnionRef(java.util.List<String> groups) implements FieldShape {
    public NestedTypeGroupUnionRef {
      groups = java.util.List.copyOf(groups);
      if (groups.isEmpty()) {
        throw new IllegalArgumentException("groups must not be empty");
      }
      for (String group : groups) {
        Objects.requireNonNull(group, "group must not be null");
        if (group.isBlank()) {
          throw new IllegalArgumentException("group must not be blank");
        }
      }
    }
  }

  /** Reference to one plain record-type group published elsewhere in the catalog. */
  record PlainTypeGroupRef(String group) implements FieldShape {
    public PlainTypeGroupRef {
      Objects.requireNonNull(group, "group must not be null");
      if (group.isBlank()) {
        throw new IllegalArgumentException("group must not be blank");
      }
    }
  }
}
