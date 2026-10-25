package dev.erst.gridgrind.contract.json;

import com.fasterxml.jackson.annotation.JsonIgnore;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolContractSupport;
import java.lang.reflect.RecordComponent;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;

/** Provides creator-contract field and member lookup for tolerant JSON objects. */
final class RequestObjectMembers {
  private RequestObjectMembers() {}

  /**
   * Indexes one raw object once so structural collection preserves source-order member ownership.
   */
  static Index index(RequestJsonObject object) {
    return new Index(object);
  }

  static void collectFields(
      Index members,
      String jsonPath,
      List<String> requiredFields,
      List<String> optionalFields,
      List<RequestStructuralProblem> problems) {
    Objects.requireNonNull(members, "members must not be null");
    collectMissingRequiredFields(members, jsonPath, requiredFields, problems);
    collectUnknownFields(
        members,
        jsonPath,
        java.util.stream.Stream.concat(requiredFields.stream(), optionalFields.stream())
            .collect(java.util.stream.Collectors.toUnmodifiableSet()),
        problems);
  }

  static void collectMissingRequiredFields(
      Index members,
      String jsonPath,
      List<String> requiredFields,
      List<RequestStructuralProblem> problems) {
    Objects.requireNonNull(members, "members must not be null");
    for (String required : requiredFields) {
      if (members.member(required).isEmpty()) {
        problems.add(new RequestMissingRequiredField(childPath(jsonPath, required)));
      }
    }
  }

  static void collectUnknownFields(
      Index members,
      String jsonPath,
      Set<String> allowedFields,
      List<RequestStructuralProblem> problems) {
    Objects.requireNonNull(members, "members must not be null");
    Objects.requireNonNull(allowedFields, "allowedFields must not be null");
    for (RequestJsonMember member : members.members()) {
      if (!allowedFields.contains(member.name())) {
        problems.add(
            new RequestUnknownField(childPath(jsonPath, member.name()), member.nameByteOffset()));
      }
    }
  }

  static List<RecordComponent> visibleRecordComponents(Class<? extends Record> recordType) {
    List<RecordComponent> components = new ArrayList<>();
    for (RecordComponent component : recordType.getRecordComponents()) {
      if (isVisible(component)) {
        components.add(component);
      }
    }
    return List.copyOf(components);
  }

  static List<String> optionalFieldNames(
      List<RecordComponent> components, List<String> requiredFields) {
    return components.stream()
        .map(GridGrindProtocolContractSupport::wireFieldName)
        .filter(field -> !requiredFields.contains(field))
        .toList();
  }

  static String childPath(String parent, String child) {
    return parent.isEmpty() ? child : parent + "." + child;
  }

  /** Immutable, source-ordered lookup for every occurrence in one tolerant JSON object. */
  static final class Index {
    private final List<RequestJsonMember> members;
    private final Map<String, RequestJsonMember> firstMembers;
    private final Map<String, List<RequestJsonMember>> membersByName;

    @SuppressWarnings(
        "PMD.UseConcurrentHashMap") // Thread-confined insertion order preserves authored order.
    private Index(RequestJsonObject object) {
      Objects.requireNonNull(object, "object must not be null");
      members = object.members();
      Map<String, RequestJsonMember> first = new LinkedHashMap<>();
      Map<String, List<RequestJsonMember>> named = new LinkedHashMap<>();
      for (RequestJsonMember member : members) {
        first.putIfAbsent(member.name(), member);
        addNamedMember(named, member);
      }
      firstMembers = Collections.unmodifiableMap(first);
      Map<String, List<RequestJsonMember>> immutableNamed = new LinkedHashMap<>();
      named.forEach((name, occurrences) -> immutableNamed.put(name, List.copyOf(occurrences)));
      membersByName = Collections.unmodifiableMap(immutableNamed);
    }

    /** Returns the deterministic first occurrence selected for a structurally valid field. */
    Optional<RequestJsonMember> member(String name) {
      Objects.requireNonNull(name, "name must not be null");
      return Optional.ofNullable(firstMembers.get(name));
    }

    /** Returns every authored occurrence in source order, including duplicate keys. */
    List<RequestJsonMember> membersNamed(String name) {
      Objects.requireNonNull(name, "name must not be null");
      return membersByName.getOrDefault(name, List.of());
    }

    /** Returns every object member in the original source order. */
    List<RequestJsonMember> members() {
      return members;
    }

    private static void addNamedMember(
        Map<String, List<RequestJsonMember>> named, RequestJsonMember member) {
      named.computeIfAbsent(member.name(), _ -> new ArrayList<>()).add(member);
    }
  }

  private static boolean isVisible(RecordComponent component) {
    return !component.getAccessor().isAnnotationPresent(JsonIgnore.class);
  }
}
