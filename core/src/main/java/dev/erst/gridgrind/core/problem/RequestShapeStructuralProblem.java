package dev.erst.gridgrind.core.problem;

/** A structural defect in an otherwise syntactically valid request value or creator contract. */
public sealed interface RequestShapeStructuralProblem extends RequestStructuralProblem
    permits RequestUnknownField,
        RequestMissingRequiredField,
        RequestExplicitNullField,
        RequestMissingTypeDiscriminator,
        RequestUnknownTypeDiscriminator,
        RequestUnsupportedEnumValue,
        RequestMalformedScalar {}
