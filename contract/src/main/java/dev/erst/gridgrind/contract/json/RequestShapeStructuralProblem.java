package dev.erst.gridgrind.contract.json;

/** A structural defect in an otherwise syntactically valid request value or creator contract. */
sealed interface RequestShapeStructuralProblem extends RequestStructuralProblem
    permits RequestUnknownField,
        RequestMissingRequiredField,
        RequestExplicitNullField,
        RequestMissingTypeDiscriminator,
        RequestUnknownTypeDiscriminator,
        RequestUnsupportedEnumValue,
        RequestMalformedScalar {}
