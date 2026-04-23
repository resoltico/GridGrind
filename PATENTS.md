---
afad: "3.5"
version: "0.55.0"
domain: LEGAL
updated: "2026-04-19"
route:
  keywords: [gridgrind, patent, patents, legal, contributors, dependencies, non-assertion, mit, apache, bsd]
  questions: ["what is the patent posture for gridgrind", "does gridgrind include a patent non-assertion statement", "what do contributors need to know about patents"]
---

# Patent Considerations

GridGrind is licensed under the MIT License, which does not include explicit patent grant language.

## Summary

GridGrind makes no patent claims and knowingly infringes no patents.

If explicit patent protection is a concern for your use case, consult qualified legal counsel.

## GridGrind (MIT License)

The MIT License grants broad permissions to use, copy, modify, merge, publish, distribute, and
sublicense the software without restriction. It does not include an explicit patent grant or patent
retaliation clause.

| Component | License | Explicit Patent Grant |
|:----------|:--------|:----------------------|
| GridGrind | MIT | No (implicit only) |
| Apache POI | Apache 2.0 | Yes |
| Apache XMLBeans | Apache 2.0 | Yes |
| Apache Log4j Core / API | Apache 2.0 | Yes |
| Apache Santuario xmlsec | Apache 2.0 | Yes |
| Apache Commons (Codec, Collections, Compress, IO, Lang, Math) | Apache 2.0 | Yes |
| Jackson Databind / Core / Annotations | Apache 2.0 | Yes |
| SparseBitSet | Apache 2.0 | Yes |
| Bouncy Castle (bcpkix, bcprov, bcutil) | MIT variant | No |
| SLF4J API | MIT | No |
| CurvesAPI | BSD 3-Clause | No |

## Bundled Dependencies

GridGrind's executable JAR bundles Apache POI, Apache XMLBeans, Apache Log4j, Apache Santuario
xmlsec, Apache Commons libraries, Jackson, SparseBitSet, Bouncy Castle, SLF4J API, and
CurvesAPI. See NOTICE for the complete list and attributions.

The Apache 2.0 components include an explicit patent grant in Section 3:

> Subject to the terms and conditions of this License, each Contributor hereby grants to You a
> perpetual, worldwide, non-exclusive, no-charge, royalty-free, irrevocable patent license to
> make, have made, use, offer to sell, sell, import, and otherwise transfer the Work...

This Apache 2.0 patent grant applies to contributions made by each respective project's
contributors. It does not extend to GridGrind's own code, which is licensed separately under MIT.

CurvesAPI is licensed under the BSD 3-Clause License, which does not include an explicit patent
grant.

## What This Means

GridGrind is an independent implementation that uses Apache POI as a runtime library. It does not
fork, modify, or derive algorithms from Apache POI's internal implementation. GridGrind's code is
original work.

For GridGrind users:

1. The Apache 2.0 patent grant from Apache POI, XMLBeans, Log4j, Commons, Jackson, and
   SparseBitSet contributors applies to those bundled components.
2. CurvesAPI (BSD 3-Clause) and GridGrind's own MIT-licensed code do not carry explicit patent
   grants.
3. The copyright holder (Ervins Strauhmanis) makes no patent claims on GridGrind's implementation
   and is not aware of any patents this implementation infringes.

## Patent Non-Assertion

The copyright holder states:

GridGrind makes no patent claims. If the copyright holder holds any patents that relate to this
implementation, permission is granted under the MIT License to use this implementation without
patent liability.

This is a statement of intent, not a formal legal instrument.

## For Contributors

By contributing to GridGrind, you grant MIT License permissions for your contributions. The MIT
License does not include an explicit patent grant. Do not contribute code you know infringes
patents you or others hold.

## Legal Disclaimer

This document is informational only and does not constitute legal advice. Patent law is complex
and jurisdiction-specific. For patent-related concerns, consult qualified legal counsel.
