---
afad: "5.0.1"
version: "0.75.0"
domain: LEGAL
updated: "2026-08-31"
route:
  keywords: [gridgrind, patent, patents, legal, dependencies, mit, apache, bsd, container]
  questions: ["what patent rights accompany gridgrind", "does gridgrind offer a patent non-assertion pledge", "what patent terms apply to bundled dependencies"]
---

# Patent Considerations

This document describes which distributed license texts contain an express patent grant. It is
informational only, is not a patent search or legal opinion, and does not add to or modify any
license.

## GridGrind Code

GridGrind's own code is distributed under the MIT License in [LICENSE](LICENSE). That license
contains broad copyright permissions but no express patent-license clause or patent-retaliation
clause.

No patent search, freedom-to-operate analysis, patent-clearance opinion, or non-infringement
determination has been performed or is represented by this repository. Absence of a known patent
claim is not a representation that no relevant patent exists.

GridGrind does not provide a separate patent covenant, patent non-assertion pledge, contributor
patent agreement, or express patent grant for its own code. If a formal patent grant is needed,
obtain qualified legal advice rather than relying on repository prose.

## Bundled JAR Components

The executable JAR combines separately licensed components. [NOTICE](NOTICE) identifies the
resolved component families and preserves their attribution notices.

| Component family | Distributed terms | Express patent clause |
|:--|:--|:--|
| GridGrind | MIT | No |
| Apache POI, XMLBeans, Log4j, Santuario, and Commons | Apache-2.0 | Yes, Apache License 2.0 Section 3 |
| Apache POI Custom XML recipe workbook | Apache-2.0 | Yes, Apache License 2.0 Section 3 |
| Jackson, SparseBitSet, and Woodstox Core | Apache-2.0 | Yes, Apache License 2.0 Section 3 |
| FastDoubleParser and Schubfach code shaded into Jackson Core | MIT, BSL-1.0, and BSD-2-Clause | No |
| Bouncy Castle and SLF4J API | MIT-style / MIT | No |
| Stax2 API | BSD-2-Clause | No |
| CurvesAPI | BSD-3-Clause | No |
| Jakarta Activation API and Jakarta XML Binding API | Eclipse Distribution License v1.0 (SPDX BSD-3-Clause) | No |

The Apache License patent grant is component-specific. It applies only to patent claims licensable
by each Apache-licensed component's contributors that are necessarily infringed by their
contributions as described in Section 3. It does not create a patent grant for GridGrind's MIT-
licensed code or for components under other licenses. The termination condition in Apache License
2.0 Section 3 also remains part of those component licenses.

MIT and BSD licenses in this distribution do not contain an express patent-license clause. This
document does not characterize whether any implied patent rights might arise under a particular
jurisdiction.

## Source Distribution

The public source tree also includes the Gradle wrapper under Apache-2.0. The component-specific
Section 3 patent grant and termination terms described above apply to that artifact; they do not
extend to GridGrind's MIT-licensed code. The Apache POI Custom XML workbook listed in the JAR table
is also present as a source fixture. [NOTICE](NOTICE) records the exact provenance and distribution
scope of both artifacts.

## Container Image

The published container image includes the GridGrind JAR, an Azul Zulu build of OpenJDK, and
operating-system and font packages. Those additional components retain their own license and patent
terms; they are not covered by GridGrind's MIT license or by one combined patent statement. Their
legal materials remain in the image's JRE legal directory and package documentation directories,
and published images include an SBOM attestation for content inventory.

## Contributors

The repository currently documents no contributor license agreement, developer certificate of
origin, or separate contributor patent grant. Contributors must have the right to submit their
work and should not submit material they know they lack permission to contribute. The legal effect
of a contribution depends on the applicable facts and law; this document does not determine it.

## Cryptography and Jurisdiction

GridGrind includes cryptographic functionality for OOXML encryption, signing, and XML security.
Patent, import, export, and cryptography rules vary by jurisdiction. Users and redistributors are
responsible for determining which requirements apply to their use or distribution.

## Disclaimer

This document is not legal advice. Consult qualified counsel for patent clearance, licensing,
redistribution, contribution, or jurisdiction-specific questions.
