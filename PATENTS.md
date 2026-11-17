# Patent Considerations

This document describes which distributed license texts contain an express patent grant. It is
informational only, is not a patent search or legal opinion, and does not add to or modify any
license.

## GridGrind Code

GridGrind's original code uses GNU Affero General Public License version 3 only
(`AGPL-3.0-only`), as identified in [PROJECT-NOTICE.md](PROJECT-NOTICE.md) and
[LICENSE](LICENSE). Section 11 describes contributors' essential patent claims and the
non-exclusive, worldwide, royalty-free patent license specified there. Its conditions and
limitations remain those of the license. This document supplies no separate patent grant.

No patent search, freedom-to-operate analysis, patent-clearance opinion, or non-infringement
determination has been performed or is represented by this repository. Absence of a known patent
claim is not a representation that no relevant patent exists.

GridGrind does not provide a separate patent covenant, patent non-assertion pledge, or contributor
patent agreement beyond the rights granted by its license.

## Bundled JAR Components

The executable JAR combines separately licensed components. [NOTICE](NOTICE) identifies the
resolved component families and preserves their attribution notices.

| Component family | Distributed terms | Express patent clause |
|:--|:--|:--|
| GridGrind original code | AGPL-3.0-only | Yes, AGPLv3 Section 11, with its stated scope and conditions |
| Apache POI, XMLBeans, Log4j, Santuario, and Commons | Apache-2.0 | Yes, Apache License 2.0 Section 3 |
| Apache POI Custom XML recipe workbook | Apache-2.0 | Yes, Apache License 2.0 Section 3 |
| Jackson, SparseBitSet, and Woodstox Core project code | Apache-2.0 | Yes, Apache License 2.0 Section 3 |
| MSV/xsdlib and RELAX NG code embedded by Woodstox | Component-specific BSD-3-Clause, Apache-1.1 and Apache-2.0 | Section 3 applies to Apache-2.0 portions; no express patent clause in the BSD/Apache-1.1 portions |
| ISO-RELAX code embedded by Woodstox | MIT-style | No express patent clause |
| FastDoubleParser and Schubfach code shaded into Jackson Core | MIT, BSL-1.0, and BSD-2-Clause | No |
| Bouncy Castle and SLF4J API | MIT-style / MIT | No |
| Stax2 API | BSD-2-Clause | No |
| CurvesAPI | BSD-3-Clause | No |
| Jakarta Activation API and Jakarta XML Binding API | Eclipse Distribution License v1.0 (SPDX BSD-3-Clause) | No |

The Apache License patent grant is component-specific. It applies only to patent claims licensable
by each Apache-licensed component's contributors that are necessarily infringed by their
contributions as described in Section 3. It does not create a patent grant for GridGrind's AGPL-licensed original
code or for components under other licenses. The termination condition in Apache License 2.0
Section 3 also remains part of those component licenses.

MIT and BSD licenses in this distribution do not contain an express patent-license clause. This
document does not characterize whether any implied patent rights might arise under a particular
jurisdiction.

## Source Distribution

The public source tree also includes the Gradle wrapper under Apache-2.0. The component-specific
Section 3 patent grant and termination terms described above apply to that artifact; they do not
extend to GridGrind's AGPL-licensed original code. The Apache POI Custom XML workbook listed in the JAR table
is also present as a source fixture. [NOTICE](NOTICE) records the exact provenance and distribution
scope of both artifacts.

## Container Image

The container contains GridGrind, the pinned Eclipse Temurin OpenJDK/Ubuntu environment and
separately licensed operating-system/font packages. Those materials retain their own licenses
and patent terms. Their vendor legal files and package copyright files remain in the image;
[NOTICE](NOTICE) and the source manifest identify their distribution/source scope. No single
project license or patent statement replaces those terms.

## Contributors

[Contribution terms](PROJECT-NOTICE.md#contribution-terms) retain contributors' copyright and
grant the public AGPL license plus the stated perpetual, worldwide, non-exclusive,
royalty-free, irrevocable owner license to use, reproduce, modify, distribute, sublicense and
license contributions under other terms. These terms do not state an assignment, DCO or a
separate express contributor patent/trademark grant. The AGPL's own patent provisions remain
applicable according to their terms. This is not a representation of patent clearance.

## Cryptography and Jurisdiction

GridGrind includes cryptographic functionality for OOXML encryption, signing, and XML security.
Patent, import, export, and cryptography rules vary by jurisdiction. Users and redistributors are
responsible for determining which requirements apply to their use or distribution.

## Disclaimer

This document is not legal advice. Consult qualified counsel for patent clearance, licensing,
redistribution, contribution, or jurisdiction-specific questions.
