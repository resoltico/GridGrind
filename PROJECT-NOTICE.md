# GridGrind licensing and source

Copyright © 2026 Ervins Strauhmanis. Contributors retain copyright in their contributions.

Original GridGrind source, tests, tooling and accompanying documentation in this checkout are
licensed under the GNU Affero General Public License, version 3 only (`AGPL-3.0-only`), unless
separate terms are identified for a file. See the unmodified license text in [LICENSE](LICENSE).
You may copy, modify and convey this work under those terms. It comes WITHOUT ANY WARRANTY,
including the implied warranties of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE;
see sections 15–16. Third-party components retain their own licenses and notices. This notice
neither relicenses third-party material nor grants exceptions to its terms.

Commercial use is permitted by the AGPL. Charging money, internal business use or offering a
SaaS service does not, by itself, require a paid license. Sections 4–6 govern conveyance of
covered source and object code. Under section 13, a modified version that supports remote
network interaction must prominently offer all users interacting with it an opportunity to
receive its Corresponding Source, without charge, through a standard means of copying software.
The required scope follows the license's definitions; it is not automatically every unrelated
system, service or item of business data. Read the license for the actual conditions.

## Source for this build

Release publication requires the exact owned-project source archive
`gridgrind-<version>-sources.tar.gz`, captured from the final frozen source commit before
creating its release tag. It contains the tracked application, helper, repair, build and
installation inputs and is hash-bound in the retained delivery inventory. Source tags and
automatic repository archives are convenience references, not substitutes for that archive.
Release sources and binaries are made available together in a complete public prerelease at
https://github.com/resoltico/GridGrind/releases. Source assets must be verified anonymously
before the versioned runtime image is published.

Build prerequisites and commands are in [Development](docs/DEVELOPER.md). Dependency versions
and retrieval inputs are pinned in the included build inputs and lockfiles. [NOTICE](NOTICE)
identifies separately licensed components. The native-provider capsule retains its mapped
source repairs and helper. Container delivery separately includes
`gridgrind-container-sources-<version>.tar.gz` with the matching Temurin and Ubuntu source/build
inputs described in NOTICE and its source manifest; those vendor materials retain their terms.

A modifier or distributor must supply the Corresponding Source required for the work they
convey, including their changes and required source/build/install inputs. An upstream release
page or an unchanged archive does not supply modifications it omits. Do not include private
runtime configuration, credentials, customer workbooks or other private data in source
publications. The delivery mechanisms are aids, not a determination that every distribution or
deployment meets every license obligation.

Upstream: https://github.com/resoltico/GridGrind.
For separate commercial licensing availability, see [LICENSE-COMMERCIAL.md](LICENSE-COMMERCIAL.md).

## Contribution terms

By submitting a contribution (code, documentation, tests or other material) to this
repository, you:

- confirm that you wrote it or otherwise have the right to submit it on these terms;
- license it to everyone under this repository's public license, `AGPL-3.0-only`; and
- also grant Ervins Strauhmanis a perpetual, worldwide, non-exclusive, royalty-free,
  irrevocable license to use, reproduce, modify, distribute and sublicense it, and to license
  it under other terms, including the commercial terms described in
  [LICENSE-COMMERCIAL.md](LICENSE-COMMERCIAL.md).

You keep the copyright in your contribution. Do not submit material you cannot license on
these terms; third-party components keep their own licenses. Each pull request confirms these
terms with the checkbox in its template.
