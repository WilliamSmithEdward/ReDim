# Notes for agents

<!-- repo-standards:begin. Copied from WilliamSmithEdward/repo-standards, templates/agents/AGENTS-block.md. Change it there; the weekly rescan fails a copy that differs. -->
## Releases, CI and security

These rules are the same in every WilliamSmithEdward repository.

- **How a release happens here:** pushing a `vX.Y.Z` tag runs Publish, which builds the release files in CI and creates the GitHub release with them, their signed provenance and the security reports. Any other step, such as a marketplace upload, is described elsewhere in this file.
- **Starting a workflow by hand never releases anything.** Publish and every
  release report are dry runs when started with `gh workflow run` or the Run
  workflow button. They build, scan and assemble the release files exactly
  as a release would, and upload them as the `release-preview` artifact
  instead. Run one after changing anything on the release path:
  `gh workflow run <file> --ref main`, then
  `gh run download <run-id> -n release-preview`.
- **Do not create, publish, edit or delete a release or a `v*` tag** unless
  the owner asks for it. A `v*` tag cannot be moved or deleted once pushed.
- **Every change to `main` goes through a pull request** that passes CI
  passed, Security passed and Malware scan passed. No one can push to `main`
  directly or skip the checks, admins included. Push a branch, open a pull
  request, and let it merge itself: `gh pr merge --auto --squash <number>`.
- **Pins.** Actions by full commit SHA with the version as a comment. Images
  by digest, in `.github/security/<tool>/Dockerfile`. Python tools from the
  hash-locked `.github/requirements/<purpose>.txt`, compiled from the `.in`
  beside it with
  `uv pip compile <purpose>.in --universal --generate-hashes --python-version 3.12 -o <purpose>.txt`.
  Runners are named releases, never `-latest`.
- **Updates merge themselves.** Dependabot and the Update YARA rules workflow
  open pull requests that merge once the three checks pass, except a
  third-party major version, which waits for the owner. Leave them alone
  unless asked.
- **A scanner finding is fixed or accepted with a written reason** in the
  repository's accepted list. Never silence a scanner without one.
<!-- repo-standards:end -->

## Releasing

1. Bump `REDIM_VERSION` in `src/ReDimUI.cls` and add a
   `## X.Y.Z - YYYY-MM-DD` section to CHANGELOG.md. The release's notes are
   taken from that section, and Publish fails without one.
2. Run `python tools/stamp_release.py` to write the new header into every
   release source. Publish runs it again and fails if it changes anything.
3. Merge to `main` through a pull request.
4. Optionally dry-run Publish: `gh workflow run publish.yml --ref main`,
   then `gh run download <run-id> -n release-preview`.
5. The owner tags the merged commit:
   `git tag vX.Y.Z && git push origin vX.Y.Z`.

Publish refuses a tag that is not `REDIM_VERSION`. It releases
`ReDimUI.cls`, `ReDimHost.bas` and the `ROneCOne.cls` pinned in the
workflows, with CRLF endings, and no workbooks. The demos are built from
`demo/vba/` with `python tools/build_workbooks.py`.
