# Engineering protocol rollout evidence

Repository: `The-Marketing-Pros/lead-magnets`. Base: `722eecdb1ae5a3d5b674d35aab6f0cf949bf588e`.

## Candidate scope

Preservation-first protocol, single-human owner policy, independent agent review records, dependency maintenance and a fail-closed aggregate CI gate. New static checks validate the root landing page and local links/assets. Existing downloadable files are preserved byte-for-byte.

## V1 — implementer verification

Codex is the implementer. Record exact candidate and hosted CI URLs in the PR; an evidence-only commit does not inherit a review for a different code snapshot. Local gate tests and workflow validation are pending until their actual results are recorded.

## Independent review

Pending. No R1/R2 or GitHub approval is claimed without an actual review receipt. Andrew is the sole human owner; another human is not required.

## Remaining limitations

There is no package build or conventional application test suite. Calculator behavior, browser layout and spreadsheet correctness are not established by link checks.

No application or production acceptance is inferred from protocol validation. No merge, deletion, production deployment, live migration or external message is performed by this change.

## Current local protocol verification — 2026-09-12

The dependency-free gate regression suite passed, and all GitHub workflow/Dependabot YAML parsed with Ruby/Psych. Every required result failure, cancellation, skip, missing/malformed dependency and workflow dependency-set mismatch is exercised. This is structural and synthetic-result validation, not a rerun of application suites or live integration acceptance. The candidate adds/modifies only protocol/CI/documentation files; no existing file or branch is deleted. Hosted application checks and independent review are recorded separately when they finish.
