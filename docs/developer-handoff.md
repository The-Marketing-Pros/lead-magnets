# Developer handoff

Repository: `The-Marketing-Pros/lead-magnets`. Main baseline: `722eecdb1ae5a3d5b674d35aab6f0cf949bf588e`.

Static lead-magnet HTML and downloadable spreadsheet/template assets.

## Verification

New static checks validate the root landing page and local links/assets. Existing downloadable files are preserved byte-for-byte.

```sh
python3 .github/scripts/validate_static.py .
```

Run the gate regression tests with `python3 -m unittest discover -s .github/scripts -p 'test_*.py' -v`. All PR change categories trigger `.github/workflows/engineering-protocol.yml`; missing, skipped, failed or cancelled required jobs fail the aggregate. A failed whole workflow or provider outage remains non-successful.

## Known limits

There is no package build or conventional application test suite. Calculator behavior, browser layout and spreadsheet correctness are not established by link checks.

No application source, data, file or branch is deleted by this rollout. Existing deployment workflows retain their own triggers; a feature-branch push may create an existing provider preview. No production deployment is authorized by the protocol.
