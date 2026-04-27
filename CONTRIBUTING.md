<!-- SPDX-License-Identifier: Apache-2.0 -->

# Contributing to Modeleon

Thanks for your interest. Modeleon is open-source under the Apache
License 2.0 and welcomes contributions of every size.

> **A note up front.** Modeleon is currently maintained by one person.
> I'll do my best to respond, but please expect **weeks, not days** for
> issue triage and PR reviews. If something feels stuck, a polite
> nudge in the thread is fine.

## License of contributions

By submitting a contribution to this repository (a pull request,
patch, or content via an issue), **you agree to license your
contribution under the [Apache License, Version 2.0](LICENSE)**, the
same license that covers the rest of the project.

This matches Apache 2.0's standard "Submission of Contributions"
clause (Section 5). It means:

- You retain copyright on your contribution.
- You grant Modeleon and every downstream user a perpetual, worldwide,
  royalty-free license to use, modify, and redistribute your
  contribution under Apache 2.0.
- You confirm that you have the right to submit it — i.e. the work is
  yours, or you have explicit permission from the rights holder
  (employer, collaborators).

No separate Contributor License Agreement (CLA) is needed. Submitting
the PR is itself the agreement.

## Where to go

- **Bug?** Open an [Issue](https://github.com/modeleonai/modeleon/issues)
  with a minimal reproducer (5–20 lines of Python ideally), the full
  traceback, your Python version (`python --version`), and your
  modeleon version (`pip show modeleon`).
- **Question?** Open a
  [Discussion](https://github.com/modeleonai/modeleon/discussions). I
  read them when I can — sometimes promptly, sometimes after a while.
  Other users may help before I get there.
- **Feature idea?** Open a Discussion first to talk through the use
  case before opening an issue. I'd rather discuss the problem than
  the solution.

## Submitting a pull request

For non-trivial changes, **open an issue first** so we can align on
scope. Tiny fixes (typos, doc clarifications, one-line bugs) can skip
this — just open the PR.

The bar for merge:

1. **One concern per PR.** A bug fix and a refactor in the same PR is
   harder to review and harder to revert.
2. **Tests.** Every code change needs a test — either showing the new
   behaviour or pinning the bug you fixed.
3. **Lint passes.** `ruff check src/ tests/` must be clean.
4. **Commit messages.** Imperative present tense ("fix: handle empty
   cash-flow list in IRR"), first line under 70 chars, longer body
   when the *why* isn't obvious from the diff.

I may not get to your PR right away. That's not a rejection — it's
the realistic pace of a small-team project.

## Code style

- Python 3.12+. `from __future__ import annotations` in most files;
  keep that pattern in new ones.
- `ruff` is the source of truth for style. Line length 100. Config
  lives in `pyproject.toml`.
- Type hints on public APIs. Internal helpers can skip them when
  the type is obvious.
- Docstrings only on public functions/classes — and prefer one tight
  paragraph over a long structured block.
- Avoid new dependencies. The engine intentionally has only three
  runtime deps (`numpy`, `openpyxl`, `python-dateutil`). New deps
  need a strong justification in the PR description.

## Running tests

```bash
cd packages/engine
pip install -e ".[dev]"
pytest                              # all tests
pytest tests/test_variable.py -v    # one file
ruff check src/ tests/              # lint
```

The full suite runs in a few seconds.

## Code of conduct

Be respectful in issues, discussions, and PR review. Don't make
Modeleon a place where people regret opening a thread. We follow the
[Contributor Covenant](https://www.contributor-covenant.org/) in
spirit; report serious issues to admin@modeleon.ai.

## Questions?

Open a [Discussion](https://github.com/modeleonai/modeleon/discussions)
or email admin@modeleon.ai.
