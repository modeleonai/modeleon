# Notebooks

Curated Jupyter tutorials demonstrating the engine end-to-end. Run
top-to-bottom and open the resulting `.xlsx` to see live formulas.

| Notebook | Covers |
|---|---|
| [01_first_model.ipynb](01_first_model.ipynb) | Variables → formulas → tabs → cross-sheet refs → self_ref → Excel stdlib → `.xlsx` output |

## Conventions

- **Commit outputs stripped.** Notebook outputs produce noisy diffs and
  leak local paths. Run [`nbstripout`](https://github.com/kynan/nbstripout)
  before committing, or install it as a pre-commit hook:
  ```sh
  pip install nbstripout
  nbstripout --install     # per-clone, writes to .git/config
  ```
- **One thesis per notebook.** Don't make scratchpads here — use
  `packages/engine/scratch/` for throwaway work (gitignored).
- **Must run top-to-bottom on a fresh kernel.** Each notebook should
  build its own explicit root — e.g.
  ``with mo.MultiVariable("Tour") as model: ...`` — so
  cell re-runs don't accumulate global state.
- **Assume the reader has never seen the library.** Notebooks are the
  first hands-on experience for new users — favor clarity over cleverness.

## Running

```sh
cd packages/engine
pip install -e ".[dev]" jupyterlab
jupyter lab notebooks/
```
