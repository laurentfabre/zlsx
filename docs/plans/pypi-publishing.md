# PyPI publishing — deferred by decision

**Status: DEFERRED (2026-07-26), not blocked-and-forgotten.** The
remaining step is a PyPI-side action only the account owner can perform,
and it has been consciously postponed rather than overlooked.
Everything on the repository side is correct and verified.

Practical consequence: **`pip install py-zlsx` does not work and will
not until this is done.** The wheels build fine on every release; they
simply never reach the index. Anyone needing the Python binding installs
from source in the meantime.

Nothing here needs re-investigating when it is picked up — the values
below are verified against the current workflow, and the release job
reminds you at the moment it matters (it preflights on tag push and
names every field). Do not re-raise it as a discovered problem; it is a
known, priced decision.

---

## Where it stands

`py-zlsx` has **never been published**. Every release from v0.2.4
onward built all five wheels successfully and then failed at the final
step:

```
Trusted publishing exchange failure:
* `invalid-publisher`: valid token, but no corresponding publisher
  (Publisher with matching claims was not found)
```

7/7 runs, all the same. `https://pypi.org/pypi/py-zlsx/json` returns
**404** — the project does not exist, so `pip install py-zlsx` has
never worked for anyone.

## Why it is not a workflow bug

The workflow requests a GitHub OIDC token and offers it to PyPI. PyPI
looks for a publisher whose claims match, finds none, and refuses. That
is correct behaviour on both sides: **nobody has told PyPI that this
workflow is allowed to publish this project.**

Because the project does not exist yet, the publisher has to be created
through the **pending publisher** form. A normal (existing-project)
publisher cannot be attached to a name nobody owns — which is the trap
this sat in.

## The fix (≈2 minutes, requires the PyPI account)

Go to <https://pypi.org/manage/account/publishing/>, and under
**"Add a new pending publisher"** enter exactly:

| Field | Value |
|---|---|
| PyPI Project Name | `py-zlsx` |
| Owner | `laurentfabre` |
| Repository name | `zlsx` |
| Workflow name | `pypi.yml` |
| Environment name | `pypi` |

Every value is asserted against the workflow — `environment: pypi` is
declared on the `publish` job, and the file is `.github/workflows/pypi.yml`.
A mismatch in any single field reproduces the same error.

Then publish by pushing a tag:

```sh
git tag v0.5.0 && git push origin v0.5.0
```

or run the workflow manually with `publish=true`.

## Alternative, if you would rather not use trusted publishing

Create an API token at <https://pypi.org/manage/account/token/>, add it
as the repository secret `PYPI_API_TOKEN`, and give the publish step:

```yaml
      - name: Publish to PyPI
        uses: pypa/gh-action-pypi-publish@release/v1
        with:
          password: ${{ secrets.PYPI_API_TOKEN }}
```

Trusted publishing is the better default — no long-lived secret — so
this is only worth doing if the OIDC flow keeps giving trouble.

## What was changed on this side

The workflow now diagnoses the situation instead of failing opaquely:

- a **preflight** step checks whether the project exists on PyPI and
  emits a warning naming every field to enter if it does not;
- a **failure** step restates the remediation, so the log says what to
  do rather than only what went wrong.

Neither changes publishing behaviour. Once the pending publisher exists,
the next tag push publishes and both steps go quiet.
