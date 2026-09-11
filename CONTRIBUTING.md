# Contributing

Thanks for helping. This page is the short version; the long version is
under [`docs/development/`](docs/development/architecture.md).

## Set up

```bash
git clone https://github.com/ForLegalAI/mcp-ms-office-documents
cd mcp-ms-office-documents
python -m venv .venv && source .venv/bin/activate
pip install -r requirements-dev.txt      # runtime deps + ruff + pytest-asyncio
```

Python 3.12. No `.env` is needed for development or tests; the defaults are
local storage and INFO logging. To run the server: `python main.py`, then
point a client at `http://localhost:8958/mcp`.

## Before you push

```bash
ruff check .
pytest -m "not network"
```

That is exactly what CI runs on every pull request. `requirements-dev.txt`
matters: installing only `requirements.txt` leaves both commands without
their tool. Tests marked `network` fetch an image from a public host and are
excluded from CI so a third-party outage cannot fail an unrelated change;
run `pytest` with no marker filter to include them locally.

Generated inspection files land in `tests/output/` (git-ignored). See
[`docs/development/testing.md`](docs/development/testing.md) for the
patterns each package uses.

## Making a change

- **Branch from `master`**, open a pull request, keep it to one topic.
- **Commit messages** follow Conventional Commits with the package as scope,
  as the history does: `feat(pptx): …`, `fix(docx): …`, `docs: …`, `ci: …`.
  Add `!` for a breaking change. Say why, not just what; the body is where
  the reasoning that would otherwise live in a code comment goes.
- **A bug fix comes with a regression test** in the file that covers that
  behaviour, and a docstring note naming the issue if one exists.
- **Documentation is part of the change, not a follow-up.** Every pull
  request updates, in the same commits, each page under `docs/` that
  describes what it touched: the developer page for the package
  (`docs/development/tools/*.md`, `docs/development/*.md`) for internal
  changes, and the user reference under `docs/` for anything a user or the
  calling model sees. If nothing described what you touched, say so in the
  commit message.
- **A user-visible change touches three places together**: the tool
  description in `main.py` (what the calling model reads), the reference
  page under `docs/`, and the tests. They drift otherwise.
- **Environment variables** are read only in `config.py`. Every one must
  appear in `.env.example` and `docs/configuration.md`; a test fails if not.
- **Keep the rules in [`AGENTS.md`](AGENTS.md).** They apply to people too.

Adding a document type or a storage backend has its own checklist:
[`adding-a-tool.md`](docs/development/adding-a-tool.md),
[`adding-a-backend.md`](docs/development/adding-a-backend.md).

## Repository layout

```
main.py                 tool declarations, startup, server launch
config.py               all configuration
{docx,xlsx,pptx,email,xml}_tools/   one package per document type
upload_tools/           strategy dispatch + one module per backend
admin/                  optional template-admin UI
default_templates/      shipped templates
custom_templates/       your templates (git-ignored except the example letter)
config/                 the three commented master YAML files (tracked) and
                        admin-written *_templates.d/ directories (ignored)
docs/                   user reference and development docs
tests/                  one file per behaviour
```

`custom_templates/` and everything under `config/` except the three master
YAML files are git-ignored on purpose. If you need to track a new shipped
file there, add a `!` rule to `.gitignore`.

## Releases

A release is a GitHub release. Publishing one runs the `cd.yml` workflow,
which builds the image for amd64 and arm64 and pushes
`georgx22/mcp-office-docs:<tag>`. A full release also moves `latest`; a
prerelease does not, so nobody pulling the default tag gets a beta.

There is no changelog file; the release notes are the changelog.

## Reporting problems

Open an [issue](https://github.com/ForLegalAI/mcp-ms-office-documents/issues).
For a security problem, read [SECURITY.md](SECURITY.md) first.
