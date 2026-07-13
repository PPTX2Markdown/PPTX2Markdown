# pptx2markdown Agent Installation Instructions

You are an AI coding agent helping the user install `pptx2markdown` on this
machine. Install the released package from PyPI. Do not install from a repository
clone, local wheel, `dist/`, or another local build artifact unless the user
explicitly asks for a development install.

You may need permission to run commands, install Python or an isolated Python
tool manager, modify a shell profile, or install LibreOffice. Explain each
state-changing command before running it. Ask before administrator, `sudo`,
password, system package-manager, shell-profile, or PATH changes.

## 1. Choose the conversation language

Ask which language the user wants for installation prompts and the final report.
Use that language throughout the installation.

## 2. Inspect the environment

Detect the operating system, CPU architecture, current shell, and available
Python/tool managers with non-destructive commands. Check, where applicable:

```bash
python3 --version
python --version
uv --version
pipx --version
```

`pptx2markdown` requires Python 3.12 or newer. Do not replace a system Python or
change the user's default interpreter automatically. If a suitable interpreter
is missing, explain the platform-appropriate official installation choices and
ask which one to use before installing anything.

If this agent cannot run commands, stop and direct the user to the manual
installation section in the README. Do not claim installation succeeded.

## 3. Ask for the installation scope

Offer these choices:

1. Isolated CLI installation with `uv tool` (recommended when `uv` is present).
2. Isolated CLI installation with `pipx`.
3. Installation into the currently active virtual environment.

Do not install into the system Python with elevated privileges. Do not create a
new virtual environment in an arbitrary directory without confirming its path
with the user.

## 4. Install from PyPI

Use the command matching the user's choice.

With `uv`:

```bash
uv tool install --upgrade pptx2markdown
```

With `pipx`, check whether the package is already installed and then install or
upgrade it:

```bash
pipx install pptx2markdown
pipx upgrade pptx2markdown
```

Inside the user-approved active virtual environment:

```bash
python -m pip install --upgrade pptx2markdown
```

If the command directory is not on PATH, report the exact tool-generated advice.
Ask before editing a shell profile or running commands such as `uv tool update-shell`
or `pipx ensurepath`. Prefer telling the user how to refresh or restart the shell
instead of silently changing profile files.

## 5. Ask whether legacy PPT support is needed

Normal `.pptx` conversion does not require LibreOffice. LibreOffice is optional
for legacy `.ppt` conversion and EMF/WMF handling.

Ask whether the user needs those features. If not, skip LibreOffice. If yes,
detect whether `soffice` or LibreOffice is already available. Before using
Homebrew, APT, another system package manager, administrator privileges, or
`sudo`, explain the exact command and ask for approval.

Do not install Microsoft PowerPoint or change Office settings.

## 6. Verify the installation

Open a new or refreshed shell when PATH changes require it, then run:

```bash
pptx2markdown --help
```

Confirm that help lists the static parser options, including
`--output-format {markdown,json}` and `--headings {auto,strict}`. It must not
list Surya, OCR, VLM, or reading-order provider options.

If the installation method exposes its own package listing, use it to report the
installed version (`uv tool list`, `pipx list`, or Python package metadata).

Do not convert personal presentation files merely to test installation. Offer to
run a conversion only if the user provides or selects a file and approves the
output directory.

## 7. Final report

Summarize only verified facts:

- detected operating system, shell, and Python version;
- selected installation method and installed package version;
- whether CLI help verification passed;
- whether LibreOffice was installed, already present, skipped, or declined;
- any PATH refresh or manual action still required;
- the exact command for a first conversion:

```bash
pptx2markdown deck.pptx
```

Explain that output is written under `./output/<deck-name>/` by default. If any
step failed, report the failure and the safest next action instead of describing
the installation as complete.
