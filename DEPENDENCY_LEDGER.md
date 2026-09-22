# Dependency license ledger

Every declared dependency, its license, and why it is here. Enforced by
`tests/unit/test_dependency_ledger.py`, which fails the build if
`pyproject.toml` grows a dependency that is not listed here. Ported from
KitchenSink4Web, which carried the only ledger in the family until the
2026-09-15 license audit (finding D-02).

Licenses below were read from the installed package metadata in this repo's
virtual environment (`importlib.metadata`), not from a search result.

## The rule

**No copyleft and no source-available dependency anywhere in the REQUIRED
install.** Anything questionable lives behind an optional extra, so a
dependency's license never becomes the server's problem. The ship license is
`AGPL-3.0-only` and every permissive license below is one-way compatible into
it, which is the direction that matters: the obligations run to whoever
redistributes those packages, and `pip` resolves each one from its own
publisher with its own license files. KitchenSink4PPT redistributes none of
them, so no attribution obligation attaches to the wheel.

## Required install

The package name is the FIRST column and the license the SECOND, in every
table in this file. The enforcing test reads them positionally.

| Package | License | Direction | Note |
|---|---|---|---|
| `fastmcp` | Apache-2.0 | permissive, one-way into anything | The MCP server framework. Pinned `>=3.4,<4` because a minor bump moved the visibility API mid-build. |
| `lxml` | BSD-3-Clause | permissive | The XML engine, reached directly in 34 source files. This server edits the OOXML itself rather than through an object model, so lxml is the file tier. The BSD-3 no-endorsement clause is the only obligation and it binds redistribution, which does not happen here. |
| `svgelements` | MIT | permissive | SVG parsing for the native vector path, where an SVG becomes editable grouped shapes rather than a pasted image. |
| `latex2mathml` | MIT | permissive | First half of the equation path: LaTeX in, MathML out. |
| `mathml2omml` | MIT | permissive | Second half: MathML to the OMML that PowerPoint actually stores. The wheel's License metadata field reads UNKNOWN, which is a packaging defect rather than an unlicensed release: the classifier says OSI Approved MIT License, and the wheel ships an MIT license text naming amedama that matches the upstream repository. Confirmed 2026-09-15 against the package's own license file, closing the caveat the license audit raised. |
| `regex` | Apache-2.0 AND CNRI-Python | both permissive | Backs the caller-pattern guard in `ops/_regex.py`. Needed rather than convenient: stdlib `re` has no match timeout, the server is single-threaded stdio, and one pathological caller pattern would deny service to the whole session. |
| `packaging` | Apache-2.0 OR BSD-2-Clause | either, permissive | Version comparison in `core/update_check.py`. |
| `pywin32` | PSF | permissive | The COM tier, which drives a real PowerPoint. Declared under a win32 platform marker, so it is never installed anywhere it cannot work. |

No copyleft. No obligation triggered by the current distribution model.

## Optional extras

| Package | License | Extra | Why it is optional |
|---|---|---|---|
| `pillow` | MIT-CMU | `optimize` | Image recompression in `ops/optimize.py`. Optional because a deck ships fine unoptimised. MIT-CMU is the historical-permission-notice variant Pillow ships; permissive, attribution on redistribution only. |
| `fonttools` | MIT | `metrics` (and `dev`) | Real advance widths for the text-fit model in `ops/fontmetrics.py`, reached only from `_overflow_heuristic` and its callers (`fit_text`, `get_autofit_state`, `check_layout`). Optional because the fit model has a working fallback: without it the model keeps the average-glyph-width estimate it always used and every result says which of the two produced it. The import is lazy and inside a try, so a plain install never touches it. Chosen over Pillow, which is already carried for `optimize`, because fontTools is pure Python (one `py3-none-any` wheel covers 3.12 through 3.14 and every platform the day a new interpreter ships, where Pillow needs a compiled wheel per interpreter) and because it reads the `hmtx` advance widths this model wants directly rather than through a FreeType raster context. It is listed in the `dev` extra as well, so continuous integration exercises the measured path instead of skipping every test of it; that does not make it a runtime dependency. License read from the installed package metadata on 2026-09-22. |
| `pytest` | MIT | `dev` | Test-time only, never distributed. |
| `pytest-timeout` | MIT | `dev` | Test-time only, and `required_plugins` in pyproject makes its absence fail the run rather than silently voiding every timeout. A COM hang once ran unbounded during a field round because of exactly that. |
| `python-pptx` | MIT | `dev` | Primarily the independent test oracle: this server writes OOXML directly, so a second implementation reading the result back is the only check that is not marking its own homework. **It is also reached at runtime.** `ops/slides.py` imports it lazily inside `create_presentation` when no template is given, to copy the bytes of the default template python-pptx bundles. On a plain install that call raises ImportError, because this package is declared only in the `dev` extra. Recorded here as found on 2026-09-15 rather than fixed: moving it into the required install or vendoring a blank template is a packaging decision, not a license one, and the MIT license carries no risk either way. |

## Transitive obligations

Transitive dependencies arrive through the packages above and are resolved by
`pip` from their own publishers. None is redistributed by this project, so
none creates an attribution obligation here. A full transitive audit has not
been run; it is the right companion to the one the KitchenSink4Web ledger
also defers.

## Fonts, which are the one real attribution obligation

`docs/fonts/` ships Fraunces and IBM Plex Mono as `.woff2` files on the
documentation page. Both are under the SIL Open Font License 1.1 (OFL),
which requires the copyright notice and license to travel with the font
software. `docs/fonts/OFL-Fraunces.txt` and `docs/fonts/OFL-IBMPlexMono.txt`
are those notices. Unlike every package
above, these files ARE redistributed, which is why the obligation is real
here and nowhere else in this ledger.
