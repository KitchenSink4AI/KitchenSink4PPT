# KitchenSink4PPT : License Notice

KitchenSink4PPT is Copyright (c) 2026 Alvut Consulting, LLC.

KitchenSink4PPT is dual-licensed:

**AGPL-3.0 (open source).** Free for anyone (individuals, academics, and
businesses) for any use that complies with the AGPL's terms. Those terms
include sharing source, including your modifications, when you distribute
the software or make it available over a network.

**Commercial license.** For organizations that want to build KitchenSink4PPT
into their own products or services without the AGPL's source-sharing
obligations. Contact licensing@kitchensink4.ai.

Copyright (c) 2026 Alvut Consulting, LLC. KitchenSink4AI is a product line
of Alvut Consulting, LLC.

Not affiliated with, endorsed by, or sponsored by Microsoft Corporation.
Microsoft, PowerPoint, GitHub and Copilot are trademarks of the Microsoft group of companies.

Required Notice: Copyright (c) 2026 Alvut Consulting, LLC
(https://github.com/KitchenSink4AI/KitchenSink4PPT)

## Third-party software

KitchenSink4PPT does not redistribute these packages; pip installs each
from its own publisher under its own license. The first column is the
package, the second its license.

### Required install

| Package | License | Use |
|---|---|---|
| `fastmcp` | Apache-2.0 | MCP server framework |
| `lxml` | BSD-3-Clause | OOXML parsing and editing |
| `svgelements` | MIT | SVG parsing |
| `latex2mathml` | MIT | LaTeX to MathML |
| `mathml2omml` | MIT | MathML to OMML |
| `regex` | Apache-2.0 AND CNRI-Python | Pattern matching with timeouts |
| `packaging` | Apache-2.0 OR BSD-2-Clause | Version comparison |
| `pywin32` | PSF | COM bridge to PowerPoint (Windows only) |

### Optional extras

| Package | License | Extra |
|---|---|---|
| `pillow` | MIT-CMU | `optimize` |
| `fonttools` | MIT | `metrics`, `dev` |
| `pytest` | MIT | `dev` |
| `pytest-timeout` | MIT | `dev` |
| `python-pptx` | MIT | `dev` |
| `build` | MIT | `dev` |
| `setuptools` | MIT | `dev` |
| `cryptography` | Apache-2.0 OR BSD-3-Clause | `dev` |

### Fonts

`docs/fonts/` includes Fraunces and IBM Plex Mono as `.woff2` files for the
documentation page. Both are licensed under the SIL Open Font License 1.1
(OFL). The license notices are `docs/fonts/OFL-Fraunces.txt` and
`docs/fonts/OFL-IBMPlexMono.txt`.

## Third-party components

The default deck template is a modified copy of the file shipped with python-pptx 1.0.2 (original SHA-256 e10cc9e120961f6bd4074a373c9c80d2a06c497157e8f4972977b7bea83a8f34; shipped file SHA-256 eb33715284d0a4838d3c1ffaa1245df153fef590ea5e77fa8f9d2fe4bc6867e6). Its thumbnail image and its description, last-modified-by and date properties were removed; every slide, layout, master and theme part is unchanged. It is used under the MIT License.

Copyright (c) 2013 Steve Canny, https://github.com/scanny

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
