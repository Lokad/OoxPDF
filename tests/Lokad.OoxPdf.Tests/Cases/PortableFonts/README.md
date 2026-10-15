# Portable rendering font

Unmodified `DejaVuSans.ttf` from Debian `fonts-dejavu-core` 2.37-8, upstream
<https://dejavu-fonts.github.io/>. SHA-256:
`AE7B7855E115A5966D8B1B3F80F254CCC117EC86F9965E202EE2940453837280`.

The font is redistributed under the Bitstream Vera permission notice, with
DejaVu changes in the public domain. The complete package copyright notices are
retained in `LICENSE.txt`; the `debian/*` packaging license does not apply to the
font. This is a test input, never a shipped library dependency.

The portable rendering group maps every family/style request to this one face.
It checks composed content, image painting and PDF resources on both platforms.
It does not establish Office typography or replace independent visual gates.

`letter-normal-math.docx` derives from the pinned composition letter by adding
`m:rPr/m:nor m:val="1"` to each mathematical run. The font lacks the default
mathematical italic alphabet; authored normal math exercises its supported glyphs
and MATH metrics. The fixed portfolio is unchanged. Document/font source hashes
and the exact transformation are retained in `provenance.json`.

`math-navigation-normal.docx` derives from the public compound math/navigation
probe by the same normal-math transformation across body, cell, header and footer
runs. Its adjacent JSON pins both source and result. The portable navigation
check verifies four internal destinations and mathematical painting in both modes.
