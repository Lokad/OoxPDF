# Independent hierarchical list compositions

Synthetic public OOXML authored directly for list-content contracts; these sources
do not derive from the existing composed-report template. Regenerate with
`python tools/NewHierarchicalListFixtures.py` from the repository root. The generator
uses only the Python standard library and deterministic ZIP metadata.

`expected.json` specifies label strings, prose, equation and picture expectations
before rendering. `mixed-formats.docx` combines Roman, alphabetic and padded decimal
levels with a table continuation, picture, equation and hyperlink. The differently
structured `legal-overrides.docx` uses a concrete outer-level override and legal
decimal references, including a subordinate restart after its parent advances.

WordprocessingML placeholders use the referenced level's effective format;
`isLgl` applies decimal formatting to all references in that level. Latin letter
numbering repeats letters after Z: AA, BB, then AAA, BBB after ZZ. Expectations
follow ECMA-376 Part 1 §§17.9.12,17.9.4,17.18.59 and Microsoft's
[level-text](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.leveltext?view=openxml-3.0.1)
and [legal-numbering](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.islegalnumberingstyle?view=openxml-3.0.1)
descriptions. No Office export or application-view reference is claimed.

Both documents should occupy one page in the pinned-font Final/Review checks.
Labels must paint in source order, descriptions must clear their labels, the
picture must retain both colors, and the equation and navigation target must remain
present. Source semantics establish label content; these checks do not establish
new exact Word geometry or arbitrary numbering/revision support.
