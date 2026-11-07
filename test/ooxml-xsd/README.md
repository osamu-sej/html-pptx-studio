# ECMA-376 schemas (Transitional)

The Office Open XML schemas (ECMA-376 Part 1, Transitional), as python-docx keeps them in `ref/xsd/`. The PowerPoint
export tests (`test/test_pptx_timing.py`) validate every slide part they write against `pml.xsd`, so a file PowerPoint
would have to repair (an attribute out of range, elements out of order) fails the test. `mc:AlternateContent` is
checked through its `mc:Fallback` (PowerPoint 2010's and 2016's own elements are not in these schemas).
