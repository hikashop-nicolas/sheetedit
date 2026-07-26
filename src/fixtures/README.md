# Test fixtures

Binary files used by the tests, kept here because they cannot be built from source.

| File | Source | Licence |
|---|---|---|
| `compound.xls` | Produced locally by LibreOffice from `demo/c-outline.xlsx` | Ours |
| `macros-cp950.xlsm` | Apache POI test data, `test-data/spreadsheet/45431.xlsm` | Apache-2.0 |
| `macros-cp1252.xlsm` | Apache POI test data, `test-data/spreadsheet/61495-test.xlsm` | Apache-2.0 |

The two `.xlsm` files carry a real `vbaProject.bin`, which is the only way to test the VBA reader
end to end: the container format and the compression can each be checked in isolation, but their
combination needs a genuine project written by Excel. They use different code pages (950 and 1252)
on purpose, since decoding is part of what is being tested.

Their macros ARE executed by the tests now: one of them runs end to end against a stub object model
in `vba-excel.test.ts`, and again against the real one. That is the point of having genuine files
rather than hand-written ones, since a macro nobody wrote for this project is the only honest test
of an interpreter.
