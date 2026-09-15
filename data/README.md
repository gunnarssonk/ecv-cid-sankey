# data/

The mapping workbook goes here.

Download `gcos-ecv-esa-to-ipcc-cid_v1.0.xlsx` from https://doi.org/10.5281/zenodo.21871535 and place it in this directory. The script will find it there without a `--workbook` argument; otherwise pass `--workbook PATH`.

The filename must match exactly: the script is written against v1.0 of the workbook and looks for that name only. For a later version, pass the file with `--workbook PATH` (or rename it) once its sheets have been checked against the script.

The workbook is archived separately on Zenodo under CC BY 4.0 and is deliberately not bundled with this code, so that the two stay independently versioned.