# Cisco BAT XLT to HTML

Convert a Cisco BAT (Bulk Administration Tool) `bat.xlt` Excel template into a single, self-contained HTML/JS application that runs in any modern browser — no macros, no ActiveX, no network access, and no elevated privileges required.

Cisco ships `bat.xlt` with Cisco Unified Communications Manager (CUCM) to help administrators build bulk-import files. The template relies on VBA macros, which many environments block. This project parses the xlt once and generates a standalone HTML file that replicates the template's core functionality in plain JavaScript.

## What the generated HTML does

- **One tab per BAT transaction sheet**, with the same columns and `(Type[maxlen] MANDATORY/OPTIONAL)` constraints parsed from the xlt's header row.
- **"Create File Format" builder** for the dynamic tabs (Phones, Phones-Users, User Device Profiles, Remote Destination Profile, VG gateways), driven by the xlt's hidden `Database` sheet — pick device/line/intercom fields and repeat counts just like the original dialog.
- **"Export to BAT Format"** produces the same comma-separated `.txt` (CRLF line endings, normalized uppercase header line) that the VBA macro's export routine writes, ready for upload to CUCM's Bulk Administration.
- **Validation** mirroring the macro's checks: mandatory fields, max lengths, numeric/boolean types, no commas in values, and per-line Directory Number requirements.
- **Paste from Excel** — copy rows in Excel and paste them straight into the grid (tab- or comma-separated).
- **Dummy MAC address** option on the sheets that support it.

## Requirements

- Python 3
- [`xlrd`](https://pypi.org/project/xlrd/) (reads the legacy `.xlt` binary format):

  ```bash
  pip install xlrd
  ```

## Getting the bat.xlt template

The `bat.xlt` file is Cisco's and is **not included in this repository** — download it from your own CUCM server so it matches your cluster version:

1. Log in to **Cisco Unified CM Administration**.
2. Go to **Bulk Administration > Upload/Download Files**.
3. Find and download the `bat.xlt` file.

## Usage

```bash
python3 src/bat_xlt_to_html.py bat.xlt [output.html]
```

If no output path is given, the HTML is written next to the input with an `.html` extension.

Example:

```bash
python3 src/bat_xlt_to_html.py bat.xlt output/bat.html
```

Then open the generated HTML file in any browser and start building your BAT import files.

## Repository layout

| Path | Contents |
|---|---|
| `src/` | The converter script (`bat_xlt_to_html.py`) |
| `output/` | An example of the generated standalone HTML application (from the CUCM 14.0.1 template) |

## Notes

- Always generate the HTML from the `bat.xlt` of the cluster you are importing into, so the columns and constraints match your CUCM version.
- The repeating-group definitions (Speed Dials, BLF Speed Dials, BLF Directed Call Parks, Remote Destinations) are compiled into the VBA macros rather than the `Database` sheet, so they are mirrored as stable constants in the script.

## License

Licensed under the [Apache License 2.0](LICENSE).
