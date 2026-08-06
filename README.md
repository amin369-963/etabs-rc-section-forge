# ETABS RC Section Forge

An open-source Python utility for generating reinforced-concrete beam and column section libraries in ETABS.

The current alpha release is designed for **ETABS 20 and ETABS 21 on Windows**. It connects to ETABS through the CSI COM API, reads concrete and reinforcing-steel materials from the active model, resolves reinforcing-bar sizes from ETABS data, and creates rectangular beam or column sections.

> This project is independent of Computers and Structures, Inc. ETABS is a trademark of Computers and Structures, Inc.

## Current status

- Version: `0.1.0-alpha`
- ETABS 20.3: tested successfully
- ETABS 21: code path included; integration verification is still required
- Windows only for direct ETABS automation
- Python 3.9 or newer
- License: MIT

This is an engineering automation tool, not a design-code compliance checker. Review all generated sections before using them in production models.

## Main features

- Attaches to an active ETABS instance or attempts to launch ETABS 20/21.
- Reads the first available concrete material as the default concrete material.
- Reads the first available rebar material as the default longitudinal and confinement material.
- Supports metric and US reinforcing-bar input notation.
- Supports these working-unit presets:
  - `kgf-cm-C`
  - `kN-mm-C`
  - `kN-m-C`
  - `kip-in-F`
  - `kip-ft-F`
- Restores the model's original ETABS display units after execution.
- Generates rectangular reinforced-concrete beam sections.
- Generates feasible reinforced-concrete column reinforcement arrangements.
- Checks ETABS API return codes for section and reinforcement assignments.
- Includes unit tests that run without ETABS by using a fake ETABS model.

## Repository files

```text
etabs-rc-section-forge/
├── etabs_rc_section_builder.py
├── test_etabs_rc_section_builder.py
├── requirements.txt
├── README.md
├── LICENSE
└── .gitignore
```

## Installation

Use a 64-bit Python installation that matches the ETABS installation architecture.

```powershell
python -m pip install -r requirements.txt
```

The dependency file currently pins:

```text
comtypes==1.4.16
```

## Running the tests

The tests do not require ETABS:

```powershell
python -m unittest -v test_etabs_rc_section_builder.py
```

## Running with ETABS

1. Back up the ETABS model.
2. Close unnecessary ETABS instances.
3. Open ETABS 20 or ETABS 21.
4. Open the target model and wait until it is fully loaded.
5. Confirm the configuration values near the top of `etabs_rc_section_builder.py`.
6. Run:

```powershell
python etabs_rc_section_builder.py
```

The program prints the detected ETABS version and model path before creating sections. Confirm that the displayed model path is the intended model.

## Configuration

The default active path creates beam sections only:

```python
DO_BEAMS = 1
DO_COLUMNS = 0
```

To run the column generator:

```python
DO_BEAMS = 0
DO_COLUMNS = 1
```

Select a working unit preset:

```python
WORKING_UNITS = "kgf-cm-C"
```

Select the preferred reinforcing-bar notation:

```python
REBAR_SYSTEM = "metric"  # or "us"
```

Material values default to `None`, which causes the program to select the first matching material returned by ETABS:

```python
MATERIALS = {
    "concrete": None,
    "rebar_long": None,
    "rebar_tie": None,
}
```

A specific ETABS material name may be assigned instead of `None`.

## Metric and US reinforcing bars

Metric examples:

```python
COL_BAR_DIAM_MM = [20, 25]
COL_TIE_SIZE = "8d"
```

US examples:

```python
REBAR_SYSTEM = "us"
COL_BAR_DIAM_MM = ["#6", "#8"]
COL_TIE_SIZE = "#3"
```

The program matches requested bars against the actual diameter data reported by ETABS. It does not silently accept a large diameter mismatch.

## `comtypes` troubleshooting

`comtypes` generates Python wrappers for COM type libraries. The wrappers are normally generated automatically when a COM object exposes type information. Stale wrappers can occasionally remain after changing the ETABS version, Python environment, or `comtypes` version.

Official `comtypes` client documentation:

- https://comtypes.readthedocs.io/en/stable/client.html

### `Operation unavailable` from `GetActiveObject`

This usually means no accessible ETABS instance is registered in the Windows Running Object Table.

1. Open ETABS manually.
2. Open the target model.
3. Keep only one ETABS instance open when possible.
4. Run the script again.

### `WinError 740` or `The requested operation requires elevation`

ETABS and Python are running at different Windows privilege levels.

Use one of these configurations:

- Run both ETABS and Python normally; or
- Run both ETABS and Python as Administrator.

Running only one of them as Administrator can block COM attachment and process launch.

### `Class not registered`, `Invalid class string`, or COM error 429

- Confirm that ETABS 20 or ETABS 21 is correctly installed.
- Confirm that Python and ETABS use compatible 64-bit architecture.
- Start ETABS manually and retry.
- Repair the ETABS installation if the COM registration is missing.

### Stale or corrupted generated wrappers

Close ETABS and all Python processes, then run:

```powershell
python -m comtypes.clear_cache
```

Restart ETABS and rerun the program.

The generated-wrapper directory is controlled by `comtypes.client.gen_dir`. See the official `comtypes` documentation for details.

### Multiple ETABS installations

The script searches these default executable paths:

```text
C:\Program Files\Computers and Structures\ETABS 21\ETABS.exe
C:\Program Files\Computers and Structures\ETABS 20\ETABS.exe
```

Attaching to an already open target model is preferred over launching a new instance.

## Safety notes

- Always work on a backup model during testing.
- The script can redefine sections when overwrite behavior is enabled.
- The script does not automatically save the ETABS model.
- Verify generated beam and column reinforcement data inside ETABS.
- Do not treat generated candidate sections as proof of compliance with ACI, Eurocode, CSA, Iranian codes, or any other design standard.

## Contributing

Bug reports and pull requests are welcome. For ETABS integration issues, include:

- ETABS version
- Python version and architecture
- `comtypes` version
- Windows version
- Relevant console output
- Whether ETABS and Python were run normally or as Administrator

Do not upload proprietary ETABS models or confidential project data.
