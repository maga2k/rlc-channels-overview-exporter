# RLC Channel Overview Exporter

MATLAB utilities for exploring the Motion Appkied (former McLaren Applied) System Monitor API and exporting `.rlc` Channel Overview data to Excel (`.xlsx`).

## Features

- **API Explorer:** inspects installed System Monitor DLLs using .NET reflection and generates text reports of exported types, properties and methods.
- **RLC to Excel:** exports parameter IDs, descriptions, applications, sizes, Prime flags, channel configurations and logging rates.
- **Excel formatting:** applies filters, freezes headers and highlights rows in yellow when the `Note` cell is not empty.

## Requirements

- MATLAB on Windows with .NET support.
- McLaren Electronic Systems System Monitor 8.85.1.2 installed.
- `SystemMonitorAPI.dll` available locally.
- Microsoft Excel installed for optional workbook formatting.

## Usage

1. Clone or download this repository.
2. Open each script in MATLAB.
3. Update the configuration paths at the top of the script.
4. Run `explore_SM_API_dll.m` to inspect the installed DLLs.
5. Run `rlc_to_xlsx.m` to export the loaded `.rlc` configuration to Excel.

The scripts create their output directories automatically when needed.

## Project Structure

```text
rlc-channel-overview-exporter/
├── rlc_to_xlsx.m
├── explore_SM_API_dll.m
├── output/
│   └── .gitkeep
├── sm_dump/
│   └── .gitkeep
├── .gitignore
├── LICENSE
└── README.md
```

- `output/` — generated Excel workbooks.
- `sm_dump/` — generated DLL inspection reports.
- `.gitkeep` — preserves empty output directories in Git.

## Notes

This project is an independent utility and is not affiliated with or endorsed by McLaren. System Monitor and its API are proprietary software and must be installed separately.

## License

See [LICENSE](LICENSE) for the license applying to this repository.