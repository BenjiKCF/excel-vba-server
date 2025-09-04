# Excel VBA MCP

An MCP (Model Context Protocol) server that enables running VBA macros in Excel workbooks using xlwings.

## Features

- Execute VBA macros in Excel workbooks programmatically
- Inject custom VBA code into workbooks
- Save workbooks after macro execution
- Asynchronous execution using FastMCP

## Requirements

- Python 3.10 or higher
- Microsoft Excel installed on the system
- xlwings Python package

## Installation

1. Clone the repository:
```bash
git clone https://github.com/BenjiKCF/excel-vba-server.git
cd excel-vba-server
```

2. Install dependencies using uv:
```bash
uv sync
```

## Usage

### Running the Server

Start the MCP server:
```bash
uv run python server.py
```

The server communicates via stdio and can be integrated with MCP-compatible clients.

### Available Tools

#### run_vba_macro

Executes a VBA macro in an Excel workbook.

**Parameters:**
- `filename`: Path to the Excel workbook file (.xlsm, .xlsx)
- `macro_name`: Name of the VBA macro to execute
- `vba_code`: VBA code to inject and run

**Example:**
```python
await run_vba_macro(
    filename="workbook.xlsm",
    macro_name="MyMacro",
    vba_code="""
Sub MyMacro()
    MsgBox "Hello from VBA!"
End Sub
"""
)
```

## How It Works

1. Opens the specified Excel workbook using xlwings
2. Creates or updates a VBA module named "AutoModule"
3. Injects the provided VBA code into the module
4. Saves the workbook
5. Executes the specified macro
6. Returns the result or any error messages

## Security Note

This tool allows execution of arbitrary VBA code in Excel workbooks. Ensure you trust the source of the VBA code and workbook files to avoid potential security risks.

## Dependencies

- `xlwings`: For Excel automation
- `mcp[cli]`: Model Context Protocol implementation
- `httpx`: HTTP client (currently unused but included for future extensions)

## License

[Add your license here]

## Contributing

[Add contribution guidelines here]
