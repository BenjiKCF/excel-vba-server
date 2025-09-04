from typing import Any
import httpx
from mcp.server.fastmcp import FastMCP
import xlwings as xw
import asyncio
import re

# Initialize FastMCP server
mcp = FastMCP("weather")

@mcp.tool()
async def run_vba_macro(filename: str, macro_name: str, vba_code: str) -> str:
    """Run a VBA macro in an Excel workbook.

    Args:
        filename: The path to the Excel workbook file.
        macro_name: The name of the macro to run.
        vba_code: The VBA code to inject and run.
    """
    def _run_macro():
        try:
            wb = xw.Book(filename)
            vb_project = wb.api.VBProject

            # Check if module already exists
            module_name = "AutoModule"
            try:
                vb_module = vb_project.VBComponents(module_name)
                # Clear existing code
                vb_module.CodeModule.DeleteLines(1, vb_module.CodeModule.CountOfLines)
            except:
                # Create new module if not found
                vb_module = vb_project.VBComponents.Add(1)
                vb_module.Name = module_name

            # Add new code
            vb_module.CodeModule.AddFromString(vba_code)

            # Save and run
            wb.save()
            wb.macro(macro_name)()

            return f"Macro '{macro_name}' executed successfully."

        except Exception as e:
            return f"Error running macro: {e}"

    # Run the synchronous xlwings code in a thread pool
    loop = asyncio.get_event_loop()
    return await loop.run_in_executor(None, _run_macro)


if __name__ == "__main__":
    # Initialize and run the server
    mcp.run(transport='stdio')