"""Console script for pptx_handler."""
import pptx_handler

import typer
from rich.console import Console
from pathlib import Path

app = typer.Typer()
console = Console()


@app.command()
def main():
    # """Console script for pptx_handler."""
    # console.print("Replace this message by putting your code into "
    #            "pptx_handler.cli.main")
    # console.print("See Typer documentation at https://typer.tiangolo.com/")

    # get parent template direcotory via parent direcotory
    cli = Path().resolve()
    parent_dir = cli.parent
    print(f"parent_dir: {parent_dir}")
    powerpoint_dir = parent_dir / 'template'
    print(f"powerpoint_dir: {powerpoint_dir}")





if __name__ == "__main__":
    app()
