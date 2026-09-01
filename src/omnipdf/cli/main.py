"""
FarazanPDFConverter CLI: Command-line interface for layout-preserving PDF-to-Word conversion.
"""

import os
import json
import typer
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TimeElapsedColumn

from omnipdf.core.converter import OmniConverter, ConversionMode
from omnipdf.core.inspector import PDFInspector

app = typer.Typer(
    name="omnipdf",
    help="FarazanPDFConverter: Advanced Layout-Preserving PDF to Word (DOCX) Converter",
    add_completion=False,
)
console = Console()


@app.command("convert")
def convert(
    input_pdf: str = typer.Argument(..., help="Path to input PDF file"),
    output_docx: str = typer.Option(None, "-o", "--output", help="Path to output DOCX file"),
    mode: str = typer.Option("auto", "-m", "--mode", help="Conversion mode: academic, exact, flow, auto"),
    math: bool = typer.Option(True, "--math/--no-math", help="Convert math/LaTeX to native Word OMML equations"),
    code: bool = typer.Option(True, "--code/--no-code", help="Highlight code and algorithms in callout boxes"),
    images: bool = typer.Option(True, "--images/--no-images", help="Extract and embed high-res figures"),
    dpi: int = typer.Option(300, "--dpi", help="DPI for figure and background extraction"),
):
    """Convert a PDF file to an editable Word document (.docx)."""
    if not os.path.exists(input_pdf):
        console.print(f"[bold red]Error:[/] PDF file not found: {input_pdf}")
        raise typer.Exit(code=1)

    if not output_docx:
        base, _ = os.path.splitext(input_pdf)
        output_docx = f"{base}.docx"

    console.print(Panel.fit(
        f"[bold cyan]FarazanPDFConverter Converter[/]\n"
        f"Input: [yellow]{input_pdf}[/]\n"
        f"Output: [green]{output_docx}[/]\n"
        f"Mode: [magenta]{mode.upper()}[/]",
        title="Document Conversion"
    ))

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TimeElapsedColumn(),
        console=console,
    ) as progress:
        task = progress.add_task("[cyan]Converting PDF...", total=100)

        def on_progress(cur, total, msg):
            pct = int((cur / max(1, total)) * 100)
            progress.update(task, completed=pct, description=f"[cyan]{msg}")

        converter = OmniConverter(
            convert_math=math,
            highlight_code=code,
            extract_images=images,
            image_dpi=dpi,
        )

        try:
            mode_enum = ConversionMode(mode.lower())
            result = converter.convert(
                pdf_path=input_pdf,
                docx_path=output_docx,
                mode=mode_enum,
                on_progress=on_progress,
            )
            progress.update(task, completed=100, description="[bold green]Conversion Complete!")
        except Exception as e:
            console.print(f"\n[bold red]Conversion failed:[/] {e}")
            raise typer.Exit(code=1)

    # Print summary
    table = Table(title="Conversion Summary", border_style="cyan")
    table.add_column("Metric", style="bold")
    table.add_column("Value", style="green")

    table.add_row("Status", result.get("status", "success"))
    table.add_row("Mode Used", result.get("mode_used", mode))
    table.add_row("Document Type", result.get("document_type", "N/A"))
    table.add_row("Pages Converted", str(result.get("total_pages", 0)))
    table.add_row("Processing Time", f"{result.get('duration_seconds', 0)}s")
    table.add_row("Output Size", f"{result.get('output_size_kb', 0)} KB")

    stats = result.get("stats", {})
    if "equations_converted" in stats:
        table.add_row("Equations Converted", str(stats.get("equations_converted", 0)))
    if "code_blocks" in stats:
        table.add_row("Code Blocks", str(stats.get("code_blocks", 0)))
    if "tables_extracted" in stats:
        table.add_row("Tables Extracted", str(stats.get("tables_extracted", 0)))
    if "figures_embedded" in stats:
        table.add_row("Figures Embedded", str(stats.get("figures_embedded", 0)))

    console.print(table)
    console.print(f"\n✨ [bold green]Successfully saved to:[/] [underline]{output_docx}[/]\n")


@app.command("inspect")
def inspect(
    input_pdf: str = typer.Argument(..., help="Path to input PDF file"),
    as_json: bool = typer.Option(False, "--json", help="Output inspection report as raw JSON"),
):
    """Inspect PDF structure, layout, columns, formulas, tables, and compatibility."""
    if not os.path.exists(input_pdf):
        console.print(f"[bold red]Error:[/] PDF file not found: {input_pdf}")
        raise typer.Exit(code=1)

    inspector = PDFInspector(input_pdf)
    report = inspector.inspect()

    if as_json:
        print(json.dumps(report.to_dict(), indent=2))
        return

    table = Table(title=f"PDF Inspection: {report.file_name}", border_style="cyan")
    table.add_column("Property", style="bold")
    table.add_column("Details", style="yellow")

    table.add_row("File Size", f"{round(report.file_size_bytes / 1024, 2)} KB")
    table.add_row("Total Pages", str(report.total_pages))
    table.add_row("Detected Columns", str(report.detected_columns))
    table.add_row("Document Type", report.document_type.replace("_", " ").title())
    table.add_row("Math Density", report.math_density.title())
    table.add_row("Code Blocks Detected", str(report.code_blocks_detected))
    table.add_row("Tables Detected", str(report.tables_detected))
    table.add_row("Images Detected", str(report.images_detected))
    table.add_row("Drawings / Vector Art", str(report.vector_drawings_detected))
    table.add_row("Recommended Mode", f"[bold green]{report.recommended_mode.upper()}[/]")
    table.add_row("Compatibility Score", f"{report.compatibility_score}/100")

    console.print(table)

    if report.notes:
        console.print("\n[bold]Notes:[/]")
        for note in report.notes:
            console.print(f"  • {note}")


@app.command("batch")
def batch(
    input_dir: str = typer.Argument(..., help="Directory containing PDF files"),
    output_dir: str = typer.Option(None, "-o", "--output", help="Directory to save converted DOCX files"),
    mode: str = typer.Option("auto", "-m", "--mode", help="Conversion mode: academic, exact, flow, auto"),
):
    """Batch convert all PDF files in a directory."""
    if not os.path.isdir(input_dir):
        console.print(f"[bold red]Error:[/] Input directory not found: {input_dir}")
        raise typer.Exit(code=1)

    if not output_dir:
        output_dir = os.path.join(input_dir, "converted_docx")
    os.makedirs(output_dir, exist_ok=True)

    pdf_files = [f for f in os.listdir(input_dir) if f.lower().endswith(".pdf")]
    if not pdf_files:
        console.print(f"[yellow]No PDF files found in {input_dir}[/]")
        return

    console.print(f"[bold cyan]Found {len(pdf_files)} PDF files to convert.[/]\n")

    converter = OmniConverter()
    mode_enum = ConversionMode(mode.lower())

    for idx, f in enumerate(pdf_files, 1):
        in_path = os.path.join(input_dir, f)
        out_name = f"{os.path.splitext(f)[0]}.docx"
        out_path = os.path.join(output_dir, out_name)

        console.print(f"[{idx}/{len(pdf_files)}] Converting [yellow]{f}[/]...")
        try:
            res = converter.convert(in_path, out_path, mode=mode_enum)
            console.print(f"  ✓ Saved to [green]{out_name}[/] ({res.get('duration_seconds')}s)")
        except Exception as e:
            console.print(f"  ✗ [red]Failed:[/] {e}")

    console.print(f"\n✨ [bold green]Batch conversion finished. Output directory:[/] {output_dir}")


@app.command("serve")
def serve(
    host: str = typer.Option("127.0.0.1", "--host", help="Host address for web server"),
    port: int = typer.Option(8000, "--port", "-p", help="Port for web server"),
):
    """Launch the interactive FarazanPDFConverter Web Application."""
    import uvicorn
    console.print(Panel.fit(
        f"[bold green]Starting FarazanPDFConverter Web Interface[/]\n"
        f"URL: [bold underline cyan]http://{host}:{port}[/]\n"
        f"Press Ctrl+C to stop.",
        title="FarazanPDFConverter Web Server"
    ))
    uvicorn.run("omnipdf.web.app:app", host=host, port=port, reload=False)


def main():
    app()


if __name__ == "__main__":
    main()
