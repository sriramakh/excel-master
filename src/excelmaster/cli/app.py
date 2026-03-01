"""Excel Master CLI — generate data, build dashboards, run full pipeline."""
from __future__ import annotations
import sys
from pathlib import Path
from typing import Optional
import typer
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich import print as rprint

app = typer.Typer(
    name="excelmaster",
    help="AI-powered Excel Dashboard Engine",
    add_completion=False,
)
console = Console()

DATASET_TYPES = [
    "extreme_load", "moderate", "feature_rich", "sparse",
    "finance", "supply_chain", "executive", "hr_admin", "marketing"
]

TEMPLATE_TYPES = [
    "executive_summary", "hr_analytics", "dark_operational",
    "financial", "supply_chain", "marketing", "minimal_clean"
]

THEME_TYPES = [
    "corporate_blue", "hr_purple", "dark_mode", "supply_green",
    "finance_green", "marketing_orange", "slate_minimal", "executive_navy"
]


# ── generate-data ──────────────────────────────────────────────────────────────

@app.command("generate-data")
def generate_data(
    dataset_type: str = typer.Argument(
        ...,
        help=f"Dataset type. One of: {', '.join(DATASET_TYPES)}, or 'all'"
    ),
    output_dir: Path = typer.Option(
        Path("data"), "--output-dir", "-o",
        help="Directory to save generated Excel files"
    ),
):
    """Generate synthetic dataset(s) as Excel files."""
    from ..data.data_engine import generate_dataset, generate_all

    output_dir.mkdir(parents=True, exist_ok=True)

    if dataset_type == "all":
        console.print(Panel(
            "[bold cyan]Generating all 9 datasets...[/bold cyan]",
            title="Excel Master Data Engine"
        ))
        results = generate_all(output_dir)
        _print_data_results(results)
    else:
        if dataset_type not in DATASET_TYPES:
            console.print(f"[red]Unknown dataset type: {dataset_type}[/red]")
            console.print(f"Available: {', '.join(DATASET_TYPES)}")
            raise typer.Exit(1)
        console.print(f"[cyan]Generating [{dataset_type}] dataset...[/cyan]")
        path = generate_dataset(dataset_type, output_dir)
        console.print(f"[green]✓ Saved: {path}[/green]")


def _print_data_results(results: dict) -> None:
    table = Table(title="Generated Datasets", show_header=True)
    table.add_column("Dataset", style="cyan")
    table.add_column("Status", style="green")
    table.add_column("Path")
    for ds_type, path in results.items():
        table.add_row(ds_type, "✓ Done", str(path))
    console.print(table)


# ── build-dashboard ────────────────────────────────────────────────────────────

@app.command("build-dashboard")
def build_dashboard(
    data_path: Path = typer.Argument(..., help="Path to input Excel (.xlsx) file"),
    output_path: Optional[Path] = typer.Option(
        None, "--output", "-o",
        help="Output dashboard path (auto-named if not specified)"
    ),
    sheet: Optional[str] = typer.Option(
        None, "--sheet", "-s",
        help="Sheet name to use as data source (auto-detected if not specified)"
    ),
    industry: str = typer.Option(
        "", "--industry", "-i",
        help="Industry hint for LLM (e.g. 'Finance', 'HR', 'Marketing')"
    ),
    template: Optional[str] = typer.Option(
        None, "--template", "-t",
        help=f"Force template. Options: {', '.join(TEMPLATE_TYPES)}"
    ),
    theme: Optional[str] = typer.Option(
        None, "--theme",
        help=f"Force theme. Options: {', '.join(THEME_TYPES)}"
    ),
    no_llm: bool = typer.Option(
        False, "--no-llm",
        help="Use rule-based template selection instead of LLM"
    ),
):
    """Build an AI-powered Excel dashboard from an existing data file."""
    from ..dashboard.dashboard_engine import DashboardEngine

    if not data_path.exists():
        console.print(f"[red]File not found: {data_path}[/red]")
        raise typer.Exit(1)

    console.print(Panel(
        f"[bold]Input:[/bold] {data_path}\n"
        f"[bold]LLM:[/bold] {'Disabled (rule-based)' if no_llm else 'Enabled'}\n"
        f"[bold]Template:[/bold] {template or 'Auto (LLM-selected)'}\n"
        f"[bold]Theme:[/bold] {theme or 'Auto (LLM-selected)'}",
        title="[cyan]Excel Master Dashboard Engine[/cyan]"
    ))

    engine = DashboardEngine(use_llm=not no_llm)
    result = engine.build(
        data_path=data_path,
        output_path=output_path,
        sheet_name=sheet,
        industry=industry,
        template_override=template,
        theme_override=theme,
        verbose=True,
    )

    if result.success:
        console.print(Panel(
            f"[green]✓ Dashboard built successfully![/green]\n\n"
            f"[bold]Output:[/bold] {result.output_path}\n"
            f"[bold]Template:[/bold] {result.template_used}\n"
            f"[bold]KPIs:[/bold] {result.kpi_count} | "
            f"[bold]Charts:[/bold] {result.chart_count}",
            title="[green]Success[/green]"
        ))
    else:
        console.print(Panel(
            f"[red]✗ Build failed[/red]\n\n{result.error}",
            title="[red]Error[/red]"
        ))
        raise typer.Exit(1)


# ── run (full pipeline) ────────────────────────────────────────────────────────

@app.command("run")
def run_pipeline(
    dataset_type: str = typer.Argument(
        ...,
        help=f"Dataset type. One of: {', '.join(DATASET_TYPES)}, or 'all'"
    ),
    data_dir: Path = typer.Option(Path("data"), "--data-dir"),
    output_dir: Path = typer.Option(Path("output"), "--output-dir", "-o"),
    template: Optional[str] = typer.Option(None, "--template", "-t"),
    theme: Optional[str] = typer.Option(None, "--theme"),
    no_llm: bool = typer.Option(False, "--no-llm"),
):
    """Full pipeline: generate dataset → profile → LLM select → build dashboard."""
    from ..data.data_engine import generate_dataset, generate_all
    from ..dashboard.dashboard_engine import DashboardEngine

    data_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    console.print(Panel(
        f"[bold cyan]Excel Master Full Pipeline[/bold cyan]\n"
        f"Dataset: {dataset_type} | LLM: {'Off' if no_llm else 'On'}",
        title="Excel Master"
    ))

    engine = DashboardEngine(use_llm=not no_llm)

    if dataset_type == "all":
        console.print("\n[cyan]Step 1: Generating all datasets...[/cyan]")
        files = generate_all(data_dir)

        console.print("\n[cyan]Step 2: Building dashboards for all datasets...[/cyan]")
        results = []
        for ds_type, data_path in files.items():
            out = output_dir / f"{ds_type}_dashboard.xlsx"
            r = engine.build(data_path, output_path=out,
                               template_override=template,
                               theme_override=theme, verbose=True)
            results.append(r)

        _print_pipeline_summary(results)
    else:
        if dataset_type not in DATASET_TYPES:
            console.print(f"[red]Unknown: {dataset_type}[/red]")
            raise typer.Exit(1)

        console.print(f"\n[cyan]Step 1: Generating [{dataset_type}] dataset...[/cyan]")
        data_path = generate_dataset(dataset_type, data_dir)

        console.print(f"\n[cyan]Step 2: Building dashboard...[/cyan]")
        out = output_dir / f"{dataset_type}_dashboard.xlsx"
        result = engine.build(data_path, output_path=out,
                               template_override=template,
                               theme_override=theme, verbose=True)

        if result.success:
            console.print(f"\n[green]✓ Pipeline complete! Dashboard: {result.output_path}[/green]")
        else:
            console.print(f"\n[red]✗ Pipeline failed: {result.error}[/red]")
            raise typer.Exit(1)


def _print_pipeline_summary(results: list) -> None:
    table = Table(title="Pipeline Results", show_header=True)
    table.add_column("Dataset", style="cyan", no_wrap=True)
    table.add_column("Status")
    table.add_column("Template")
    table.add_column("KPIs")
    table.add_column("Charts")
    table.add_column("Output", overflow="fold")

    for r in results:
        status = "[green]✓[/green]" if r.success else "[red]✗[/red]"
        table.add_row(
            r.dataset, status, r.template_used,
            str(r.kpi_count), str(r.chart_count),
            Path(r.output_path).name if r.output_path else r.error[:30]
        )
    console.print(table)


# ── list ───────────────────────────────────────────────────────────────────────

@app.command("list")
def list_options():
    """List all available datasets, templates, and themes."""
    ds_table = Table(title="Available Datasets")
    ds_table.add_column("Type", style="cyan")
    ds_table.add_column("Industry")
    ds_table.add_column("Description")

    descriptions = {
        "extreme_load": ("Multi-Industry", "5 datasets, 10K+ rows, 15+ cols each"),
        "moderate": ("E-Commerce", "Orders data, 2,500 rows, 14 cols"),
        "feature_rich": ("Investment", "Portfolio analytics, 22+ calculated cols"),
        "sparse": ("Research", "Survey data with 35-60% null rates"),
        "finance": ("Finance", "P&L, cash flow, accounts receivable"),
        "supply_chain": ("Logistics", "Shipments, carriers, inventory"),
        "executive": ("C-Suite", "KPI scorecard, OKRs, quarterly trends"),
        "hr_admin": ("HR", "Employee master, payroll, recruitment"),
        "marketing": ("Marketing", "Campaigns, web analytics, content"),
    }
    for ds, (ind, desc) in descriptions.items():
        ds_table.add_row(ds, ind, desc)
    console.print(ds_table)

    tpl_table = Table(title="Available Templates")
    tpl_table.add_column("Template", style="green")
    tpl_table.add_column("Best For")
    tpl_table.add_column("Default Theme")

    templates_info = {
        "executive_summary": ("Board/C-Suite KPIs", "corporate_blue"),
        "hr_analytics": ("People & Workforce", "hr_purple"),
        "dark_operational": ("Dense Operational Data", "dark_mode"),
        "financial": ("P&L, Budget, Cash Flow", "finance_green"),
        "supply_chain": ("Logistics & Freight", "supply_green"),
        "marketing": ("Campaigns & ROI", "marketing_orange"),
        "minimal_clean": ("Research, Survey, General", "slate_minimal"),
    }
    for tpl, (best, theme) in templates_info.items():
        tpl_table.add_row(tpl, best, theme)
    console.print(tpl_table)


# ── profile ────────────────────────────────────────────────────────────────────

@app.command("profile")
def profile_data(
    data_path: Path = typer.Argument(..., help="Path to Excel file to profile"),
    sheet: Optional[str] = typer.Option(None, "--sheet", "-s"),
):
    """Profile a dataset and show its structure (no dashboard generation)."""
    from ..data.data_engine import profile_dataset, profile_to_prompt_text

    if not data_path.exists():
        console.print(f"[red]File not found: {data_path}[/red]")
        raise typer.Exit(1)

    profile = profile_dataset(data_path, sheet_name=sheet)
    console.print(Panel(
        profile_to_prompt_text(profile),
        title=f"[cyan]Dataset Profile: {profile.name}[/cyan]"
    ))


# ── chat ──────────────────────────────────────────────────────────────────────

@app.command("chat")
def chat_command(
    data_path: Path = typer.Argument(..., help="Path to CSV or Excel data file"),
    output_path: Optional[Path] = typer.Option(
        None, "--output", "-o",
        help="Output dashboard path (auto-named if not specified)"
    ),
):
    """Interactive chat: build and modify dashboards with natural language."""
    from ..chat.engine import ChatEngine

    if not data_path.exists():
        console.print(f"[red]File not found: {data_path}[/red]")
        raise typer.Exit(1)

    engine = ChatEngine(data_path, output_path)
    engine.run()


def main():
    app()


if __name__ == "__main__":
    main()
