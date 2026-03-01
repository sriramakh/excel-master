"""Main dashboard engine: orchestrates data profiling → LLM selection → template rendering."""
from __future__ import annotations
import json
from pathlib import Path
import pandas as pd

from ..models import DashboardConfig, DashboardTemplate, BuildResult
from ..data.data_engine import profile_dataset, discover_and_join
from .template_selector import TemplateSelector
from .themes import get_theme
from .deep_analysis import compute_deep_stats, build_analysis_prompt, safe_parse_analysis


# Template class registry — uses new xlsxwriter-based templates
def _get_template_class(template: DashboardTemplate):
    from .templates.executive_xl import ExecutiveSummaryXL
    from .templates.hr_xl import HRAnalyticsXL
    from .templates.dark_operational_xl import DarkOperationalXL
    from .templates.financial_xl import FinancialXL
    from .templates.supply_chain_xl import SupplyChainXL
    from .templates.marketing_xl import MarketingXL
    from .templates.minimal_clean_xl import MinimalCleanXL

    mapping = {
        DashboardTemplate.EXECUTIVE_SUMMARY: ExecutiveSummaryXL,
        DashboardTemplate.HR_ANALYTICS: HRAnalyticsXL,
        DashboardTemplate.DARK_OPERATIONAL: DarkOperationalXL,
        DashboardTemplate.FINANCIAL: FinancialXL,
        DashboardTemplate.SUPPLY_CHAIN: SupplyChainXL,
        DashboardTemplate.MARKETING: MarketingXL,
        DashboardTemplate.MINIMAL_CLEAN: MinimalCleanXL,
    }
    return mapping.get(template, ExecutiveSummaryXL)


class DashboardEngine:
    """
    Main orchestrator for the Excel Master pipeline:
    1. Profile the dataset (or use existing profile)
    2. Ask LLM to select template + configure dashboard
    3. Render the dashboard using selected template
    4. Save to output path
    """

    def __init__(self, use_llm: bool = True):
        self.use_llm = use_llm
        self.selector = TemplateSelector() if use_llm else None

    def build(self, data_path: str | Path,
               output_path: str | Path | None = None,
               sheet_name: str | None = None,
               industry: str = "",
               template_override: str | None = None,
               theme_override: str | None = None,
               verbose: bool = True) -> BuildResult:
        """
        Full pipeline: data → profile → LLM config → render → save.

        Args:
            data_path: Path to input Excel (.xlsx) file
            output_path: Where to save the dashboard (auto-named if None)
            sheet_name: Which sheet to read (auto-detected if None)
            industry: Industry hint for LLM (auto-detected if empty)
            template_override: Force a specific template key
            theme_override: Force a specific theme key
            verbose: Print progress messages
        """
        data_path = Path(data_path)
        if not data_path.exists():
            return BuildResult(success=False, output_path="",
                               dataset=str(data_path), template_used="",
                               kpi_count=0, chart_count=0,
                               error=f"File not found: {data_path}")

        # 1. Set output path
        if output_path is None:
            out_dir = data_path.parent.parent / "output"
            output_path = out_dir / f"{data_path.stem}_dashboard.xlsx"
        output_path = Path(output_path)

        try:
            # 2. Profile dataset
            if verbose:
                print(f"\n{'='*60}")
                print(f"Processing: {data_path.name}")
                print(f"{'='*60}")
                print("  Profiling dataset...")

            profile = profile_dataset(data_path, sheet_name=sheet_name, industry=industry)
            if verbose:
                print(f"  Profile: {profile.rows:,} rows × {len(profile.columns)} cols | "
                      f"Sheet: {profile.sheet_name}")

            # 3. LLM configuration
            if self.use_llm:
                config = self.selector.select_with_override(
                    profile,
                    template=template_override,
                    theme=theme_override,
                )
            else:
                config = self._default_config(profile, template_override, theme_override)

            # 4. Load data (multi-sheet files already unified during profiling)
            if data_path.suffix.lower() == ".csv":
                if verbose:
                    print(f"  Loading CSV data...")
                df = pd.read_csv(data_path)
            elif len(pd.ExcelFile(data_path).sheet_names) > 1 and sheet_name is None:
                if verbose:
                    print(f"  Loading unified multi-sheet data...")
                df, _, _ = discover_and_join(data_path, verbose=False)
            else:
                if verbose:
                    print(f"  Loading data from sheet: {profile.sheet_name}...")
                df = pd.read_excel(data_path, sheet_name=profile.sheet_name)

            # 4.5 Deep Analysis — LLM-powered statistical interpretation
            try:
                if verbose:
                    print(f"  Computing deep analysis statistics...")
                stats = compute_deep_stats(df, profile)
                sys_prompt, user_prompt = build_analysis_prompt(stats, profile, config)
                if self.use_llm:
                    if verbose:
                        print(f"  Calling LLM for deep analysis...")
                    raw = self.selector.llm.generate_json(
                        sys_prompt, user_prompt, max_tokens_override=6144,
                    )
                    config.deep_analysis = safe_parse_analysis(raw)
                    if verbose:
                        print(f"  ✓ Deep analysis received "
                              f"(quality score: {config.deep_analysis.data_quality_score}/100)")
            except Exception as e:
                if verbose:
                    print(f"  ⚠ Deep analysis skipped: {e}")
                config.deep_analysis = None

            # 5. Render dashboard
            if verbose:
                print(f"  Rendering [{config.template.value}] template with [{config.theme.value}] theme...")
            template_cls = _get_template_class(config.template)
            template_instance = template_cls(config)
            template_instance._store_for_analysis(df, profile)
            saved_path = template_instance.build(df, output_path)

            if verbose:
                print(f"  ✓ Dashboard saved: {saved_path}")
                print(f"  ✓ KPIs: {len(config.kpis)} | Charts: {len(config.charts)}")

            return BuildResult(
                success=True,
                output_path=str(saved_path),
                dataset=data_path.stem,
                template_used=config.template.value,
                kpi_count=len(config.kpis),
                chart_count=len(config.charts),
            )

        except Exception as e:
            import traceback
            err = traceback.format_exc()
            if verbose:
                print(f"  ✗ Error: {e}")
            return BuildResult(
                success=False,
                output_path=str(output_path),
                dataset=data_path.stem,
                template_used=template_override or "auto",
                kpi_count=0,
                chart_count=0,
                error=str(e),
            )

    def build_from_dataframe(self, df: pd.DataFrame,
                              name: str = "dataset",
                              output_path: str | Path = "output/dashboard.xlsx",
                              industry: str = "",
                              template_override: str | None = None) -> BuildResult:
        """Build dashboard directly from a pandas DataFrame."""
        import tempfile, os
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name
        try:
            df.to_excel(tmp_path, index=False)
            return self.build(tmp_path, output_path=output_path,
                               industry=industry, template_override=template_override)
        finally:
            os.unlink(tmp_path)

    def build_all(self, data_dir: str | Path = "data",
                   output_dir: str | Path = "output",
                   verbose: bool = True) -> list[BuildResult]:
        """Build dashboards for all Excel files in data_dir."""
        data_dir = Path(data_dir)
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        xlsx_files = sorted(data_dir.glob("*.xlsx"))
        csv_files = sorted(data_dir.glob("*.csv"))
        xlsx_files = list(xlsx_files) + list(csv_files)
        if verbose:
            print(f"\nFound {len(xlsx_files)} data files in {data_dir}")

        results = []
        for f in xlsx_files:
            out = output_dir / f"{f.stem}_dashboard.xlsx"
            result = self.build(f, output_path=out, verbose=verbose)
            results.append(result)

        # Summary
        success = sum(1 for r in results if r.success)
        if verbose:
            print(f"\n{'='*60}")
            print(f"Build complete: {success}/{len(results)} dashboards generated")
            for r in results:
                status = "✓" if r.success else "✗"
                print(f"  {status} {r.dataset} → {r.template_used}"
                      f" ({r.kpi_count} KPIs, {r.chart_count} charts)")

        return results

    def _default_config(self, profile, template_override=None, theme_override=None) -> DashboardConfig:
        """Fallback config without LLM."""
        from ..models import (DashboardConfig, DashboardTemplate, ColorTheme,
                               KPIConfig, ChartConfig, AggFunc, ChartType, NumberFormat)
        from .themes import TEMPLATE_DEFAULT_THEME

        # Pick template based on industry keywords
        industry = profile.industry.lower()
        if "hr" in industry or "human" in industry:
            tpl = DashboardTemplate.HR_ANALYTICS
        elif "finance" in industry or "financial" in industry:
            tpl = DashboardTemplate.FINANCIAL
        elif "supply" in industry or "logistic" in industry:
            tpl = DashboardTemplate.SUPPLY_CHAIN
        elif "marketing" in industry:
            tpl = DashboardTemplate.MARKETING
        elif "executive" in industry or "board" in industry:
            tpl = DashboardTemplate.EXECUTIVE_SUMMARY
        else:
            tpl = DashboardTemplate.MINIMAL_CLEAN

        if template_override:
            try:
                tpl = DashboardTemplate(template_override)
            except ValueError:
                pass

        theme = TEMPLATE_DEFAULT_THEME.get(tpl)
        if theme_override:
            try:
                theme = ColorTheme(theme_override)
            except ValueError:
                pass

        # Auto KPIs
        kpis = []
        for col in profile.numeric_columns[:4]:
            kpis.append(KPIConfig(label=col.replace("_", " ").title(),
                                   column=col, aggregation=AggFunc.SUM,
                                   format=NumberFormat.NUMBER))

        # Auto charts
        charts = []
        if profile.categorical_columns and profile.numeric_columns:
            charts.append(ChartConfig(
                type=ChartType.BAR,
                title=f"{profile.categorical_columns[0].replace('_', ' ').title()} Analysis",
                x_column=profile.categorical_columns[0],
                y_columns=[profile.numeric_columns[0]],
                aggregation=AggFunc.SUM,
            ))
        if profile.date_columns and profile.numeric_columns:
            charts.append(ChartConfig(
                type=ChartType.LINE,
                title="Trend Over Time",
                x_column=profile.date_columns[0],
                y_columns=[profile.numeric_columns[0]],
                aggregation=AggFunc.SUM,
            ))
        if len(profile.categorical_columns) >= 2:
            charts.append(ChartConfig(
                type=ChartType.PIE,
                title=f"{profile.categorical_columns[0]} Distribution",
                x_column=profile.categorical_columns[0],
                y_columns=[profile.numeric_columns[0]] if profile.numeric_columns else [],
                aggregation=AggFunc.COUNT,
                top_n=8,
            ))

        return DashboardConfig(
            template=tpl,
            title=f"{profile.name.replace('_', ' ').title()} Dashboard",
            subtitle=profile.description,
            theme=theme,
            kpis=kpis,
            charts=charts,
            table_columns=list(profile.column_names[:6]),
        )
