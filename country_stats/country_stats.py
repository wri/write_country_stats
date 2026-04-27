import os
import sys
import pandas as pd
import click
import requests
from tqdm import tqdm


GFW_API_BASE = "https://data-api.globalforestwatch.org/dataset"

# Carbon tabs only carry data at thresholds >= 30 — drop the lower thresholds.
CARBON_EXCLUDED_THRESHOLDS = {0, 10, 15, 20, 25}


# --------------------------------------------------------------------------------------
# Low-level API plumbing
# --------------------------------------------------------------------------------------

def _api_query(dataset, version, sql, api_key):
    """Run a SQL query against a GFW Data API dataset and return the rows as a DataFrame."""
    url = f"{GFW_API_BASE}/{dataset}/{version}/query/json"
    resp = requests.get(url, params={"sql": sql, "x-api-key": api_key}, timeout=120)
    resp.raise_for_status()
    return pd.DataFrame(resp.json()["data"])


def _run_per_iso(dataset, version, sql_template, api_key, isos, desc):
    """Run a `{iso}`-templated SQL once per ISO; return concatenated rows (empty df if none)."""
    frames = []
    for iso in tqdm(isos, desc=desc, unit="iso"):
        try:
            df = _api_query(dataset, version, sql_template.format(iso=iso), api_key)
        except Exception as exc:
            tqdm.write(f"  [warn] {iso}: {exc}")
            continue
        if not df.empty:
            frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


# --------------------------------------------------------------------------------------
# SQL builders for gadm__tcl__iso_change / gadm__tcl__adm1_change
# --------------------------------------------------------------------------------------

def _build_loss_sql(level, by_driver, primary, iso=None):
    """
    Build a tree-cover-loss SUM(...) GROUP BY query.

    - level: "iso" or "adm1"
    - by_driver: include wri_google_tree_cover_loss_drivers__driver in SELECT/GROUP BY
    - primary: add `is__umd_regional_primary_forest_2001 = true` to WHERE
    - iso: optional per-ISO filter (use a `'{iso}'` placeholder if you need to .format() later)
    """
    grouping = ["iso"]
    if level == "adm1":
        grouping.append("adm1")
    grouping.append("umd_tree_cover_loss__year")
    if by_driver:
        grouping.append("wri_google_tree_cover_loss_drivers__driver")

    where = ["umd_tree_cover_density_2000__threshold = 30"]
    if primary:
        where.append("is__umd_regional_primary_forest_2001 = true")
    if iso is not None:
        where.append(f"iso = '{iso}'")

    select_cols = ", ".join(grouping) + ", SUM(umd_tree_cover_loss__ha) AS umd_tree_cover_loss__ha"
    group_by = ", ".join(grouping)
    return (
        f"SELECT {select_cols} FROM data WHERE {' AND '.join(where)} "
        f"GROUP BY {group_by} ORDER BY {group_by}"
    )


def _build_area_sql(primary):
    """Per-ISO area summary against gadm__tcl__iso_summary."""
    where = ["umd_tree_cover_density_2000__threshold = 30"]
    if primary:
        where.append("is__umd_regional_primary_forest_2001 = true")
    return (
        "SELECT iso, SUM(area__ha) AS area__ha FROM data "
        f"WHERE {' AND '.join(where)} GROUP BY iso ORDER BY iso"
    )


# --------------------------------------------------------------------------------------
# GADM name attach helpers — rename the gadm columns to the join keys to avoid the
# "country" name collision the API responses would otherwise create.
# --------------------------------------------------------------------------------------

def _attach_iso_name(df, gadm_df):
    """Inner-merge `iso_name` onto `df` keyed on iso code. Drops rows whose iso is not in gadm."""
    name_lookup = (
        gadm_df[["country", "iso_name"]].drop_duplicates().rename(columns={"country": "iso"})
    )
    return df.merge(name_lookup, on="iso", how="inner")


def _attach_adm1_name(df, gadm_df):
    """Inner-merge `iso_name` + `adm1_name` onto `df` keyed on (iso, adm1)."""
    df = df.dropna(subset=["adm1"]).copy()
    df["adm1"] = df["adm1"].astype("Int64")
    name_lookup = (
        gadm_df[["country", "iso_name", "subnational1", "adm1_name"]]
        .drop_duplicates()
        .rename(columns={"country": "iso", "subnational1": "adm1"})
    )
    name_lookup["adm1"] = name_lookup["adm1"].astype("Int64")
    return df.merge(name_lookup, on=["iso", "adm1"], how="inner")


# --------------------------------------------------------------------------------------
# Output shapers (raw API rows → tab-ready frames)
# --------------------------------------------------------------------------------------

def _shape_primary_loss_iso(loss_df, area_df, gadm_df, years):
    cols = ["country", "threshold", "area__ha"] + [f"tc_loss_ha_{y}" for y in years]
    if loss_df.empty:
        return pd.DataFrame(columns=cols)

    wide = loss_df.pivot_table(
        index="iso",
        columns="umd_tree_cover_loss__year",
        values="umd_tree_cover_loss__ha",
        fill_value=0,
    ).reset_index()

    for y in years:
        if y not in wide.columns:
            wide[y] = 0

    wide["threshold"] = 30
    wide = wide.merge(area_df, on="iso", how="left")
    wide = _attach_iso_name(wide, gadm_df)

    out = wide[["iso_name", "threshold", "area__ha"] + list(years)].copy()
    out = out.rename(columns={"iso_name": "country", **{y: f"tc_loss_ha_{y}" for y in years}})

    numeric_cols = ["area__ha"] + [f"tc_loss_ha_{y}" for y in years]
    out[numeric_cols] = out[numeric_cols].fillna(0).round().astype("Int64")
    return out.sort_values(by=["country"]).reset_index(drop=True)


def _shape_primary_loss_adm1(loss_df, gadm_df, years):
    cols = ["country", "subnational1", "threshold"] + [f"tc_loss_ha_{y}" for y in years]
    if loss_df.empty:
        return pd.DataFrame(columns=cols)

    wide = loss_df.pivot_table(
        index=["iso", "adm1"],
        columns="umd_tree_cover_loss__year",
        values="umd_tree_cover_loss__ha",
        fill_value=0,
    ).reset_index()

    for y in years:
        if y not in wide.columns:
            wide[y] = 0

    merged = _attach_adm1_name(wide, gadm_df)
    merged["threshold"] = 30

    out = merged[["iso_name", "adm1_name", "threshold"] + list(years)].copy()
    out = out.rename(
        columns={
            "iso_name": "country",
            "adm1_name": "subnational1",
            **{y: f"tc_loss_ha_{y}" for y in years},
        }
    )
    numeric_cols = [f"tc_loss_ha_{y}" for y in years]
    out[numeric_cols] = out[numeric_cols].fillna(0).round().astype("Int64")
    return out.sort_values(by=["country", "subnational1"]).reset_index(drop=True)


def _shape_drivers(df, gadm_df, level):
    """Long-format driver tab. level in {'iso', 'adm1'}."""
    base_cols = ["country", "threshold", "driver", "year", "tc_loss_ha"]
    if level == "adm1":
        base_cols.insert(1, "subnational1")
    if df.empty:
        return pd.DataFrame(columns=base_cols)

    rename = {
        "iso_name": "country",
        "wri_google_tree_cover_loss_drivers__driver": "driver",
        "umd_tree_cover_loss__year": "year",
        "umd_tree_cover_loss__ha": "tc_loss_ha",
    }

    if level == "iso":
        df = _attach_iso_name(df, gadm_df)
        df["threshold"] = 30
        out = df[
            [
                "iso_name",
                "threshold",
                "wri_google_tree_cover_loss_drivers__driver",
                "umd_tree_cover_loss__year",
                "umd_tree_cover_loss__ha",
            ]
        ].rename(columns=rename)
        sort_cols = ["country", "year", "driver"]
    else:  # adm1
        df = _attach_adm1_name(df, gadm_df)
        df["threshold"] = 30
        rename["adm1_name"] = "subnational1"
        out = df[
            [
                "iso_name",
                "adm1_name",
                "threshold",
                "wri_google_tree_cover_loss_drivers__driver",
                "umd_tree_cover_loss__year",
                "umd_tree_cover_loss__ha",
            ]
        ].rename(columns=rename)
        sort_cols = ["country", "subnational1", "year", "driver"]

    out["tc_loss_ha"] = out["tc_loss_ha"].fillna(0).round().astype("Int64")
    return out.sort_values(by=sort_cols).reset_index(drop=True)


# --------------------------------------------------------------------------------------
# Public fetchers — thin wrappers that pick the right SQL + shaper
# --------------------------------------------------------------------------------------

def fetch_iso_primary_loss(api_version, api_key, gadm_df, years):
    """Country primary-forest loss, threshold=30. Wide format (one column per year)."""
    sql_loss = _build_loss_sql(level="iso", by_driver=False, primary=True)
    sql_area = _build_area_sql(primary=True)
    with tqdm(total=2, desc="ISO primary loss", unit="query") as bar:
        loss = _api_query("gadm__tcl__iso_change", api_version, sql_loss, api_key)
        bar.update(1)
        area = _api_query("gadm__tcl__iso_summary", api_version, sql_area, api_key)
        bar.update(1)
    if loss.empty:
        click.echo("WARNING: ISO primary-loss query returned no rows.", err=True)
    return _shape_primary_loss_iso(loss, area, gadm_df, years)


def fetch_adm1_primary_loss(api_version, api_key, gadm_df, years):
    """Subnational-1 primary-forest loss, threshold=30. Wide format (one column per year)."""
    isos = sorted(gadm_df["country"].unique().tolist())
    sql_t = _build_loss_sql(level="adm1", by_driver=False, primary=True, iso="{iso}")
    raw = _run_per_iso(
        "gadm__tcl__adm1_change", api_version, sql_t, api_key, isos, "ADM1 primary loss"
    )
    if raw.empty:
        click.echo("WARNING: ADM1 primary-loss queries returned no rows.", err=True)
    return _shape_primary_loss_adm1(raw, gadm_df, years)


def fetch_iso_drivers(api_version, api_key, gadm_df):
    """Country tree-cover-loss by driver, threshold=30. Long format."""
    sql = _build_loss_sql(level="iso", by_driver=True, primary=False)
    with tqdm(total=1, desc="ISO drivers", unit="query") as bar:
        df = _api_query("gadm__tcl__iso_change", api_version, sql, api_key)
        bar.update(1)
    if df.empty:
        click.echo("WARNING: ISO drivers query returned no rows.", err=True)
    return _shape_drivers(df, gadm_df, level="iso")


def fetch_adm1_drivers(api_version, api_key, gadm_df):
    """Subnational-1 tree-cover-loss by driver, threshold=30. Long format."""
    isos = sorted(gadm_df["country"].unique().tolist())
    sql_t = _build_loss_sql(level="adm1", by_driver=True, primary=False, iso="{iso}")
    raw = _run_per_iso(
        "gadm__tcl__adm1_change", api_version, sql_t, api_key, isos, "ADM1 drivers"
    )
    if raw.empty:
        click.echo("WARNING: ADM1 drivers queries returned no rows.", err=True)
    return _shape_drivers(raw, gadm_df, level="adm1")


def fetch_iso_primary_drivers(api_version, api_key, gadm_df):
    """Country primary-forest loss by driver, threshold=30. Long format."""
    sql = _build_loss_sql(level="iso", by_driver=True, primary=True)
    with tqdm(total=1, desc="ISO primary drivers", unit="query") as bar:
        df = _api_query("gadm__tcl__iso_change", api_version, sql, api_key)
        bar.update(1)
    if df.empty:
        click.echo("WARNING: ISO primary drivers query returned no rows.", err=True)
    return _shape_drivers(df, gadm_df, level="iso")


def fetch_adm1_primary_drivers(api_version, api_key, gadm_df):
    """Subnational-1 primary-forest loss by driver, threshold=30. Long format."""
    isos = sorted(gadm_df["country"].unique().tolist())
    sql_t = _build_loss_sql(level="adm1", by_driver=True, primary=True, iso="{iso}")
    raw = _run_per_iso(
        "gadm__tcl__adm1_change", api_version, sql_t, api_key, isos, "ADM1 primary drivers"
    )
    if raw.empty:
        click.echo("WARNING: ADM1 primary drivers queries returned no rows.", err=True)
    return _shape_drivers(raw, gadm_df, level="adm1")


# --------------------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------------------

@click.command()
@click.argument("iso")
@click.argument("adm1")
@click.option("--output-dir", default=None, help="Output directory (default: <script_dir>/output)")
@click.option(
    "--api-version",
    default="v20260424",
    help="GFW Data API dataset version for primary-loss / drivers queries",
)
def cli(iso, adm1, output_dir, api_version):
    api_key = os.environ.get("GFW_API_KEY")
    if not api_key:
        click.echo("ERROR: GFW_API_KEY env var is not set.", err=True)
        sys.exit(1)

    path = os.path.dirname(os.path.abspath(__file__))
    out_dir = output_dir or os.path.join(path, "output")
    os.makedirs(out_dir, exist_ok=True)

    gadm_df = pd.read_csv(os.path.join(path, "gadm41.csv"))

    loss_years = list(range(2001, 2026))
    primary_years = list(range(2002, 2026))

    iso_df = (
        pd.merge(gadm_df, pd.read_csv(iso, sep="\t"), how="inner", on="country")
        .drop(["subnational1", "adm1_name", "subnational2", "adm2_name"], axis=1)
        .drop_duplicates()
        .sort_values(by=["iso_name", "umd_tree_cover_density_2000__threshold"])
        .reset_index(drop=True)
    )

    adm1_df = (
        pd.merge(gadm_df, pd.read_csv(adm1, sep="\t"), how="inner", on=["country", "subnational1"])
        .drop(["subnational1", "subnational2", "adm2_name"], axis=1)
        .drop_duplicates()
        .sort_values(by=["iso_name", "adm1_name", "umd_tree_cover_density_2000__threshold"])
        .reset_index(drop=True)
    )

    iso_primary_df = fetch_iso_primary_loss(api_version, api_key, gadm_df, primary_years)
    adm1_primary_df = fetch_adm1_primary_loss(api_version, api_key, gadm_df, primary_years)
    iso_drivers_df = fetch_iso_drivers(api_version, api_key, gadm_df)
    adm1_drivers_df = fetch_adm1_drivers(api_version, api_key, gadm_df)
    iso_primary_drivers_df = fetch_iso_primary_drivers(api_version, api_key, gadm_df)
    adm1_primary_drivers_df = fetch_adm1_primary_drivers(api_version, api_key, gadm_df)

    # Tab order: tcl → primary → drivers → primary drivers → carbon, for both Country and Subnational 1.
    write_to_excel(
        out_dir,
        "global",
        loss_years,
        primary_years,
        (iso_df, "Country tcl"),
        (iso_primary_df, "Country primary"),
        (iso_drivers_df, "Country drivers"),
        (iso_primary_drivers_df, "Country primary drivers"),
        (iso_df, "Country carbon"),
        (adm1_df, "Subnational 1 tcl"),
        (adm1_primary_df, "Subnational 1 primary"),
        (adm1_drivers_df, "Subnational 1 drivers"),
        (adm1_primary_drivers_df, "Subnational 1 primary drivers"),
        (adm1_df, "Subnational 1 carbon"),
    )

    click.echo(f"Wrote {os.path.join(out_dir, 'global.xlsx')}")


# --------------------------------------------------------------------------------------
# Excel writer — one specs table drives all 10 sheets
# --------------------------------------------------------------------------------------

def _build_sheet_specs(loss_years, primary_years):
    tc_col = ["umd_tree_cover_gain__ha"] + [f"umd_tree_cover_loss_{y}__ha" for y in loss_years]
    tc_col_alias = ["gain_2000-2012_ha"] + [f"tc_loss_ha_{y}" for y in loss_years]
    carbon_cols = [
        "umd_tree_cover_density_2000__threshold",
        "umd_tree_cover_extent_2000__ha",
        "gfw_aboveground_carbon_stocks_2000__Mg_C",
        "avg_gfw_aboveground_carbon_stocks_2000__Mg_C_ha-1",
        "gfw_forest_carbon_gross_emissions__Mg_CO2e_yr-1",
        "gfw_forest_carbon_gross_removals__Mg_CO2_yr-1",
        "gfw_forest_carbon_net_flux__Mg_CO2e_yr-1",
    ]
    carbon_emissions_yearly = [f"gfw_forest_carbon_gross_emissions_{y}__Mg_CO2e" for y in loss_years]

    area_stats = [
        "umd_tree_cover_density_2000__threshold",
        "area__ha",
        "umd_tree_cover_extent_2000__ha",
        "umd_tree_cover_extent_2010__ha",
    ]
    area_stats_alias = ["threshold", "area_ha", "extent_2000_ha", "extent_2010_ha"]

    primary_loss_cols = [f"tc_loss_ha_{y}" for y in primary_years]
    drivers_cols_country = ["country", "threshold", "driver", "year", "tc_loss_ha"]
    drivers_cols_subn1 = ["country", "subnational1", "threshold", "driver", "year", "tc_loss_ha"]

    return {
        "Country tcl": dict(
            sheet="Country tree cover loss",
            columns=["iso_name"] + area_stats + tc_col,
            header=["country"] + area_stats_alias + tc_col_alias,
        ),
        "Country primary": dict(
            sheet="Country primary loss",
            columns=["country", "threshold", "area__ha"] + primary_loss_cols,
        ),
        "Country drivers": dict(
            sheet="Country drivers",
            columns=drivers_cols_country,
        ),
        "Country primary drivers": dict(
            sheet="Country primary drivers",
            columns=drivers_cols_country,
        ),
        "Country carbon": dict(
            sheet="Country carbon data",
            columns=["iso_name"] + carbon_cols + carbon_emissions_yearly,
            header=["country"] + carbon_cols + carbon_emissions_yearly,
            filter_carbon=True,
        ),
        "Subnational 1 tcl": dict(
            sheet="Subnational 1 tree cover loss",
            columns=["iso_name", "adm1_name"] + area_stats + tc_col,
            header=["country", "subnational1"] + area_stats_alias + tc_col_alias,
        ),
        "Subnational 1 primary": dict(
            sheet="Subnational 1 primary loss",
            columns=["country", "subnational1", "threshold"] + primary_loss_cols,
        ),
        "Subnational 1 drivers": dict(
            sheet="Subnational 1 drivers",
            columns=drivers_cols_subn1,
        ),
        "Subnational 1 primary drivers": dict(
            sheet="Subnational 1 primary drivers",
            columns=drivers_cols_subn1,
        ),
        "Subnational 1 carbon": dict(
            sheet="Subnational 1 carbon data",
            columns=["iso_name", "adm1_name"] + carbon_cols + carbon_emissions_yearly,
            header=["country", "subnational1"] + carbon_cols + carbon_emissions_yearly,
            filter_carbon=True,
        ),
    }


def write_to_excel(out_dir, dataset, loss_years, primary_years, *dfs):
    specs = _build_sheet_specs(loss_years, primary_years)
    out_path = os.path.join(out_dir, f"{dataset}.xlsx")

    with pd.ExcelWriter(out_path) as writer:
        for frame, kind in dfs:
            if kind not in specs:
                raise ValueError(f"Unknown sheet kind: {kind!r}")
            spec = specs[kind]

            if spec.get("filter_carbon"):
                frame = frame[
                    ~frame["umd_tree_cover_density_2000__threshold"].isin(CARBON_EXCLUDED_THRESHOLDS)
                ]

            kwargs = dict(
                sheet_name=spec["sheet"],
                float_format="%.2f",
                freeze_panes=(1, 1),
                index=False,
                columns=spec["columns"],
            )
            if "header" in spec:
                kwargs["header"] = spec["header"]

            frame.to_excel(writer, **kwargs)


if __name__ == "__main__":
    cli()
