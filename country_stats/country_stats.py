import pandas as pd
import os
import click
import numpy as np


@click.command()
@click.argument("iso")
@click.argument("adm1")
@click.argument("adm2")
def cli(iso, adm1, adm2):
    path = os.path.dirname(__file__)
    gadm_df = pd.read_csv(os.path.join(path, "gadm36.csv"))

    iso_df = read_from_folder(iso)
    iso_df = pd.merge(gadm_df, iso_df, how='inner', on="country", copy=True).drop(['subnational1', "adm1_name", 'subnational2', "adm2_name"], axis=1).drop_duplicates().sort_values(by=['iso_name', "umd_tree_cover_density__threshold"])

    iso_df = iso_df.set_index(["country", "iso_name", "umd_tree_cover_density__threshold", "area__ha", "umd_tree_cover_gain_2000-2012__ha", "umd_tree_cover_extent_2000__ha", "umd_tree_cover_extent_2010__ha"] + ["umd_tree_cover_loss_{}__ha".format(year) for year in range(2001,2020)]).iloc[:, :].div(1000000, axis=0)
    iso_df = iso_df.reset_index()

    adm1_df = read_from_folder(adm1)
    adm1_df = pd.merge(gadm_df, adm1_df, how='inner', on=["country", "subnational1"], copy=True).drop(['subnational1', 'subnational2', "adm2_name"], axis=1).drop_duplicates().sort_values(by=['iso_name', "adm1_name", "umd_tree_cover_density__threshold"])

    adm1_df = adm1_df.set_index(["country", "iso_name", "adm1_name", "umd_tree_cover_density__threshold", "area__ha", "umd_tree_cover_gain_2000-2012__ha", "umd_tree_cover_extent_2000__ha", "umd_tree_cover_extent_2010__ha"] + ["umd_tree_cover_loss_{}__ha".format(year) for year in range(2001,2020)]).iloc[:, :].div(1000000, axis=0)
    adm1_df = adm1_df.reset_index()

    adm2_df = read_from_folder(adm2)
    adm2_df = pd.merge(gadm_df, adm2_df, how='inner', on=["country", "subnational1", "subnational2"], copy=True).drop(['subnational1', 'subnational2'], axis=1).drop_duplicates().sort_values(by=['iso_name', "adm1_name", "adm2_name", "umd_tree_cover_density__threshold"])

    adm2_df = adm2_df.set_index(["country", "iso_name", "adm1_name", "adm2_name", "umd_tree_cover_density__threshold", "area__ha", "umd_tree_cover_gain_2000-2012__ha", "umd_tree_cover_extent_2000__ha", "umd_tree_cover_extent_2010__ha"] + ["umd_tree_cover_loss_{}__ha".format(year) for year in range(2001,2020)]).iloc[:, :].div(1000000, axis=0)
    adm2_df = adm2_df.reset_index()

    countries = gadm_df.country.unique()

    write_to_excel("global",
                   (iso_df, "Country"),
                   (adm1_df, "Subnational 1"))

    for country in countries:
        write_to_excel(country,
                       (iso_df.loc[iso_df['country'] == country], "Country"),
                       (adm1_df.loc[adm1_df['country'] == country], "Subnational 1"),
                       (adm2_df.loc[adm2_df['country'] == country], "Subnational 2"))


def read_from_folder(folder):
    import glob

    all_files = glob.glob(os.path.join(folder, "*.csv"))

    li = []

    for filename in all_files:
        click.echo(filename)
        df = pd.read_csv(filename, index_col=None, header=0, sep="\t")
        li.append(df)

    return pd.concat(li, axis=0, ignore_index=True)


def write_to_excel(dataset, *dfs):
    tc_col = ["umd_tree_cover_gain_2000-2012__ha"] + ["umd_tree_cover_loss_{}__ha".format(year) for year in range(2001, 2021)]
    tc_col_alias = ["gain_2000-2012_ha"] + ["tc_loss_ha_{}".format(year) for year in range(2001,2020)]
    carbon_cols = [
        "umd_tree_cover_density__threshold",
        "umd_tree_cover_extent_2000__ha",
        "whrc_aboveground_biomass_stock_2000__Mg",
        "avg_whrc_aboveground_biomass_2000_Mg_ha-1",
        "gfw_gross_emissions_co2e_all_gases__Mg_yr-1",
        "gfw_gross_cumulative_aboveground_belowground_co2_removals__Mg_yr-1",
        "gfw_net_flux_co2e__Mg_yr-1",
    ]
    carbon_emissions_yearly = ["gfw_gross_emissions_co2e_all_gases__Mg__{}".format(year) for year in range(2001, 2021)]

    area_stats = ["umd_tree_cover_density__threshold", "area__ha", "umd_tree_cover_extent_2000__ha", "umd_tree_cover_extent_2010__ha"]
    area_stats_alias = ["threshold", "area_ha", "extent_2000_ha", "extent_2010_ha"]

    gadm = ["iso_name", "adm1_name", "adm2_name"]
    gadm_alias = ["country", "subnational1", "subnational2"]

    with pd.ExcelWriter('{}.xlsx'.format(dataset)) as writer:
        for df in dfs:
            if df[1] == "Country":
                df[0].to_excel(writer, sheet_name=df[1] + " tree cover loss", float_format="%.2f", freeze_panes=(1, 1),
                               index=False, columns=gadm[:1] + area_stats + tc_col, header=gadm_alias[:1] + area_stats_alias + tc_col_alias)
                df[0].to_excel(writer, sheet_name=df[1] + " carbon data", float_format="%.2f", freeze_panes=(1, 1),
                               index=False, columns=gadm[:1] + carbon_cols + carbon_emissions_yearly, header=gadm_alias[:1] + carbon_cols + carbon_emissions_yearly)
            elif df[1] == "Subnational 1":
                df[0].to_excel(writer, sheet_name=df[1] + " tree cover loss", float_format="%.2f", freeze_panes=(1, 1),
                               index=False, columns=gadm[:2] + area_stats + tc_col, header=gadm_alias[:2] + area_stats_alias + tc_col_alias)
                df[0].to_excel(writer, sheet_name=df[1] + " carbon data", float_format="%.2f", freeze_panes=(1, 1),
                               index=False, columns=gadm[:2] + carbon_cols + carbon_emissions_yearly, header=gadm_alias[:2] + carbon_cols + carbon_emissions_yearly)
            else:
                df[0].to_excel(writer, sheet_name=df[1] + " tree cover loss", float_format="%.2f", freeze_panes=(1, 1),
                               index=False, columns=gadm + area_stats + tc_col, header=gadm_alias + area_stats_alias + tc_col_alias)
                df[0].to_excel(writer, sheet_name=df[1] + " carbon data", float_format="%.2f", freeze_panes=(1, 1),
                               index=False, columns=gadm + carbon_cols + carbon_emissions_yearly, header=gadm_alias + carbon_cols + carbon_emissions_yearly)


if __name__ == "__main__":
    cli()