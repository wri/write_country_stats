import pandas as pd
import os
import click
import numpy as np


TROPICAL_COUNTRIES = ["ABW", "AGO", "AIA", "ATG", "BDI", "BGD", "BEN", "BFA", "BDG", "BHS", "BLZ", "BOL", "BRA",
                          "BRB", "BRN", "BTN", "BWA", "CAF", "CIV", "CMR", "COD", "COG", "COL", "CRI", "CUB", "CYM",
                          "DJI", "DMA", "DOM", "ECU", "ERI", "ETH", "GAB", "GHA", "GIN", "GLP", "GMB", "GNB", "GNQ",
                          "GRD", "GTM", "GUF", "GUY", "HND", "HTI", "IDN", "IND", "JAM", "KEN", "KHM", "KNA", "LAO",
                          "LBR", "LCA", "LKA", "MDG", "MEX", "MLI", "MMR", "MOZ", "MRT", "MSR", "MTQ", "MWI", "MYS",
                      "NAM", "NER", "NGA", "NIC", "NPL", "PAN", "PER", "PHL", "PLW", "PNG", "PRI", "PRY", "RWA",
                          "SDN", "SEN", "SGP", "SLB", "SLE", "SLV", "SOM", "SUR", "SWZ", "TCA", "TCD", "TGO", "THA",
                          "TLS", "TTO", "TZA", "UGA", "VCT", "VEN", "VGB", "VIR", "VNM", "ZMB", "ZWE"]


@click.command()
@click.argument("iso")
@click.argument("adm1")
@click.argument("adm2")
def cli(iso, adm1, adm2):


    path = os.path.dirname(__file__)
    gadm_df = pd.read_csv(os.path.join(path, "gadm36.csv"))

    iso_df = read_from_folder(iso)
    for i in range(2001, 2019):
        iso_df.loc[~iso_df["country"].isin(TROPICAL_COUNTRIES), "biomass_loss_Mt_{}".format(i)] = np.nan
        iso_df.loc[~iso_df["country"].isin(TROPICAL_COUNTRIES), "co2_emissions_Mt_{}".format(i)] = np.nan

    iso_df = pd.merge(gadm_df, iso_df, how='inner', on="country", copy=True).drop(['subnational1', "adm1_name", 'subnational2', "adm2_name"], axis=1).drop_duplicates().sort_values(by=['iso_name', "threshold"])

    iso_df = iso_df.set_index(["country", "iso_name", "threshold", "area_ha", "gain_2000_2012_ha", "extent_2000_ha", "extent_2010_ha"] + ["tc_loss_ha_{}".format(year) for year in range(2001,2019)]).iloc[:, :].div(1000000, axis=0)
    iso_df = iso_df.reset_index()

    adm1_df = read_from_folder(adm1)
    for i in range(2001, 2019):
        adm1_df.loc[~adm1_df["country"].isin(TROPICAL_COUNTRIES), "biomass_loss_Mt_{}".format(i)] = np.nan
        adm1_df.loc[~adm1_df["country"].isin(TROPICAL_COUNTRIES), "co2_emissions_Mt_{}".format(i)] = np.nan
    adm1_df = pd.merge(gadm_df, adm1_df, how='inner', on=["country", "subnational1"], copy=True).drop(['subnational1', 'subnational2', "adm2_name"], axis=1).drop_duplicates().sort_values(by=['iso_name', "adm1_name", "threshold"])

    adm1_df = adm1_df.set_index(["country", "iso_name", "adm1_name", "threshold", "area_ha", "gain_2000_2012_ha", "extent_2000_ha", "extent_2010_ha"] + ["tc_loss_ha_{}".format(year) for year in range(2001,2019)]).iloc[:, :].div(1000000, axis=0)
    adm1_df = adm1_df.reset_index()

    adm2_df = read_from_folder(adm2)
    for i in range(2001, 2019):
        adm2_df.loc[~adm2_df["country"].isin(TROPICAL_COUNTRIES), "biomass_loss_Mt_{}".format(i)] = np.nan
        adm2_df.loc[~adm2_df["country"].isin(TROPICAL_COUNTRIES), "co2_emissions_Mt_{}".format(i)] = np.nan
    adm2_df = pd.merge(gadm_df, adm2_df, how='inner', on=["country", "subnational1", "subnational2"], copy=True).drop(['subnational1', 'subnational2'], axis=1).drop_duplicates().sort_values(by=['iso_name', "adm1_name", "adm2_name", "threshold"])

    adm2_df = adm2_df.set_index(["country", "iso_name", "adm1_name", "adm2_name", "threshold", "area_ha", "gain_2000_2012_ha", "extent_2000_ha", "extent_2010_ha"] + ["tc_loss_ha_{}".format(year) for year in range(2001,2019)]).iloc[:, :].div(1000000, axis=0)
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

    tc_col = ["gain_2000_2012_ha"] + ["tc_loss_ha_{}".format(year) for year in range(2001,2019)]
    biomass_col = ["biomass_stock_2000_Mt", "avg_biomass_per_ha_Mt"]
    biomass_alias = ["biomass_stock_2000_Mt", "avg_biomass_per_ha_Mt"]
    biomass_loss = ["biomass_loss_Mt_{}".format(year) for year in range(2001, 2019)]
    co2_col = ["co2_stock_2000_Mt"]
    co2_alias = ["co2_stock_2000_Mt"]
    co2_loss = ["co2_emissions_Mt_{}".format(year) for year in range(2001, 2019)]
    area_stats = ["threshold", "area_ha", "extent_2000_ha", "extent_2010_ha"]

    gadm = ["iso_name", "adm1_name", "adm2_name"]
    gadm_alias = ["country", "subnational1", "subnational2"]

    with pd.ExcelWriter('{}.xlsx'.format(dataset)) as writer:

        for df in dfs:

            if dataset in TROPICAL_COUNTRIES or dataset == "global":
                if df[1] == "Country":
                    df[0].to_excel(writer, sheet_name=df[1] + " tree cover loss", float_format="%.2f", freeze_panes=(1, 1),
                                   index=False, columns=gadm[:1] + area_stats + tc_col, header=gadm_alias[:1] + area_stats + tc_col)
                    df[0].to_excel(writer, sheet_name=df[1] + " biomass loss", float_format="%.2f", freeze_panes=(1, 1),
                                   index=False, columns=gadm[:1] + area_stats + biomass_col + biomass_loss, header=gadm_alias[:1] + area_stats + biomass_alias + biomass_loss)
                    df[0].to_excel(writer, sheet_name=df[1] + " co2 emissions", float_format="%.2f", freeze_panes=(1, 1),
                                   index=False, columns=gadm[:1] + area_stats + co2_col + co2_loss, header=gadm_alias[:1] + area_stats + co2_alias + co2_loss)
                elif df[1] == "Subnational 1":
                    df[0].to_excel(writer, sheet_name=df[1] + " tree cover loss", float_format="%.2f", freeze_panes=(1, 1),
                                   index=False, columns=gadm[:2] + area_stats + tc_col, header=gadm_alias[:2] + area_stats + tc_col)
                    df[0].to_excel(writer, sheet_name=df[1] + " biomass loss", float_format="%.2f", freeze_panes=(1, 1),
                                   index=False, columns=gadm[:2] + area_stats + biomass_col + biomass_loss, header=gadm_alias[:2] + area_stats + biomass_alias + biomass_loss)
                    df[0].to_excel(writer, sheet_name=df[1] + " co2 emissions", float_format="%.2f", freeze_panes=(1, 1),
                                   index=False, columns=gadm[:2] + area_stats + co2_col + co2_loss, header=gadm_alias[:2] + area_stats + co2_alias + co2_loss)
                else:
                    df[0].to_excel(writer, sheet_name=df[1] + " tree cover loss", float_format="%.2f", freeze_panes=(1, 1),
                                   index=False, columns=gadm + area_stats + tc_col, header=gadm_alias + area_stats + tc_col)
                    df[0].to_excel(writer, sheet_name=df[1] + " biomass loss", float_format="%.2f", freeze_panes=(1, 1),
                                   index=False, columns=gadm + area_stats + biomass_col + biomass_loss, header=gadm_alias + area_stats + biomass_alias + biomass_loss)
                    df[0].to_excel(writer, sheet_name=df[1] + " co2 emissions", float_format="%.2f", freeze_panes=(1, 1),
                                   index=False, columns=gadm + area_stats + co2_col + co2_loss, header=gadm_alias + area_stats + co2_alias + co2_loss)
            else:
                if df[1] == "Country":
                    df[0].to_excel(writer, sheet_name=df[1] + " tree cover loss", float_format="%.2f", freeze_panes=(1, 1),
                                   index=False, columns=gadm[:1] + area_stats + tc_col, header=gadm_alias[:1] + area_stats + tc_col)
                    df[0].to_excel(writer, sheet_name=df[1] + " biomass loss", float_format="%.2f", freeze_panes=(1, 1),
                                   index=False, columns=gadm[:1] + area_stats + biomass_col, header=gadm_alias[:1] + area_stats + biomass_alias)
                    df[0].to_excel(writer, sheet_name=df[1] + " co2 emissions", float_format="%.2f", freeze_panes=(1, 1),
                                   index=False, columns=gadm[:1] + area_stats + co2_col, header=gadm_alias[:1] + area_stats + co2_alias)
                elif df[1] == "Subnational 1":
                    df[0].to_excel(writer, sheet_name=df[1] + " tree cover loss", float_format="%.2f", freeze_panes=(1, 1),
                                   index=False, columns=gadm[:2] + area_stats + tc_col, header=gadm_alias[:2] + area_stats + tc_col)
                    df[0].to_excel(writer, sheet_name=df[1] + " biomass loss", float_format="%.2f", freeze_panes=(1, 1),
                                   index=False, columns=gadm[:2] + area_stats + biomass_col, header=gadm_alias[:2] + area_stats + biomass_alias)
                    df[0].to_excel(writer, sheet_name=df[1] + " co2 emissions", float_format="%.2f", freeze_panes=(1, 1),
                                   index=False, columns=gadm[:2] + area_stats + co2_col, header=gadm_alias[:2] + area_stats + co2_alias)
                else:
                    df[0].to_excel(writer, sheet_name=df[1] + " tree cover loss", float_format="%.2f", freeze_panes=(1, 1),
                                   index=False, columns=gadm + area_stats + tc_col, header=gadm_alias + area_stats + tc_col)
                    df[0].to_excel(writer, sheet_name=df[1] + " biomass loss", float_format="%.2f", freeze_panes=(1, 1),
                                   index=False, columns=gadm + area_stats + biomass_col, header=gadm_alias + area_stats + biomass_alias)
                    df[0].to_excel(writer, sheet_name=df[1] + " co2 emissions", float_format="%.2f", freeze_panes=(1, 1),
                                   index=False, columns=gadm + area_stats + co2_col, header=gadm_alias + area_stats + co2_alias)


if __name__ == "__main__":
    cli()