"""Convert the GADM 4.1 administrative-boundaries shapefile to the flat CSV
schema used by gadm36.csv.
"""

import argparse
import re
from pathlib import Path

import geopandas as gpd
import pandas as pd

DEFAULT_SHP = Path(__file__).resolve().parent / "gadm_administrative_boundaries.shp"
DEFAULT_OUT = Path(__file__).resolve().parent / "gadm41.csv"

GID_NUM_RE = re.compile(r"(\d+)_\d+$")
SUFFIX_RE = re.compile(r"_\d+$")
TRAILING_DIGITS_RE = re.compile(r"(\d+)$")

SUB1_OVERRIDES = {
    ("CHN", "Hong Kong"): 998,
    ("CHN", "Macau"): 999,
    ("GBR", "England"): 1,
    ("GBR", "Scotland"): 3,
    ("UKR", "?"): 999,
}


def parse_sub(gid: object) -> "int | None":
    if gid is None or (isinstance(gid, float) and pd.isna(gid)):
        return None
    m = GID_NUM_RE.search(str(gid))
    return int(m.group(1)) if m else None


def parse_gid2_adm1(gid2: object) -> "int | None":
    if gid2 is None or (isinstance(gid2, float) and pd.isna(gid2)):
        return None
    stripped = SUFFIX_RE.sub("", str(gid2))
    if "." not in stripped:
        return None
    parent = stripped.rsplit(".", 1)[0]
    m = TRAILING_DIGITS_RE.search(parent)
    return int(m.group(1)) if m else None


def format_field(val: object) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "NULL"
    if isinstance(val, int):
        return str(val)
    s = str(val).replace('"', '""')
    return f'"{s}"'


def convert(shp_path: Path, out_path: Path) -> None:
    gdf = gpd.read_file(shp_path)
    df = pd.DataFrame(gdf.drop(columns="geometry"))
    cols = ["GID_0", "COUNTRY", "GID_1", "NAME_1", "GID_2", "NAME_2"]
    df = df[cols]

    written = 0
    dropped = 0
    with out_path.open("w", encoding="utf-8", newline="\n") as f:
        f.write('"country","iso_name","subnational1","adm1_name","subnational2","adm2_name"\n')
        for gid0, country, gid1, name1, gid2, name2 in df.itertuples(index=False, name=None):
            raw_sub1 = parse_sub(gid1)
            sub2 = parse_sub(gid2)
            gid2_adm1 = parse_gid2_adm1(gid2)
            if raw_sub1 is not None and gid2_adm1 is not None and raw_sub1 != gid2_adm1:
                dropped += 1
                continue
            sub1 = raw_sub1 if raw_sub1 is not None else SUB1_OVERRIDES.get((gid0, name1))
            row = (
                format_field(gid0),
                format_field(country),
                format_field(sub1),
                format_field(name1),
                format_field(sub2),
                format_field(name2),
            )
            f.write(",".join(row) + "\n")
            written += 1

    print(f"wrote {out_path} ({written} rows, dropped {dropped} GID_1/GID_2 adm1 mismatches)")


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--input", type=Path, default=DEFAULT_SHP)
    p.add_argument("--output", type=Path, default=DEFAULT_OUT)
    args = p.parse_args()
    convert(args.input, args.output)


if __name__ == "__main__":
    main()
