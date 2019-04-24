from win32com.client import Dispatch
import glob
import os
import click

@click.command()
@click.argument("info_file")
@click.argument("country_folder")
def cli(info_file, country_folder)

    all_files = glob.glob(os.path.join(country_folder, "*.xlsx"))

    for file_name in all_files:
        xl = Dispatch("Excel.Application")

        wb1 = xl.Workbooks.Open(Filename=info_file)
        wb2 = xl.Workbooks.Open(Filename=file_name)

        ws1 = wb1.Worksheets(1)
        ws1.Copy(Before=wb2.Worksheets(1))

        wb2.Close(SaveChanges=True)
        xl.Quit()