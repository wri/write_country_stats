import xlwings as xw
import glob
import os
import click
import pdb

@click.command()
@click.argument("info_file")
@click.argument("country_folder")
def cli(info_file, country_folder):
    all_files = glob.glob(os.path.join(country_folder, "*.xlsx"))

    wb1 = xw.Book(info_file)
    ws1 = wb1.sheets(1)

    for file_name in all_files:
        click.echo(file_name)
        wb2 = xw.Book(file_name)

        ws1.api.copy_worksheet(before_=wb2.sheets(1).api)

        wb2.save()
        wb2.close()

    wb1.close()
    wb1.app.quit()


if __name__ == "__main__":
    cli()