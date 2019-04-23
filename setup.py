from setuptools import setup, find_packages

setup(
    name="country_stats",
    version="0.1",
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        "Click", "pandas", "openpyxl"

    ],
    entry_points="""
        [console_scripts]
        country_stats=country_stats.country_stats:cli
    """,
)