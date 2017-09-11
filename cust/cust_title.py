import parms
import pandas as pd


def do(pd_series, sheet_title, default):
    if sheet_title == "Example":
        return pd_series[parms.COLUMN_TITLE()]
    else:
        return default


if __name__ == "__main__":
    print(do(pd.Series(["Title Example"], index=[parms.COLUMN_TITLE()]), "Example", "Default"))
    print(do(pd.Series(["Title"], index=[parms.COLUMN_TITLE()]), "Not Example", "Default"))
