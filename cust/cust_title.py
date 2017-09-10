import parms
import pandas as pd


def do(pd_series, sheet_title, default):
    if sheet_title == "Example":
        return pd_series[parms.COLUMN_TITLE()]
    else:
        return default


if __name__ == "__main__":
    d = {parms.COLUMN_TITLE(): ["Item"], parms.COLUMN_DESCRIPTION(): ["Example Item"]}
    print(do(pd.DataFrame(data=d, columns=[parms.COLUMN_TITLE, parms.COLUMN_DESCRIPTION]), "Example", "Default"))
    print(do(pd.DataFrame(data=d, columns=[parms.COLUMN_TITLE, parms.COLUMN_DESCRIPTION]), "Not example", "Default"))
