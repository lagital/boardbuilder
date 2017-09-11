import parms
import pandas as pd


def do(pd_series, sheet_title, default):
    if sheet_title == "Example":
        return pd_series[parms.COLUMN_DESCRIPTION()]
    else:
        return default


if __name__ == "__main__":
    d = {parms.COLUMN_TITLE: ["Item"], parms.COLUMN_DESCRIPTION: ["Example Item"]}
    print(do(pd.Series(["Description Example"], index=["Description"]), "Example", "Default"))
    print(do(pd.Series(["Description"], index=["Description"]), "Not Example", "Default"))
