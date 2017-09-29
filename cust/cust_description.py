import parms
import pandas as pd


def do(pd_series, sheet_title, default):
    description = str(pd_series[parms.COLUMN_DESCRIPTION()])

    if sheet_title == "Example":
        return description
    elif sheet_title == "Example2":
        description = adopt_text(pd_series["Next Location"])
    else:
        return default

    return description


def adopt_text(text):
    return str.encode(text).replace(b'\015', b'\n').decode(errors='strict')


if __name__ == "__main__":
    d = {parms.COLUMN_TITLE: ["Item"], parms.COLUMN_DESCRIPTION: ["Example Item"]}
    print(do(pd.Series(["Description Example"], index=["Description"]), "Example", "Default"))
    print(do(pd.Series(["Description"], index=["Description"]), "Not Example", "Default"))
