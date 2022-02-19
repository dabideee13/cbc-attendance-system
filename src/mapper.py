from pathlib import Path

import pandas as pd


def to_dict(df: pd.DataFrame) -> dict[str, str]:
    return {
        df.loc[index, "Name"]: df.loc[index, "Dept"]
        for index in range(len(df))
    }


def mapper(name: str, dept_dict: dict[str, str]) -> str:
    return dept_dict[name]


if __name__ == "__main__":

    file = Path.joinpath(Path.cwd(), "data", "dept.csv")
    df = pd.read_csv(file)

    file2 = Path.joinpath(Path.cwd(), "data", "processed", "Attendance Report - 01.16.2022.xlsx")
    df2 = pd.read_excel(file2, sheet_name="Sheet1", header=1)
