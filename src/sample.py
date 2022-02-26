from pathlib import Path
import pandas as pd

file_dept = Path.joinpath(Path.cwd(), "data", "dept.csv")
df_dept = pd.read_csv(file_dept)
dept_dict = to_dict(df_dept)
print(dept_dict)
