import pandas as pd
from os.path import dirname, abspath


def clear_AccessKey():
    project_path = dirname(dirname(abspath(__file__)))
    print(project_path)

    df1 = pd.read_excel(project_path+"/excel/coding.xlsx",
                        engine='openpyxl', header=0)
    for i in range(len(df1["设备ID"])):
        if df1["设备ID"][i]:
            print(df1["设备ID"][i])
            df1["鉴权码"][i] = ""
        with pd.ExcelWriter(project_path+"/excel/coding.xlsx") as writer:
            df1.to_excel(writer, sheet_name="code", index=False)


def get_AccessKey():
    project_path = dirname(dirname(abspath(__file__)))
    print(project_path)

    df1 = pd.read_excel(project_path+"/excel/coding.xlsx",
                        engine='openpyxl', header=0)

    for i in range(len(df1["设备ID"])):
        if df1["设备ID"][i]:
            print(df1["设备ID"][i])
            code = 0
            for char in df1["设备ID"][i]:
                code += ord(char)
            df1["鉴权码"][i] = code

    print(df1["鉴权码"])

    with pd.ExcelWriter(project_path+"/excel/coding.xlsx") as writer:
        df1.to_excel(writer, sheet_name="code", index=False)


if __name__ == "__main__":
    get_AccessKey()
