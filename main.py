import pandas as pd
import re


def source_title_update(data: object) -> object:
    data_list = data["Source Title"].to_list()
    result_list = []
    for elem in data_list:
        if isinstance(elem, float):
            result_list.append("not")
            continue
        elem = elem.upper()
        elem = re.sub("[^A-Za-z0-9]", "", elem)
        result_list.append(elem)
    data["KEY"] = result_list
    return data


def get_result(s: object, data: object) -> object:
    result_s_df = pd.merge(left=s, right=data, left_on="KEY", right_on="KEY")
    result_s_df.drop(['Unnamed: 0', 'Source Title_y'], axis=1, inplace=True)
    result_s_df.rename(columns={"Source Title_x": "Source Title"}, inplace=True)
    return result_s_df


def get_n(data: object) -> object:
    data_list = data["Affiliations"].to_list()
    result_list = []
    for elem in data_list:
        n1, n2, = 0, 0
        elem_split = elem.split("; ")
        for item in elem_split:
            if item.find("Omsk State Technical University") != -1:
                n1 += 1
            else:
                n2 += 1
        n = n1 / (n1 + n2)
        result_list.append(n)
    data["N"] = result_list
    return data


def main():
    scopus_df = pd.read_excel("scopus2020.xlsx")
    s1_df = pd.read_excel("s1.xlsx")
    s2_df = pd.read_excel("s2.xlsx")
    s3_df = pd.read_excel("s3.xlsx")
    s4_df = pd.read_excel("s4.xlsx")
    s_none_df = pd.read_excel("s_none.xlsx")

    scopus_df_update = scopus_df.filter(["Authors", " Title", "Source Title", "Affiliations"])
    scopus_df_and_key = source_title_update(scopus_df_update)
    s1_df_and_key = source_title_update(s1_df)
    s2_df_and_key = source_title_update(s2_df)
    s3_df_and_key = source_title_update(s3_df)
    s4_df_and_key = source_title_update(s4_df)
    s_none_df_and_key = source_title_update(s_none_df)

    result_s1_df = get_result(s1_df_and_key, scopus_df_and_key)
    result_s1_df = get_n(result_s1_df)

    result_s2_df = get_result(s2_df_and_key, scopus_df_and_key)
    result_s2_df = get_n(result_s2_df)

    result_s3_df = get_result(s3_df_and_key, scopus_df_and_key)
    result_s3_df = get_n(result_s3_df)

    result_s4_df = get_result(s4_df_and_key, scopus_df_and_key)
    result_s4_df = get_n(result_s4_df)

    result_s_none_df = get_result(s_none_df_and_key, scopus_df_and_key)
    result_s_none_df = get_n(result_s_none_df)

    result_s1_df.to_excel("s1_result.xlsx", index=False)
    result_s2_df.to_excel("s2_result.xlsx", index=False)
    result_s3_df.to_excel("s3_result.xlsx", index=False)
    result_s4_df.to_excel("s4_result.xlsx", index=False)
    result_s_none_df.to_excel("s_none_result.xlsx", index=False)
