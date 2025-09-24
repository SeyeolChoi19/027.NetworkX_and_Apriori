import pandas as pd 

from mlxtend.frequent_patterns import apriori 
from mlxtend.frequent_patterns import association_rules

def create_apriori_input(df: pd.DataFrame, item_col: str, key_column: int):
    def create_dummy_columns(input_dict: dict):
        ap_df = pd.DataFrame(input_dict)
        
        for i in ap_df["single_item"].unique():
            if (i != ""):
                ap_df.loc[ap_df["single_item"] == i, i] = 1
                ap_df.loc[ap_df["single_item"] != i, i] = 0

                if (i == ap_df["single_item"].unique()[-1]):
                    ap_df = ap_df.drop(columns = ["single_item"])

        return ap_df.groupby("item_key", as_index = False).sum()
    
    def convert_to_zero_one(input_data: pd.DataFrame):
        for column in input_data.columns:
            if (input_data[column].dtype == "float64"):
                input_data.loc[input_data[column] >= 1, column] = 1
                input_data.loc[input_data[column]  < 1, column] = 0

        return input_data

    output_dict = {
        "item_key"    : [],
        "single_item" : []
    }

    for (row, number) in zip(list(df[item_col]), list(df[key_column])):
        item_array = str(row).split(",")

        for item in item_array:
            output_dict["item_key"].append(number)
            output_dict["single_item"].append(item)

    return convert_to_zero_one(create_dummy_columns(output_dict))

def return_apriori_values(input_dataset: pd.DataFrame, support_value: float = 0.009):
    frequent_itemsets = apriori(input_dataset, min_support = support_value, use_colnames = True)
    frequency_rules   = association_rules(frequent_itemsets, metric = "lift", min_threshold = 1)    

    return frequent_itemsets, frequency_rules[["antecedents", "consequents", "antecedent support", "consequent support", "support", "confidence", "lift"]]

def convert_frozenset_to_string(df: pd.DataFrame, columns_to_convert: list[str] = ["antecedents", "consequents"]):
    for i in columns_to_convert:
        df[i] = df.apply(lambda x: ",".join(x[i]), axis = 1)

    return df.sort_values(by = "confidence", ascending = False)

def get_multi_order_results(df: pd.DataFrame, support_count: float) -> list[pd.DataFrame]:
    data      = df[df["category"] != "IM,IM"]
    df_list   = []
    keys_list = []

    for data_type in data["Type"].unique():
        df2 = data[data["Type"] == data_type]
        df2["index_column"] = [f"KEY_{str(i).zfill(6)}" for i in range(len(df2))]
        ap_df1 = create_apriori_input(df2, "item", "index_column").drop(columns = "item_key")
        
        frequent_items1, frequency_rules1 = return_apriori_values(ap_df1, support_count)
        frequency_rules1 = convert_frozenset_to_string(frequency_rules1)
        df_list.append(frequency_rules1)
        keys_list.append(f"Multi Order - {data_type}")
    
    return df_list, keys_list

def get_sa_multi_order_results(df: pd.DataFrame, support_count: float) -> list[pd.DataFrame]:
    data      = df[df["category"] != "IM,IM"]
    df_list   = []
    keys_list = []

    for indicator in data["SA/Non-SA"].unique():
        for data_type in data["Type"].unique():
            df2 = data[(data["Type"] == data_type) & (data["SA/Non-SA"] == indicator)]
            df2["index_column"] = [f"KEY_{str(i).zfill(6)}" for i in range(len(df2))]
            ap_df1 = create_apriori_input(df2, "item", "index_column").drop(columns = "item_key")
            
            frequent_items1, frequency_rules1 = return_apriori_values(ap_df1, support_count)
            frequency_rules1 = convert_frozenset_to_string(frequency_rules1)
            df_list.append(frequency_rules1)
            keys_list.append(f"{indicator} {data_type}")
        
    return df_list, keys_list

def save_results_to_excel(data_list: list[pd.DataFrame], filename: str, sheet_names_list: list[str]):
    writer = pd.ExcelWriter(filename, engine = "openpyxl")

    for (name, table) in zip(sheet_names_list, data_list):
        table.to_excel(writer, sheet_name = name, index = False)
    
    writer.save()

if __name__ == "__main__":
    df1  = pd.read_excel(r"C:\Users\User\Youmake\data\2023\Data Preprocessing Results\(내부용)_2023 YouMake Multi Order SA 데이터전처리 결과_v0.1 (DIC 4.4, 구독권 제외) (230627).xlsx").fillna("")
    df2  = pd.read_excel(r"C:\Users\User\Youmake\data\2023\Data Preprocessing Results\(내부용)_2023 YouMake Multi Order 데이터전처리 결과_v0.1 (DIC 4.4, 구독권 제외) (230627).xlsx").fillna("")
    df1  = df1[df1["Site code"].isin(["AU", "IN", "SEC", "US"])]
    df2  = df2[df2["Site code"].isin(["AU", "IN", "SEC", "US"])]
    
    for figure in [0.09, 0.08, 0.07, 0.06, 0.05, 0.04, 0.03, 0.02, 0.01, 0.009, 0.008, 0.007, 0.006]:
        try:
            sa_multi   = get_sa_multi_order_results(df1, figure) 
            multi_data = get_multi_order_results(df2, figure)
            save_results_to_excel(sa_multi[0] + multi_data[0], f"C:/Users/User/Youmake/data/2023/Apriori Algorithm Results/(내부용)_2023 YouMake Multi Order 연관성분석 결과_v0.1 (DIC 4.4, Support Count {figure}, 국가 필터 적용) (230627).xlsx", sa_multi[1] + multi_data[1])    
        except:
            pass
