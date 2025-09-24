import pandas            as pd 
import numpy             as np
import networkx          as nx 
import matplotlib.pyplot as plt 

def create_network_graphs(data_dict: dict, english_map: pd.DataFrame, figure_name: float, node_size: int = 40000, layout_type: str = "Spring"):
    def korean_to_english(input_string: str):
        return_value = input_string
        output_array = []

        for (korean, english) in zip(list(english_map["korean"]), list(english_map["english"])):
            for substring in str(input_string).split(","):
                if (substring == korean):
                    return_value = english 
                    output_array.append(return_value)
                    break 

        return ",".join(output_array)

    def create_graph(save_file_name: str, data: pd.DataFrame):
        layout_dict = {
            "Spring"   : nx.spring_layout,
            "Circular" : nx.circular_layout,
            "Shell"    : nx.shell_layout,
            "Random"   : nx.random_layout
        }

        graph    = nx.from_pandas_edgelist(data, "antecedents", "consequents", ["lift"])
        position = layout_dict[layout_type](graph)
        pagerank = nx.pagerank(graph)
        plt.figure(figsize = (12,10))

        nx.draw_networkx_edges(
            graph, 
            position,
            arrows     = True, 
            arrowsize  = 30,
            edge_color = "silver",
            width      = 2.0
        )
        
        nx.draw_networkx_nodes(
            graph, 
            position, 
            node_color = list(pagerank.values()),
            node_size  = [node_size * v for v in pagerank.values()],
            cmap       = plt.cm.Blues,
            edgecolors = "gray",
            alpha      = 0.75
        )

        nx.draw_networkx_labels(
            graph,
            position, 
            font_size = 7
        )

        plt.savefig(f"C:/Users/User/Youmake/data/2023/Network X Figures/(내부용) {layout_type} {save_file_name} NetworkX 차트 (DIC 4.4, Support Count {figure_name}) (230627) v0.2.png", dpi = 600)
        plt.close()

        return pagerank

    result_list = []

    for (filename, dataframe) in data_dict.items():
        dataframe["antecedents"] = dataframe.apply(lambda x: korean_to_english(x["antecedents"]), axis = 1)
        dataframe["consequents"] = dataframe.apply(lambda x: korean_to_english(x["consequents"]), axis = 1)

        rank_data = create_graph(filename, dataframe)
        rank_df   = pd.DataFrame({
            "Node" : [i for i in rank_data.keys()],
            "Rank" : [i for i in rank_data.values()]
        })

        result_list.append(rank_df)

    return result_list

if __name__ == "__main__":
    for figure in [0.01, 0.02, 0.03, 0.009, 0.008, 0.007, 0.006]:
        try:
            data_dict = {
                "Multi Order - YouMake" : pd.read_excel(f"C:/Users/User/Youmake/data/2023/Apriori Algorithm Results/(내부용)_2023 YouMake Multi Order 연관성분석 결과_v0.1 (DIC 4.4, Support Count {figure}) (230627).xlsx", sheet_name = "Multi Order - YouMake"),
                "Multi Order - S.com"   : pd.read_excel(f"C:/Users/User/Youmake/data/2023/Apriori Algorithm Results/(내부용)_2023 YouMake Multi Order 연관성분석 결과_v0.1 (DIC 4.4, Support Count {figure}) (230627).xlsx", sheet_name = "Multi Order - S.com"),
                "SA YouMake"            : pd.read_excel(f"C:/Users/User/Youmake/data/2023/Apriori Algorithm Results/(내부용)_2023 YouMake Multi Order 연관성분석 결과_v0.1 (DIC 4.4, Support Count {figure}) (230627).xlsx", sheet_name = "SA YouMake"),
                "Non-SA YouMake"        : pd.read_excel(f"C:/Users/User/Youmake/data/2023/Apriori Algorithm Results/(내부용)_2023 YouMake Multi Order 연관성분석 결과_v0.1 (DIC 4.4, Support Count {figure}) (230627).xlsx", sheet_name = "Non-SA YouMake"),
                "SA S.com"              : pd.read_excel(f"C:/Users/User/Youmake/data/2023/Apriori Algorithm Results/(내부용)_2023 YouMake Multi Order 연관성분석 결과_v0.1 (DIC 4.4, Support Count {figure}) (230627).xlsx", sheet_name = "SA S.com"),
                "Non-SA S.com"          : pd.read_excel(f"C:/Users/User/Youmake/data/2023/Apriori Algorithm Results/(내부용)_2023 YouMake Multi Order 연관성분석 결과_v0.1 (DIC 4.4, Support Count {figure}) (230627).xlsx", sheet_name = "Non-SA S.com")
            }   

            result_list = create_network_graphs(data_dict, pd.read_excel(r"C:/Users/User/Youmake/config/(내부용)_제품명 한 to 영 매핑 파일.xlsx"), figure, 9000, "Spring")
            writer      = pd.ExcelWriter(f"C:/Users/User/Youmake/data/2023/Network X Datasets/(내부용) Rank Score 데이터 (DIC 4.4, Support Count {figure}) (230627).v0.2.xlsx", engine = "openpyxl")

            for (file_name, dataframe) in zip(data_dict.keys(), result_list):
                dataframe.to_excel(writer, sheet_name = file_name, index = False)

            writer.save()
        except Exception as e:
            print(e)
            pass
