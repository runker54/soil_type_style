import os
import json
import pandas as pd

def write_json_dict(df, json_path):
    # 创建一个字典来存储土种和对应的土类、亚类、土属{土种：{"土类":土类，"亚类":亚类，"土属":土属}}
    soil_type_dict = {}
    for index, row in df.iterrows():
        soil_species = row['土种']
        soil_type_dict[soil_species] = {
            "土类": row['土类'],
            "亚类": row['亚类'], 
            "土属": row['土属']
        }
    # 将字典写入json文件
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(soil_type_dict, f, ensure_ascii=False, indent=4)

if __name__ == "__main__":
    # 读取土种和土类数据
    df = pd.read_excel(r"D:\worker\工作\work\三普\数据\色标\更改后\分类表_新_全质地.xlsx")
    # 指定输出路径
    output_path = r"D:\worker_code\Terrain_Test\data\soil_dict\soil_type_dict_new_textures.json"
    # 写入json文件
    write_json_dict(df, output_path)
    print(f"土种和土类数据已写入{output_path}")
