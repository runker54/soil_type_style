import pandas as pd
import numpy as np
import re

def create_complete_soil_type_table(input_file, output_file):
    # 读取原始表格
    df = pd.read_excel(input_file)
    
    # 定义所有可能的质地类型
    texture_types = ['砂质', '砂壤', '壤质', '黏壤', '黏质']
    
    # 创建新的DataFrame来存储结果
    new_rows = []
    
    # 遍历每一行
    for idx, row in df.iterrows():
        soil_species = row['土种']
        # 检查土种名中是否包含五种质地之一
        found_texture = None
        for texture in texture_types:
            if texture in soil_species:
                found_texture = texture
                break
        # 如果包含质地关键词
        if found_texture:
            # 生成其它四种质地的名称
            for texture in texture_types:
                if texture == found_texture:
                    continue
                # 替换质地关键词
                new_soil_species = soil_species.replace(found_texture, texture, 1)
                # 检查新名称是否已存在于原表
                if not ((df['土种'] == new_soil_species).any() or any(r['土种'] == new_soil_species for r in new_rows)):
                    new_row = row.copy()
                    new_row['土种'] = new_soil_species
                    new_rows.append(new_row)
    
    # 合并新旧数据
    new_df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
    
    # 保存结果
    new_df.to_excel(output_file, index=False)
    print(f"处理完成！结果已保存到 {output_file}")

if __name__ == "__main__":
    input_file = r"D:\worker\工作\work\三普\数据\色标\更改后\分类表_新_全质地.xlsx"  # 输入文件路径
    output_file = r"D:\worker\工作\work\三普\数据\色标\更改后\分类表_新_全质地_补充.xlsx"  # 输出文件路径
    create_complete_soil_type_table(input_file, output_file)
