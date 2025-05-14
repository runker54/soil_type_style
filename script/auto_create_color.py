import pandas as pd
import numpy as np
from collections import defaultdict

def clean_text(text):
    """清理文本中的换行和多余空格"""
    return ' '.join(str(text).strip().split())

def hex_to_rgb(hex_color):
    """将Hex颜色转换为RGB元组"""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def rgb_to_hex(rgb):
    """将RGB元组转换为Hex颜色"""
    return '#{:02x}{:02x}{:02x}'.format(int(rgb[0]), int(rgb[1]), int(rgb[2]))

def interpolate_colors(color1, color2, count):
    """在两个颜色之间生成等间隔的过渡色
    
    Args:
        color1: 最深颜色
        color2: 最浅颜色
        count: 需要生成的颜色数量
    """
    if count <= 1:
        return [color1]
        
    rgb1 = np.array(hex_to_rgb(color1))
    rgb2 = np.array(hex_to_rgb(color2))
    
    colors = []
    for i in range(count):
        # 使用线性插值计算当前位置的颜色
        ratio = i / (count - 1)
        rgb = rgb1 + (rgb2 - rgb1) * ratio
        colors.append(rgb_to_hex(rgb))
    
    return colors

def create_color_dict(color_df):
    """创建颜色查找字典，按色系、色调和色标名组织颜色"""
    color_dict = defaultdict(lambda: defaultdict(dict))
    
    for _, row in color_df.iterrows():
        color_system = clean_text(row['色系'])
        color_tone = clean_text(row['色调'])
        color_name = clean_text(row['色标名'])
        hex_color = row['Hex']
        # 提取色标编号
        color_number = int(color_name[-1])
        color_dict[color_system][color_tone][color_number] = hex_color
    
    return color_dict

def create_recommendation_dict(color_set_df):
    """创建推荐颜色查找字典"""
    rec_dict = {}
    for _, row in color_set_df.iterrows():
        soil_class = clean_text(row['土类'])
        color_system = clean_text(row['色系'])
        color_tone = clean_text(row['推荐色调'])
        rec_dict[soil_class] = (color_system, color_tone)
    return rec_dict

def get_colors(color_dict, color_system, color_tone, count):
    """获取指定数量的颜色，如果数量超过色标数量则进行插值"""
    available_colors = color_dict[color_system][color_tone]
    if not available_colors:
        return []
    
    # 获取色标编号
    color_numbers = sorted(available_colors.keys(), reverse=True)
    
    if count <= len(color_numbers):
        # 如果需要的颜色数量不超过色标数量，使用现有色标
        return [available_colors[num] for num in color_numbers[:count]]
    else:
        # 如果需要更多颜色，在最深和最浅色标之间插值
        deepest_color = available_colors[color_numbers[0]]  # 最深色
        lightest_color = available_colors[color_numbers[-1]]  # 最浅色
        return interpolate_colors(deepest_color, lightest_color, count)

def assign_colors(soil_type_df, color_dict, rec_dict):
    """为土壤分类分配颜色"""
    result = []
    used_colors = set()  # 用于跟踪已使用的颜色
    
    # 清理数据
    cleaned_df = soil_type_df.apply(lambda x: x.map(clean_text))
    
    # 为每个土类分配颜色
    for soil_class in sorted(cleaned_df['土类'].unique()):
        if soil_class not in rec_dict:
            print(f"警告: 土类 '{soil_class}' 在推荐颜色表中未找到")
            continue
        
        color_system, color_tone = rec_dict[soil_class]
        available_colors = color_dict[color_system][color_tone]
        
        if not available_colors:
            print(f"警告: 色系 '{color_system}' 色调 '{color_tone}' 没有可用的色标")
            continue
        
        # 获取最深的颜色作为土类颜色
        deepest_color_number = max(available_colors.keys())
        soil_class_color = available_colors[deepest_color_number]
            
        # 添加土类
        result.append({
            '层级': '土类',
            '名称': soil_class,
            '父级': '',
            '色系': color_system,
            '色调': color_tone,
            '颜色': soil_class_color
        })
        
        # 获取该土类下的所有子项
        class_df = cleaned_df[cleaned_df['土类'] == soil_class]
        
        # 为亚类分配颜色
        subclasses = sorted(class_df['亚类'].unique())
        subclass_colors = get_colors(color_dict, color_system, color_tone, len(subclasses))
        
        for i, subclass in enumerate(subclasses):
            result.append({
                '层级': '亚类',
                '名称': subclass,
                '父级': soil_class,
                '色系': color_system,
                '色调': color_tone,
                '颜色': subclass_colors[i]
            })
        
        # 为土属分配颜色
        genera = sorted(class_df['土属'].unique())
        genus_colors = get_colors(color_dict, color_system, color_tone, len(genera))
        
        for i, genus in enumerate(genera):
            result.append({
                '层级': '土属',
                '名称': genus,
                '父级': class_df[class_df['土属'] == genus]['亚类'].iloc[0],
                '色系': color_system,
                '色调': color_tone,
                '颜色': genus_colors[i]
            })
        
        # 为土种分配颜色
        species = sorted(class_df['土种'].unique())
        species_colors = get_colors(color_dict, color_system, color_tone, len(species))
        
        for i, species_name in enumerate(species):
            color = species_colors[i]
            # 检查颜色是否重复
            while color in used_colors:
                # 如果重复，生成新的过渡色
                deepest_color = available_colors[max(available_colors.keys())]
                lightest_color = available_colors[min(available_colors.keys())]
                # 生成更多的过渡色以增加可选择的颜色数量
                new_colors = interpolate_colors(deepest_color, lightest_color, len(species) + 10)
                # 选择一个未使用的颜色
                for new_color in new_colors:
                    if new_color not in used_colors:
                        color = new_color
                        break
                else:
                    # 如果所有颜色都被使用，稍微调整颜色值
                    rgb = hex_to_rgb(color)
                    adjusted_rgb = tuple(min(255, max(0, v + np.random.randint(-5, 5))) for v in rgb)
                    color = rgb_to_hex(adjusted_rgb)
            
            used_colors.add(color)
            result.append({
                '层级': '土种',
                '名称': species_name,
                '父级': class_df[class_df['土种'] == species_name]['土属'].iloc[0],
                '色系': color_system,
                '色调': color_tone,
                '颜色': color
            })
    
    return pd.DataFrame(result)

def process_soil_colors(soil_type_df, color_df, color_set_df):
    """主处理函数"""
    color_dict = create_color_dict(color_df)
    rec_dict = create_recommendation_dict(color_set_df)
    return assign_colors(soil_type_df, color_dict, rec_dict)

color_table_path = r"D:\worker\工作\work\三普\数据\色标\更改后\土壤色标表_RGB.xlsx"
soil_type_table_path = r"D:\worker\工作\work\三普\数据\色标\更改后\分类表_填充.xlsx"
color_set_table_path = r"D:\worker\工作\work\三普\数据\色标\更改后\颜色推荐表.xlsx"
color_df = pd.read_excel(color_table_path)
soil_type_df = pd.read_excel(soil_type_table_path)
color_set_df = pd.read_excel(color_set_table_path)

result_df = process_soil_colors(soil_type_df, color_df, color_set_df)
# 保存结果
result_df.to_excel(r"D:\worker\工作\work\三普\数据\色标\更改后\result_rgb1.xlsx", index=False)