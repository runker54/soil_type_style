import sqlite3
import pandas as pd
import json
import binascii # 虽然在当前代码中未使用，但保留以防未来需要

def create_content(color_rgb):
    """
    创建面状符号CONTENT的二进制内容，color_rgb为[R,G,B]列表
    """
    # 完整的JSON结构 (面状符号)
    style_json = {
        "type": "CIMPolygonSymbol",
        "symbolLayers": [
            {
                "type": "CIMSolidStroke",
                "enable": True,
                "capStyle": "Round",
                "joinStyle": "Round",
                "lineStyle3D": "Strip",
                "miterLimit": 4,
                "width": 0, # 通常轮廓线宽度为0表示无轮廓，或者可以设置为正值
                "color": {
                    "type": "CIMRGBColor",
                    "values": color_rgb + [100]  # 添加透明度100
                }
            },
            {
                "type": "CIMSolidFill",
                "enable": True,
                "color": {
                    "type": "CIMRGBColor",
                    "values": color_rgb + [100]  # 添加透明度100
                }
            }
        ],
        "angleAlignment": "Map"
    }

    # 将JSON转换为字符串，然后转换为二进制
    json_str = json.dumps(style_json, separators=(',', ':'))  # 移除空格以匹配原格式
    return json_str.encode('utf-8') + b'\x00'  # 添加结尾的null字节

def create_point_content(color_rgb, size=5):
    """
    创建点状符号CONTENT的二进制内容 (添加缺失属性, 保留多边形几何, 恢复空字节)
    color_rgb: [R, G, B] 列表
    size: 点符号的大小 (直径)
    """
    style_json = {
        "type": "CIMPointSymbol",
        "symbolLayers": [{
            "type": "CIMVectorMarker",
            "enable": True,
            "anchorPointUnits": "Relative",
            "dominantSizeAxis3D": "Y", 
            "size": size,
            "billboardMode3D": "FaceNearPlane",
            "frame": {
                "xmin": -5.0,
                "ymin": -5.0,
                "xmax": 5.0,
                "ymax": 5.0
            },
            "markerGraphics": [{
                "type": "CIMMarkerGraphic",
                "geometry": {
                    "curveRings": [[[0.0, 5.0], {
                        "a": [[0.0, 5.0], [1.5308084989341916e-16, 0], 0, 1]
                    }]]
                },
                "symbol": {
                    "type": "CIMPolygonSymbol",
                    "symbolLayers": [{
                        "type": "CIMSolidFill",
                        "enable": True,
                        "color": {
                            "type": "CIMRGBColor",
                            "values": color_rgb + [100]
                        }
                        }],
                        "angleAlignment": "Map"
                    }
                }],
                "scaleSymbolsProportionally": True,
                "respectFrame": True
            }],
        "haloSize": 1,
        "scaleX": 1,
        "angleAlignment": "Display"
    }

    json_str = json.dumps(style_json, separators=(',', ':'))
    # --- 恢复添加空字节 --- 
    return json_str.encode('utf-8') + b'\x00' 
    # --- 结束恢复 --- 

def add_style_to_stylx(stylx_path, excel_path, symbol_type='polygon', default_category='默认类别'):
    """
    从Excel读取数据并添加到stylx文件
    stylx_path: .stylx 文件的路径
    excel_path: 包含颜色和名称等信息的 Excel 文件路径
    symbol_type: 'polygon' 或 'point'，决定创建哪种类型的符号
    default_category: 如果Excel中没有 '层级' 列，或者需要覆盖时使用的默认类别
    """
    conn = None # 在try块外部初始化conn，确保finally块总能访问
    try:
        # 连接到stylx数据库 (如果文件不存在，sqlite3会自动创建)
        conn = sqlite3.connect(stylx_path)
        cursor = conn.cursor()

        # 读取Excel文件
        try:
            df = pd.read_excel(excel_path)
        except FileNotFoundError:
            print(f"错误: 无法找到 Excel 文件 '{excel_path}'")
            return # 无法继续，退出函数
        except Exception as excel_read_error:
            print(f"读取 Excel 文件时出错: {excel_read_error}")
            return # 无法继续，退出函数

        # 根据符号类型设置 CLASS ID 和内容创建函数
        if symbol_type == 'polygon':
            class_id = 5
            content_func = create_content
            print(f"配置为创建 面状符号 (Class ID: {class_id})")
        elif symbol_type == 'point':
            class_id = 3 # 点符号的 CLASS ID 通常为 1
            content_func = create_point_content
            print(f"配置为创建 点状符号 (Class ID: {class_id})")
        else:
            raise ValueError(f"不支持的 symbol_type: '{symbol_type}'. 请使用 'polygon' 或 'point'.")

        # 对每行数据进行处理
        for index, row in df.iterrows():
            try:
                # 检查必需的列是否存在
                if '名称' not in row or 'RGB' not in row:
                    print(f"警告: 第 {index + 2} 行缺少 '名称' 或 'RGB' 列，已跳过。")
                    continue

                name = str(row['名称'])
                color_str = str(row['RGB'])
                
                # --- KEY 生成逻辑 (保持之前修改后的版本) ---
                excel_key = row.get('KEY')
                if excel_key is not None:
                    base_key = str(excel_key)
                else:
                    base_key = str(index)
                key_prefix = "poly_" if symbol_type == 'polygon' else "point_"
                key = key_prefix + base_key 
                # --- 结束 KEY 生成 ---

                # --- Category 分配逻辑 (从Excel读取，否则用默认值) ---
                category = str(row.get('层级', default_category))
                # --- 结束 Category 分配 ---

                point_size = int(row.get('Size', 5)) if symbol_type == 'point' else 5

                # 解析颜色值
                try:
                    rgb = [int(x.strip()) for x in color_str.split(',')]
                    if len(rgb) != 3:
                        raise ValueError("RGB值必须包含三个数字")
                except (ValueError, AttributeError) as color_error:
                    print(f"警告: 第 {index + 2} 行颜色格式错误 ('{color_str}')，已跳过。错误: {color_error}")
                    continue # 跳过这一行

                # 生成CONTENT (根据 symbol_type 调用不同函数)
                if symbol_type == 'point':
                    content = create_point_content(rgb, size=point_size)
                else: # polygon
                    content = create_content(rgb)

                # --- 内部函数：用于插入单条记录 (添加打印语句) ---
                def insert_item(item_name, item_key, item_category):
                    try:
                        # 在执行前打印将要插入的数据
                        print(f"  准备插入: CLASS={class_id}, CATEGORY='{item_category}', NAME='{item_name}', KEY='{item_key}'")
                        cursor.execute("""
                            INSERT INTO ITEMS (CLASS, CATEGORY, NAME, TAGS, CONTENT, KEY)
                            VALUES (?, ?, ?, ?, ?, ?)
                        """, (
                            class_id, item_category, item_name, "rgb;导入", content, item_key
                        ))
                    except sqlite3.IntegrityError:
                        print(f"警告: KEY '{item_key}' 已存在于数据库中，跳过插入 NAME '{item_name}'。")
                    except sqlite3.Error as insert_err:
                        print(f"错误: 插入数据时发生数据库错误 (KEY: {item_key}, NAME: {item_name}): {insert_err}")
                # --- 结束 insert_item 修改 ---

                # --- 特殊名称处理逻辑 (应用于点和面) ---
                special_handling_done = False
                # --- 修改条件：移除对 symbol_type 的检查 --- 
                if "(" in name: 
                    # 统一处理所有带括号的符号名称
                    has_symble_name = name
                    # KEY 已包含前缀 ("poly_" 或 "point_")
                    has_key = key 
                    
                    # 1. 移除括号的名称和 KEY
                    no_symble_name = name.replace('(', '').replace(')', '') 
                    no_symble_key = has_key + '_nosym' 

                    # 2. 移除括号及内容的名称和 KEY
                    no_symble_content_name = name # 默认值以防分割失败
                    parts = name.split('(', 1)
                    if len(parts) == 2:
                        base_name = parts[0]
                        rest = parts[1].split(')', 1)
                        if len(rest) == 2:
                            suffix = rest[1]
                            no_symble_content_name = base_name + suffix
                        else:
                            print(f"警告: 第 {index + 2} 行名称 '{name}' 括号不匹配 (缺少右括号)，合并内容版本将使用原名。")
                    else:
                         print(f"警告: 第 {index + 2} 行名称 '{name}' 括号不匹配 (缺少左括号)，合并内容版本将使用原名。")
                    no_symble_content_key = has_key + '_merged' 
                    
                    # --- 修改打印信息以包含符号类型 ---
                    print(f"处理带括号符号({symbol_type}): '{has_symble_name}' -> 添加原始(KEY:{has_key})、移除括号(KEY:{no_symble_key}) '{no_symble_name}'、合并内容(KEY:{no_symble_content_key}) '{no_symble_content_name}'")
                    
                    # 插入三个条目
                    insert_item(has_symble_name, has_key, category)
                    insert_item(no_symble_name, no_symble_key, category) 
                    insert_item(no_symble_content_name, no_symble_content_key, category) 
                    special_handling_done = True
                # --- 结束特殊名称处理 ---

                # --- 默认插入 ---
                # 如果没有进行特殊处理 (即 面状符号且不含括号，或 点状符号)，则执行标准插入
                if not special_handling_done:
                    insert_item(name, key, category)

            except Exception as row_error:
                print(f"警告: 处理第 {index + 2} 行数据时发生意外错误: {row_error}，已跳过。行数据: {row.to_dict()}")
                continue # 继续处理下一行

        # 提交所有累积的更改
        conn.commit()
        print(f"成功处理完 Excel 文件，{symbol_type} 样式已添加/更新到 '{stylx_path}'")

    except sqlite3.Error as db_error:
        print(f"发生数据库错误: {db_error}")
        if conn:
            try:
                conn.rollback() # 尝试回滚事务
                print("数据库更改已回滚。")
            except sqlite3.Error as rollback_error:
                print(f"回滚事务时出错: {rollback_error}")

    except ValueError as val_error: # 捕获 symbol_type 错误
        print(f"配置错误: {val_error}")

    except Exception as e:
        print(f"发生未预料的错误: {e}")
        if conn:
            try:
                conn.rollback()
                print("数据库更改已回滚。")
            except sqlite3.Error as rollback_error:
                 print(f"回滚事务时出错: {rollback_error}")

    finally:
        # 确保数据库连接在使用后总是关闭
        if conn:
            conn.close()
            print("数据库连接已关闭。")

if __name__ == "__main__":
    # *** 指定统一的目标 .stylx 文件 ***
    target_stylx_file = r"D:\ArcGISProjects\workspace\sp2024\style\new_soil_type_gz.stylx"
    # 指定不同类型符号的数据源 Excel 文件
    point_polygon_excel_source = r"D:\worker\工作\work\三普\数据\色标\更改后\result_rgb_新.xlsx" # 假设面数据在此文件

    # --- 执行创建 ---
    # 1. 向统一文件添加面状符号
    print("-" * 20)
    print(f"开始向 {target_stylx_file} 添加面状符号...")
    add_style_to_stylx(target_stylx_file, point_polygon_excel_source,
                       symbol_type='polygon', default_category='土类') # 指定面状符号类型和默认类别
    print("面状符号添加/更新流程结束。")
    print("-" * 20)

    # 2. 向同一个文件添加点状符号
    print("\n" + "-" * 20)
    print(f"开始向 {target_stylx_file} 添加点状符号...")
    # 假设点数据 Excel (point_colors.xlsx) 可能包含 'Size' 列
    add_style_to_stylx(target_stylx_file, point_polygon_excel_source,
                       symbol_type='point', default_category='标记点') # 指定点状符号类型和默认类别
    print("点状符号添加/更新流程结束。")
    print("-" * 20)

    print(f"\n所有符号已添加/更新到文件: {target_stylx_file}")
    print("所有任务完成。")
