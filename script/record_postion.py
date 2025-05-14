# 位置获取
import pyautogui
import keyboard
import time
import json

def record_positions():
    positions = {}
    position_names = ['符号位置','描述位置','名称位置','类别位置','唯一键','应用按钮','属性按钮','颜色按钮','颜色值按钮','hex输入框','确定按钮',
                      '边框颜色按钮','边框颜色值按钮','边框hex输入框','边框确定按钮','倒数第二个位置']

    print("位置记录程序启动！")
    print("请按照提示将鼠标移动到相应位置，然后按空格键记录位置")
    print("按 'q' 键退出程序\n")
    time.sleep(2)
    
    for pos_name in position_names:
        print(f"请将鼠标移动到 {pos_name} 的位置，然后按空格键记录")
        
        while True:
            if keyboard.is_pressed('space'):
                x, y = pyautogui.position()
                positions[pos_name] = (x, y)
                print(f"{pos_name} 位置已记录: x={x}, y={y}")
                time.sleep(0.5)  # 防止重复记录
                break
            elif keyboard.is_pressed('q'):
                print("\n程序已退出")
                return None
    
    # 将记录的位置转换为代码
    print("\n以下是记录的位置代码：")
    for name, (x, y) in positions.items():
        print(f"# 移动到{name}")
        print(f"move_mouse({x}, {y})")
    
    return positions

# 运行函数
if __name__ == "__main__":
    print("3秒后开始记录位置...")
    time.sleep(3)
    positions = record_positions()
    # 保存为json文件
    with open('positions.json', 'w') as f:
        json.dump(positions, f)
