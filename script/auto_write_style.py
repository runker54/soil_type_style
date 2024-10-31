import pyautogui
import pyperclip
import keyboard
import time
import threading
import pandas as pd
from functools import wraps

# 安全设置
pyautogui.FAILSAFE = True  # 启用故障安全，移动鼠标到左上角可以终止程序

# 位置字典
position = {
    '符号位置': (335, 250),
    '描述位置': (1136, 986),
    '名称位置': (1169, 313),
    '类别位置': (1170, 369),
    '唯一键': (1195, 550),
    '应用按钮': (1293, 954),
    '属性按钮': (1184, 983),
    '颜色按钮': (1399, 373),
    '颜色值按钮': (1410, 741),
    'hex输入框': (1045, 624),
    '确定按钮': (1095, 661),
    '边框颜色按钮': (1403, 401),
    '边框颜色值按钮': (1395, 771),
    '边框hex输入框': (1042, 627),
    '边框确定按钮': (1096, 664),
    '倒数第二个位置': (340, 948)
}

class SafeController:
    def __init__(self):
        self.running = True
        self.setup_exit_handler()
        
    def setup_exit_handler(self):
        keyboard.on_press_key('esc', self.handle_exit)
    
    def handle_exit(self, _):
        print('\n检测到ESC按键，正在安全终止程序...')
        self.running = False
    
    def is_running(self):
        return self.running

def check_exit(func):
    @wraps(func)
    def wrapper(self, *args, **kwargs):
        if not self.controller.is_running():
            return False
        return func(self, *args, **kwargs)
    return wrapper

class AutomationController:
    def __init__(self, position_data):
        self.position = position_data
        self.controller = SafeController()
        
    @check_exit
    def move_mouse(self, x, y):
        pyautogui.moveTo(x, y, duration=0.1)
        time.sleep(0.1)
        return True

    @check_exit
    def click_mouse(self):
        pyautogui.click()
        time.sleep(0.1)
        return True

    @check_exit
    def input_text(self, text):
        pyperclip.copy(text)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.1)
        return True

    @check_exit
    def clear_text(self):
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('delete')
        time.sleep(0.1)
        return True

    @check_exit
    def press_down(self):
        pyautogui.press('down')
        time.sleep(0.1)
        return True

    @check_exit
    def press_enter(self):
        pyautogui.press('enter')
        time.sleep(0.1)
        return True

    def write_tl(self, index, t_name, t_class, t_hex):
        steps = [
            # 符号位置操作
            lambda: self.move_mouse(self.position['符号位置'][0], self.position['符号位置'][1]),
            lambda: self.click_mouse(),
            
            # 描述位置操作
            lambda: self.move_mouse(self.position['描述位置'][0], self.position['描述位置'][1]),
            lambda: self.click_mouse(),
            
            # 名称位置操作
            lambda: self.move_mouse(self.position['名称位置'][0], self.position['名称位置'][1]),
            lambda: self.click_mouse(),
            lambda: self.clear_text(),
            lambda: self.input_text(t_name),
            
            # 类别位置操作
            lambda: self.move_mouse(self.position['类别位置'][0], self.position['类别位置'][1]),
            lambda: self.click_mouse(),
            lambda: self.clear_text(),
            lambda: self.input_text(t_class),
            
            # 唯一键操作
            lambda: self.move_mouse(self.position['唯一键'][0], self.position['唯一键'][1]),
            lambda: self.click_mouse(),
            lambda: self.clear_text(),
            lambda: self.input_text(str(index)),
            
            # 应用按钮操作
            lambda: self.move_mouse(self.position['应用按钮'][0], self.position['应用按钮'][1]),
            lambda: self.click_mouse(),
            
            # 属性按钮操作
            lambda: self.move_mouse(self.position['属性按钮'][0], self.position['属性按钮'][1]),
            lambda: self.click_mouse(),
            
            # 颜色按钮操作
            lambda: self.move_mouse(self.position['颜色按钮'][0], self.position['颜色按钮'][1]),
            lambda: self.click_mouse(),
            lambda: self.move_mouse(self.position['颜色值按钮'][0], self.position['颜色值按钮'][1]),
            lambda: self.click_mouse(),
            lambda: self.move_mouse(self.position['hex输入框'][0], self.position['hex输入框'][1]),
            lambda: self.click_mouse(),
            lambda: self.clear_text(),
            lambda: self.input_text(t_hex),
            lambda: self.press_enter(),
            lambda: self.move_mouse(self.position['确定按钮'][0], self.position['确定按钮'][1]),
            lambda: self.click_mouse(),
            
            # 边框颜色操作
            lambda: self.move_mouse(self.position['边框颜色按钮'][0], self.position['边框颜色按钮'][1]),
            lambda: self.click_mouse(),
            lambda: self.move_mouse(self.position['边框颜色值按钮'][0], self.position['边框颜色值按钮'][1]),
            lambda: self.click_mouse(),
            lambda: self.move_mouse(self.position['边框hex输入框'][0], self.position['边框hex输入框'][1]),
            lambda: self.click_mouse(),
            lambda: self.clear_text(),
            lambda: self.input_text(t_hex),
            lambda: self.press_enter(),
            lambda: self.move_mouse(self.position['边框确定按钮'][0], self.position['边框确定按钮'][1]),
            lambda: self.click_mouse(),
            
            # 最终应用
            lambda: self.move_mouse(self.position['应用按钮'][0], self.position['应用按钮'][1]),
            lambda: self.click_mouse(),
            
            # 最后位置操作
            lambda: self.move_mouse(self.position['倒数第二个位置'][0], self.position['倒数第二个位置'][1]),
            lambda: self.click_mouse(),
            lambda: self.click_mouse(),
            lambda: self.press_down(),
        ]
        
        for step in steps:
            if not step():
                print("操作被用户终止")
                return False
        return True

    def process_data(self, df):
        print("程序将在5秒后开始运行，按ESC键可随时终止程序")
        time.sleep(5)

        for index, row in df.iterrows():
            if index <= 347:
                continue
            if not self.controller.is_running():
                print('程序已安全终止')
                return

            print(f"正在写入第{index}行,名称:{row['名称']},类别:{row['层级']},颜色:{row['颜色']}")
            t_name = str(row['名称'])
            t_class = str(row['层级'])
            t_hex = str(row['颜色']).replace('#', '')

            if "(" in t_name and t_class == '土类':
                has_symble_name = t_name
                no_symble_name = t_name.split('(')[0] + t_name.split(')')[1]
                print(has_symble_name, no_symble_name)
                if not self.write_tl(index, has_symble_name, t_class, t_hex):
                    break
                if not self.write_tl(str(index)+'_1', no_symble_name, t_class, t_hex):
                    break
            elif "(" in t_name and t_class == '土种':
                has_symble_name = t_name
                no_symble_name = t_name.replace('(', '').replace(')', '')
                no_symble_content_name = t_name.split('(')[0] + t_name.split(')')[1]
                print(has_symble_name, no_symble_name, no_symble_content_name)
                if not self.write_tl(index, has_symble_name, t_class, t_hex):
                    break
                if not self.write_tl(str(index)+'_1', no_symble_name, t_class, t_hex):
                    break
                if not self.write_tl(str(index)+'_2', no_symble_content_name, t_class, t_hex):
                    break
            else:
                print(t_name)
                if not self.write_tl(index, t_name, t_class, t_hex):
                    break
# 清理数据中的空格和特殊字符
def clean_string(x):
    if isinstance(x, str):
        # 去除所有空格，包括中间的空格
        return ''.join(x.split())
    return x
def main():
    # 这里需要准备您的DataFrame数据
    # 示例数据，请替换成您的实际数据
    info_data_path = r"D:\worker\工作\work\三普\数据\色标\result.xlsx"
    df = pd.read_excel(info_data_path)
    # 去除每一列每行数据中的空格和其他特殊字符
    df = df.map(clean_string)
    # 创建控制器实例
    controller = AutomationController(position)
    df_tl = df[df['层级'] != '土种']
    # 运行自动化流程
    controller.process_data(df_tl)
if __name__ == "__main__":
    main()