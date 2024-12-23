# %%
import pygetwindow as gw
import pyautogui
import pandas as pd
import time
import pandas.io.clipboard as cb
import os
import sys
import threading
import keyboard

# %%
def chrome_equals(wn= '公安部涉赌资金交易查控平台 - Google Chrome'):
    all_windows = gw.getWindowsWithTitle('')
    chrome_tab = ''

    for window in all_windows:
        if 'Google Chrome' in window.title:
            chrome_tab = window.title
    
    if wn == chrome_tab:
        return True
    else:
        return False

# %%
def read_data(path):    
    data_path = path
    df = pd.read_excel(data_path, dtype=str)
    if '卡号' not in df.columns:
        print('没有找到"卡号"字段')
        sys.exit(0)
    else:
        return df

# %%
def click_action(x, y, sleep_time):
    pyautogui.click(x, y)
    time.sleep(sleep_time) 

# %%
def init_path():
    BASE_DIR = os.path.dirname(os.path.realpath(sys.argv[0]))
    # BASE_DIR = './'
    SOURCES_DIR = os.path.join(BASE_DIR, 'sources')
    df_path = input('请输入银行卡调单文件(含卡号列)路径：')
    # df_path = './sources/新第二批充值调单.xlsx'
    xs_path = os.path.join(SOURCES_DIR, '参数.xlsx')
    # output_path = os.path.join(os.path.dirname(df_path), 'output')
    output_html_path = os.path.join(os.path.dirname(df_path), 'html')
    output_main_path = os.path.join(os.path.dirname(df_path), 'main')
    # if not os.path.exists(output_path):  
    #     os.mkdir(output_path)  
    if not os.path.exists(output_html_path):  
        os.mkdir(output_html_path)
    if not os.path.exists(output_main_path):  
        os.mkdir(output_main_path)
    return df_path, xs_path, output_html_path, output_main_path

# %%
# 查询当前卡调单记录
def retrieve_action(xy, bn):  
    click_action(xy.loc['卡号文本框'][0], xy.loc['卡号文本框'][1], xy.loc['卡号文本框'][2])
    pyautogui.hotkey(['ctrl', 'a'])
    pyautogui.typewrite(bn)
    click_action(xy.loc['搜索按钮'][0], xy.loc['搜索按钮'][1], xy.loc['搜索按钮'][2])

# %%
def download_html(bank_name, save_path):
    path = os.path.join(save_path, f'{bank_name}.txt')
    # print(path)
    if os.path.exists(path):
        return
    else:
        pyautogui.hotkey(['ctrl', 'u'])
        time.sleep(3) # 等待加载页面
        pyautogui.hotkey(['ctrl', 'a'])
        pyautogui.hotkey(['ctrl', 'c'])
        html_txt = cb.paste()

        with open(path, 'w') as f:
            f.write(html_txt)

        pyautogui.hotkey(['ctrl', 'w'])
        time.sleep(1) # 等待关闭

# %%
def download_execl(xy, bn):
    pyautogui.scroll(int(xy.loc['下载按钮'][3])) # 向下滚动1700像素
    time.sleep(1) # 等待加载下载页面
    click_action(xy.loc['下载按钮'][0], xy.loc['下载按钮'][1], xy.loc['下载按钮'][2]) # 点击下载
    if chrome_equals('银行卡结果列表 - Google Chrome') == True:
        return
    if chrome_equals('无标题 - Google Chrome') == False:
        print(f'程序中断，下载卡号: {bn} 中断')
        time.sleep(9999)
        sys.exit(0)
    click_action(xy.loc['下载确定'][0], xy.loc['下载确定'][1], xy.loc['下载确定'][2]) # 点击下载确定

# %%
def run_rpa():
    df_path, xs_path, output_html_path ,output_main_path = init_path()
    num_f = int(input('一张卡几个流水文件：'))
    num_d = input('输入开始下载卡号（默认从头开始）：')
    print('5秒后运行')
    time.sleep(2)

    df = read_data(df_path)
    xy = pd.read_excel(xs_path, sheet_name='坐标').set_index('index')
    df_list = df['卡号'].to_list()

    if num_d != '': # 判断是否需要断点下载
        index_of_special_element = df_list.index(num_d)
        df_list = df_list[index_of_special_element:]

    for bn in df_list:
        retrieve_action(xy, bn)
        pyautogui.scroll(int(xy.loc['搜索按钮'][3])) # 滚动像素
        pyautogui.hotkey(['f11'])   # 防止页面不显示
        time.sleep(0.5)
        pyautogui.hotkey(['f11'])
        time.sleep(0.5)
        if chrome_equals('公安部涉赌资金交易查控平台 - Google Chrome') == False:
            print(f'程序中断，下载卡号: {bn} 中断')
            time.sleep(9999)
            sys.exit(0)
        
        x_ck = xy.loc['查看按钮'][0]
        y_ck = xy.loc['查看按钮'][1]
        if num_f == 0:
            click_action(x_ck, y_ck, xy.loc['查看按钮'][2])
            if chrome_equals('持卡主体详情 - Google Chrome') == True:
                download_html(bn, output_main_path)
                pyautogui.hotkey(['ctrl', 'w']) # 关闭查看页面
                time.sleep(1)
                continue
            if chrome_equals('银行卡结果列表 - Google Chrome') == False:
                print(f'程序中断，下载卡号: {bn} 中断')
                time.sleep(9999)
                sys.exit(0)
            download_html(bn, output_html_path)
            pyautogui.hotkey(['ctrl', 'w']) # 关闭查看页面
            time.sleep(1)
        else:
            for i in range(num_f):
                click_action(x_ck, y_ck, xy.loc['查看按钮'][2])
                if chrome_equals('公安部涉赌资金交易查控平台 - Google Chrome') == True:
                    if i < num_f - 1:
                        print(f"{bn} 仅查看 {i + 1} 条数据")
                    y_ck += 36
                    continue
                if chrome_equals('持卡主体详情 - Google Chrome') == True:
                    download_html(bn, output_main_path)
                    pyautogui.hotkey(['ctrl', 'w']) # 关闭查看页面
                    time.sleep(1)
                    y_ck += 36
                    continue
                if chrome_equals('银行卡结果列表 - Google Chrome') == False:
                    print(f'程序中断，下载卡号: {bn} 中断')
                    time.sleep(9999)
                    sys.exit(0)

                download_html(bn, output_html_path)
                download_execl(xy, bn)
                pyautogui.hotkey(['ctrl', 'w']) # 关闭查看页面
                time.sleep(1)
                y_ck += 40
        pyautogui.scroll(int(xy.loc['搜索按钮'][3]) * -1) # 滚动像素

# %%
def do_something():
    keyboard.wait('esc')
    sys.exit(0)

# 创建线程并启动
t = threading.Thread(target=run_rpa, daemon=True)
t.start()

keyboard.wait('esc')
print('程序结束！')
time.sleep(3)


