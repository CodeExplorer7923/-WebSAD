# coding=utf-8
import json
import os
import webbrowser
import time
from win32com.client import Dispatch
from ctypes import wintypes
import ctypes
import sys

info_1 = '-'*10
info_2 = '-'*20
info_3 = '-'*30
info_4 = '-'*40
t = time.strftime('%Y-%m-%d  %A  %H:%M:%S')
info = f'''{t}
==============================欢迎使用本程序==============================
使用说明：【以下说明按添加时间先后排序】
1.本程序仅供学习使用，不得用于商业用途;
2.本程序最终解释权归作者所有;
3.本程序仍在开发中，若您在使用过程中遇到了问题，请积极向作者反馈;
4.详细说明文档见附件readme.txt
-----------------------------------------------------------------------
作者说：
本人常用的开发环境：python3.8
本人制作的所有程序都可以在交流群免费领取(程序都在群文件里，如需源码请在群内联系群主,发送程序名称并在末尾加上“源码”二字)
如果有人向您出售程序或者涉及到金钱交易请格外注意,那不是作者本人的意愿
免费接受程序定制,有意者群内联系群主,发送“定制程序”和要求即可,我会回复能否制作(没回复不要急,我可能在忙)
祝您生活愉快，万事顺心ヾ(•ω•`)o
-----------------------------------------------------------------------
作者:安徽工程大学-机制231-张振伟
QQ :2578713815
交流群:897871645(QQ)
=======================================================================
版本号2.1.0.0'''

class ActualFunction(object):

    def __init__(self):
        self.saved_url_path = './configuration/saved_url.json'
        self.installed_path = './configuration/installed_record.json'
        self.configuration_file_path = './configuration'

    def get_url_dict(self, path=''):
        with open(path, 'r', encoding='utf-8')as f:
            url_dict = json.load(f)
            return url_dict

    def save_json_data(self, json_data, path=''):
        with open(path, 'w', encoding='utf-8')as f:
            json.dump(json_data, f, ensure_ascii=False)
            os.fsync(f.fileno())

    def save_url(self):
        a = True
        while a:
            url = input('>>>请输入网址:')
            url_name = input('>>>请输入网址名称:')
            url_dict = self.get_url_dict(path=self.saved_url_path)
            url_keys_list = list(url_dict.keys())
            new_key = str(int(url_keys_list[-1]) + 1)
            var = [url_name, url]
            url_dict[new_key] = var
            self.save_json_data(url_dict, path=self.saved_url_path)
            exit_select = input('>>>保存成功.输q返回功能页.')
            if exit_select == 'q' or exit_select == 'Q':
                a = False
            else:
                a = True
            os.system('cls' if os.name == 'nt' else 'clear')

    def open_url(self):
        a = True
        while a:
            url_dict = self.get_url_dict(path=self.saved_url_path)
            print(info_1+'网址列表'+info_1)
            for i in url_dict:
                print(i+"."+url_dict[i][0]+':'+url_dict[i][1])
            try:
                select_url = input('>>>请输入网站序号(输q返回功能页):')
                if select_url == 'q' or select_url == 'Q':
                    a = False
                else:
                    url = url_dict[select_url][1]
                    webbrowser.open(url)
            except Exception:
                print('>>>输入无效!')
                time.sleep(0.5)
            os.system('cls' if os.name == 'nt' else 'clear')

    def delete_url(self):
        a = True
        while a:
            url_dict = self.get_url_dict(path=self.saved_url_path)
            print(info_1+'网址列表'+info_1)
            for i in url_dict:
                print(i+"."+url_dict[i][0]+':'+url_dict[i][1])
            del_select = input('>>>请选择需要删除的网址序号:')
            try:
                url_dict.pop(del_select)
                keys_list = list(url_dict.keys())

                new_keys_list = []
                for j in keys_list:
                    if int(j) < int(del_select):
                        new_keys_list.append(j)
                    elif int(j) > int(del_select):
                        j = str(int(j)-1)
                        new_keys_list.append(j)

                new_variable_list = list(url_dict.values())
                new_url_dict = {}
                for key, var in zip(new_keys_list, new_variable_list):
                    new_url_dict[key] = var
                with open(self.saved_url_path, 'w', encoding='utf-8')as f:
                    json.dump(new_url_dict, f, ensure_ascii=False)
                    os.fsync(f.fileno())
                exit_select = input('>>>删除成功.返回功能页?y/n(默认返回)')
                if exit_select == 'n' or exit_select == 'N':
                    a = True
                else:
                    a = False
            except Exception:
                print('>>>输入无效!')
                time.sleep(0.5)
                a = False
            os.system('cls' if os.name == 'nt' else 'clear')

    def clear_and_uninstall(self):
        verify = input('>>>确认卸载并清空保存的网址？y/n(ENTER默认取消)')
        if verify == 'y' or verify == 'Y':
            if os.path.exists(self.saved_url_path):
                os.remove(self.saved_url_path)
            if os.path.exists(self.installed_path):
                os.remove(self.installed_path)
            if os.path.exists(self.configuration_file_path):
                os.rmdir(self.configuration_file_path)
            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)
            input('>>>Uninstall complete.<ENTER TO EXIT>')
            sys.exit(0)
        else:
            input('>>>已取消操作.(ENTER)')

    def run(self):
        print(info_1+'功能列表'+info_1)
        func_select = input('''1.保存网址
2.打开网址
3.删除网址
4.卸载并清空程序数据
>>>请选择功能：''')
        try:
            if func_select == '1':
                os.system('cls' if os.name == 'nt' else 'clear')
                self.save_url()

            elif func_select == '2':
                os.system('cls' if os.name == 'nt' else 'clear')
                self.open_url()

            elif func_select == '3':
                os.system('cls' if os.name == 'nt' else 'clear')
                self.delete_url()

            elif func_select == '4':
                os.system('cls' if os.name == 'nt' else 'clear')
                self.clear_and_uninstall()

            else:
                print('>>>输入无效!')
                time.sleep(0.5)
        except Exception:
            print('>>>操作不合法!')
            time.sleep(0.5)


class IndexInstall(object):

    def __init__(self):
        self.info = '>>>Loading'
        self.configuration_file_path = './configuration'
        self.installed_path = './configuration/installed_record.json'
        self.saved_url_path = './configuration/saved_url.json'
        self.desktop_path = self.get_desktop_path()


    def install_and_configurate(self):
        installed_record = {'installed': True}
        shortcut_path = self.desktop_path+r'\Web_SAD.lnk'
        for i in range(5):
            print(self.info + '.' * i)
            time.sleep(0.2)
            os.system('cls' if os.name == 'nt' else 'clear')
        if not os.path.exists(self.configuration_file_path):
            os.mkdir(self.configuration_file_path)
        if not os.path.exists(path=self.installed_path):
            self.create_shortcut(path=shortcut_path, target=os.getcwd()+R'\Web_SAD2.1.0.0.exe', wDir=os.getcwd())  # *********************************
            with open(self.installed_path, 'w', encoding='utf-8')as f:
                json.dump(installed_record, f)
        if not os.path.exists(path=self.saved_url_path):
            with open(self.saved_url_path, 'w', encoding='utf-8')as f:
                json.dump({
                    "1": ["百度", "https://www.baidu.com/"],
                    "2": ["安徽工程大学官网", 'https://www.ahpu.edu.cn/'],
                    "3": ["超星学习通", 'i.chaoxing.com'],
                    "4": ["智慧团建", 'https://zhtj.youth.cn/zhtj/'],
                    '5': ['四六级报考与成绩查询', 'https://cet.neea.edu.cn/'],
                    '6': ['外网接入安徽工程大学校园网指南','https://metc.ahpu.edu.cn/_t501/2016/0516/c7234a39783/page.psp'],
                    '7': ['哔哩哔哩', 'https://www.bilibili.com/'],
                }, f, ensure_ascii=False)

    def create_shortcut(self, path, target, wDir='', icon=''):
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(path)
        shortcut.TargetPath = target
        shortcut.WorkingDirectory = wDir
        if icon == '':
            pass
        else:
            shortcut.IconLocation = icon
        shortcut.save()

    def get_desktop_path(self):  # 获取桌面路径
        CSIDL_DESKTOP = 0x0000
        SHGFP_TYPE_CURRENT = 0
        buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_DESKTOP, None, SHGFP_TYPE_CURRENT, buf)
        desktop_path = buf.value
        return desktop_path

    def run(self):
        self.install_and_configurate()
        input('>>>初始化完成.请重启程序.')


class OrganizeFunction(object):
    def init_check(self):
        try:
            with open('./configuration/installed_record.json', 'r', encoding='utf-8') as f:
                variable_1 = json.load(f)['installed']
                if variable_1:
                    return variable_1
                else:
                    install.run()
        except Exception:
            install.run()


if __name__ == '__main__':
    install = IndexInstall()
    actual_func = ActualFunction()
    organize_func = OrganizeFunction()
    start_ = organize_func.init_check()
    shortcut_path = install.get_desktop_path() + r'\Web_SAD.lnk'
    print(info)
    input('>>>ENTER开始程序')
    if start_:
        while 1:
            os.system('cls' if os.name == 'nt' else 'clear')
            actual_func.run()
            os.system('cls' if os.name == 'nt' else 'clear')