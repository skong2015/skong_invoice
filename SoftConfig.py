# ini文件识别需要用的库
import os
from configparser import ConfigParser
import chardet  # 需要安装库：pip install chardet

str = '''[软件配置.ini]

;***********************************************

API_Key = yrhFG4cV08IZl7FA0drSZUzI

Secret_Key = MZlngitpgepK0qWAbdD5uHDpL8jD2Z3G


;API_Key、Secret_Key这都是从百度云官方获取的。

;API_Key、Secret_Key获取方法：https://www.skong2015.com

;***********************************************

[水印功能]

Water_Word = 郑州上控电气技术有限公司


;水印文字自行修改，不想水印，什么都别填写就可以了。

;***********************************************'''


class SoftConfig(object):

    # 编码方式读取
    def get_encoding(self):
        with open("软件配置.ini", 'rb') as f:
            return chardet.detect(f.read())['encoding']

    # 初始化配置
    def __init__(self):
        # 软件配置.ini 文件自动生成。
        if not os.path.exists("软件配置.ini"):
            with open("软件配置.ini", "w", encoding="utf-8") as f:
                f.write(str)

        # 获取一下，AK，SK。
        parser = ConfigParser()
        parser.read("软件配置.ini", encoding=self.get_encoding())
        self.API_Key = parser.get("软件配置.ini", 'API_Key').strip()
        self.Secret_Key = parser.get("软件配置.ini", 'Secret_Key').strip()
        self.Water_Word = parser.get("水印功能", 'Water_Word').strip()


# 测试
if __name__ == '__main__':
    s_cfg = SoftConfig()
    print(s_cfg.API_Key)
    print(s_cfg.Secret_Key)
    print(s_cfg.Water_Word)
