import os
import requests
import base64
from SoftConfig import *  # 从config文件内引用AK，SK。方便以后设置
import time  # 主要是生成文件夹时，带上时间信息
import xlwt  # 需要安装库 pip install xlwt
import fitz  # pip install PyMuPDF
import shutil  # 文件夹操作
import msvcrt  # 为了实现按任意键退出

# 声明全局变量
repeat_num = 0  # 标记重复发票的个数
col_num = 1  # 标记当前行号
out_file_name = r"\发票【软件自动生成】" + time.strftime("%Y%m%d_%H%M%S")  # 保存生成的文件夹名称
out_file_path = os.getcwd() + out_file_name  # 完整的文件夹名称
file_type = ""  # 用来保存正在处理文件的类型，pdf , img
pdf_full_path = ""  # 用来存当前文件的完整路径
source_folder_name = r"\待处理发票"  # 待处理发票文件夹的名字
err_file_count = 0  # 标记不能正确处理发票的个数
multiple_Pages_path_list = []  # 用来存多页pdf文件的路径
pdf_page = 0  # 标记了一下pdf的页数


# 1_2.判断原始发票文件夹是否存在，不存在就创建一个
def source_folder_crcreate():
    if not os.path.exists(os.getcwd() + source_folder_name):
        os.makedirs(os.getcwd() + source_folder_name)


# 1_1.待处理发票文件夹遍历
def file_walk():
    # 文件夹判定
    source_folder_crcreate()
    # 文件夹遍历
    res_list = []
    for path in os.walk(os.getcwd() + source_folder_name):
        for file_name in path[2]:
            res_list.append(path[0] + "\\" + file_name)
    return res_list


# 2.获取token
def get_token():
    host = 'https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=' + s_cfg.API_Key + '&client_secret=' + s_cfg.Secret_Key
    res = requests.get(host)
    return res.json()['access_token']


# 3.输出文件夹创建
def out_file_crcreate():
    if not os.path.exists(out_file_path):
        os.makedirs(out_file_path)
        print("处理结果文件夹已自动创建！")
        print("*" * 47)


# 4_2.表头样式
def head_style():
    # 表头样式
    head_style = xlwt.XFStyle()  # 初始化表头样式
    head_font = xlwt.Font()  # 为样式创建字体
    head_font.name = '微软雅黑'
    head_font.bold = True  # 加粗
    head_style.font = head_font
    head_style.alignment.horz = 2  # 字体居中
    # 表格加边线
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    head_style.borders = borders
    return head_style


# 4_1.表头设定
def head_set(sheet):
    # 表宽度设定
    sheet.col(0).width = int(6.2 * 10000.0 / 38.35)  # 序号
    sheet.col(1).width = int(17.0 * 10000.0 / 38.35)  # 发票种类
    sheet.col(2).width = int(16.0 * 10000.0 / 38.35)  # 发票代码
    sheet.col(3).width = int(11.5 * 10000.0 / 38.35)  # 发票号码
    sheet.col(4).width = int(15.5 * 10000.0 / 38.35)  # 开票日期
    sheet.col(5).width = int(31.0 * 10000.0 / 38.35)  # 销售方名称
    sheet.col(6).width = int(27.0 * 10000.0 / 38.35)  # 开票内容
    sheet.col(7).width = int(14 * 10000.0 / 38.35)  # 价税合计
    # 表头内容
    sheet.write(0, 0, '序号', head_style())
    sheet.write(0, 1, '发票种类', head_style())
    sheet.write(0, 2, '发票代码', head_style())
    sheet.write(0, 3, '发票号码', head_style())
    sheet.write(0, 4, '开票日期', head_style())
    sheet.write(0, 5, '销售方名称', head_style())
    sheet.write(0, 6, '开票内容', head_style())
    sheet.write(0, 7, '价税合计', head_style())


# 4.创建Excel表
def create_cls():
    xls = xlwt.Workbook(encoding='utf-8')  # 创建一个表格
    sheet = xls.add_sheet('发票记录', cell_overwrite_ok=True)  # 创建一个sheet表

    # 表头处理
    head_set(sheet)

    return xls, sheet


# 5_1.发票发送给百度API处理，返回json
def get_invoice_info(file_path, token):
    global pdf_full_path, file_type
    file_path_lower = file_path.lower()  # 多写这一步转换，到后面判断文件后缀时，不会出错。
    pdf_full_path = file_path  # 记录文件的完整路径
    request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/vat_invoice"
    f = open(file_path_lower, 'rb')
    pdf_img = base64.b64encode(f.read())
    f.close()  # 文件打开使用完以后，要关闭一下
    # 判定文件大小，若大于4MB，则认为文件有错误，抛出一个假异常，触发异常处理
    size_byte = os.path.getsize(file_path)
    size_mb = float(size_byte) / 1024 / 1024
    if size_mb >= 4.0:
        raise Exception()
    # 判断文件类型，选择不同的参数
    if file_path_lower.endswith(".pdf"):
        global pdf_page, temp
        params = {"pdf_file": pdf_img}
        file_type = "pdf"  # 记录文件类型
        # 读取一下，发票的页数是否是多页，是多的话，加入保存多页路径的列表中
        temp_pdf = fitz.open(file_path)
        pdf_page = temp_pdf.page_count
        temp_pdf.close()
        temp = True  # 在这里做了一下标记，到本轮循环后，如果可以执行到，就认为是正常的发票
    elif file_path_lower.endswith(".jpg") or file_path_lower.endswith(".jpeg") or file_path_lower.endswith(
            ".png") or file_path_lower.endswith(".bmp"):
        params = {"image": pdf_img}
        file_type = "img"
    else:  # 不是正确的格式，抛出一个假异常，触发异常处理
        raise Exception()

    # 连 百度云API查询
    request_url = request_url + "?access_token=" + token
    headers = {'content-type': 'application/x-www-form-urlencoded'}
    res = requests.post(request_url, data=params, headers=headers)
    return res.json()


# 5_2.对发票查询的结果进行打印显示
def print_invoice_info(json_name):
    # 重复发票的判断
    invoice_num = json_name["words_result"]["InvoiceCode"] + json_name["words_result"]["InvoiceNum"]
    if invoice_num not in temp_list:
        # 清屏一下
        display_copyright()
        # 打印处理好的结果
        print("发票种类:\t", json_name["words_result"]["InvoiceType"])
        print("发票代码:\t", json_name["words_result"]["InvoiceCode"])
        print("发票号码:\t", json_name["words_result"]["InvoiceNum"])
        print("开票日期:\t", json_name["words_result"]["InvoiceDate"])
        print("销售方名称:\t", json_name["words_result"]["SellerName"])
        print("开票内容:\t", json_name["words_result"]["CommodityName"][0][
            "word"])  # 'CommodityName': [{'row': '1', 'word': '*塑料制品*气泡膜'}],
        print("价税合计:\t", json_name["words_result"]["AmountInFiguers"])
        print("处理进度：{0} / {1} ----->>  {2:.1%}".format(col_num, total_invoice_num, float(col_num / total_invoice_num)))


# 5_3_3.插入内存中的pdf文件中
def pdf_insert(doc):
    global pdf_full_path, file_type
    if file_type == "pdf":
        pdfdoc = fitz.open(pdf_full_path)
        doc.insert_pdf(pdfdoc)
    else:
        imgdoc = fitz.open(pdf_full_path)  # 打开图片
        pdfbytes = imgdoc.convert_to_pdf()  # 使用图片创建单页的 PDF
        imgpdf = fitz.open("pdf", pdfbytes)
        doc.insert_pdf(imgpdf)  # 将当前页插入文档


# 5_3_2.内容样式
def body_style():
    # 表头样式
    body_style = xlwt.XFStyle()  # 初始化表头样式
    head_font = xlwt.Font()  # 为样式创建字体
    head_font.name = '微软雅黑'
    body_style.font = head_font
    body_style.alignment.horz = 2  # 字体居中
    # 表格加边线
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    body_style.borders = borders
    return body_style


# 5_3_1.表格行写入
def body_write(json_name, sheet, doc):
    # 重复发票的判断
    global col_num, repeat_num
    invoice_num = json_name["words_result"]["InvoiceCode"] + json_name["words_result"]["InvoiceNum"]
    if invoice_num not in temp_list:
        # 表格的写入
        sheet.write(col_num, 0, col_num, body_style())
        sheet.write(col_num, 1, json_name["words_result"]["InvoiceType"], body_style())
        sheet.write(col_num, 2, json_name["words_result"]["InvoiceCode"], body_style())
        sheet.write(col_num, 3, json_name["words_result"]["InvoiceNum"], body_style())
        sheet.write(col_num, 4, json_name["words_result"]["InvoiceDate"], body_style())
        sheet.write(col_num, 5, json_name["words_result"]["SellerName"], body_style())
        sheet.write(col_num, 6, json_name["words_result"]["CommodityName"][0]["word"], body_style())
        sheet.write(col_num, 7, float(json_name["words_result"]["AmountInFiguers"]), body_style())
        temp_list.append(invoice_num)
        col_num += 1  # 只有写入一行后，行号才加1
        # pdf的插入
        pdf_insert(doc)
    else:
        repeat_num += 1


# 5_4.如果有多页正常的发票文件，添加列表
def multiple_page_invoice_list_add():
    if pdf_page > 1 and temp:
        multiple_Pages_path_list.append(pdf_full_path)


# 5_5_2.总金额样式：无边线
def body_style2():
    # 表头样式
    body_style = xlwt.XFStyle()  # 初始化表头样式
    head_font = xlwt.Font()  # 为样式创建字体
    head_font.name = '微软雅黑'
    body_style.font = head_font
    body_style.alignment.horz = 2  # 字体居中
    return body_style


# 5_5_1.计算总金额
def total_money():
    formula = "SUM(H2:H{0})".format(col_num)
    sheet.write(col_num, 7, xlwt.Formula(formula), body_style2())


# 6_2.对内存pdf两页合一的方法
def pdf_page2_to_1(src_doc):
    temp_doc = fitz.open()
    width, height = fitz.paper_size("a4")  # 获取A4纸的尺寸
    page_top = fitz.Rect(0, 0, width, height / 2)  # 上半张
    page_bottom = fitz.Rect(0, height / 2, width, height)  # 下半张
    r_tab = [page_top, page_bottom]  # 把矩形位置加入到列表中

    # 源页面开始按页码插入内存中的新pdf中
    for spage in src_doc:
        # 创建新的输出页面，新创建的页
        if spage.number % 2 == 0:
            page = temp_doc.new_page(-1, width=width, height=height)  # 插入一个新的空白页
        # 指定插入的矩形
        page.show_pdf_page(r_tab[spage.number % 2],  # 选择输出的矩形位置
                           src_doc,  # 源文件
                           spage.number)  # 页码
        # 文字水印
        if spage.number % 2 == 0:
            page.insert_text((5, 15), s_cfg.Water_Word, fontname="china-s", fontsize=11)
        else:
            page.insert_text((5, height / 2 + 15), s_cfg.Water_Word, fontname="china-s", fontsize=11)

    return temp_doc  # 返回的是合并好的内存中pdf


# 6_1.保存到到Excel表、pdf内
def save_cls(xls, doc):
    xls.save(out_file_path + r"\发票记录汇总.xls")
    doc = pdf_page2_to_1(doc)
    doc.save(out_file_path + r"\发票记录汇总.pdf")  # 保存pdf文件
    doc.close()


# 7.文件的转移，原文件夹清除
def do_folder():
    os.rename("待处理发票", "原始票据")
    shutil.move("原始票据", out_file_path)
    os.mkdir("待处理发票")
    print("处理好的发票已全部转移到【原始票据】文件夹！！！")


# 8.结果打印
def print_result():
    print("*" * 47)
    print("发票信息已全部处理完成！！！")
    print("共处理文件总数：%d" % total_invoice_num)
    print("有问题的文件数：%d" % err_file_count)
    print("重复的发票个数：%d" % repeat_num)
    print("正常的发票个数：{0}".format(col_num - 1))
    print("汇总文件：excel、pdf已全部生成！！！")
    # 如果有多页的文档，就在这里打印一下
    if multiple_Pages_path_list:
        print("*" * 47)
        print("注意 ---> 存在有多页的发票文件：")
        for a in multiple_Pages_path_list:
            print(a)
    print("*" * 47)
    print("")
    print("按任意键，退出~~~")
    ord(msvcrt.getch())


# 软件头部版权信息
def display_copyright():
    os.system("cls")
    print("==================================================")
    print("=      【V0.4】上控技术电子发票批量处理工具      =")
    print("==================================================")


# 文件不支持或错误的处理
def err_do():
    global err_file_count
    err_file_count += 1
    # 出错的文件复制一下，存放在【有问题的文件】这个文件夹当中
    err_path, err_name = os.path.split(pdf_full_path)
    err_relative_path = err_path[len(os.getcwd() + source_folder_name):]  # 出错文件夹相对目录的一部分 \sss1\ss2
    copy_real_err_path = os.getcwd() + out_file_name + "\\有问题的文件" + err_relative_path  # E:\群晖同步盘\原创开发\37.电子发票整理__Python__20220412\发票识别\发票【软件自动生成】\有问题的文件\sss1\ss2
    copy_real_err_full_path = copy_real_err_path + '\\' + err_name
    # 判断有没有存放错误文件的文件夹，没有的话，就创建一个
    if not os.path.exists(copy_real_err_path):
        os.makedirs(copy_real_err_path)
    # 有问题文件复制
    shutil.copyfile(pdf_full_path, copy_real_err_full_path)


# 主程序
if __name__ == '__main__':
    # 清屏一下
    display_copyright()
    # 软件配置信息，实例化一下
    s_cfg = SoftConfig()

    # 1.需要处理的文件的遍历成列表
    invoice_list = file_walk()  # ['待处理发票\\pdf1.pdf', '待处理发票\\pdf2.pdf', '待处理发票\\pic1.jpg', '待处理发票\\pic3.png']
    total_invoice_num = len(invoice_list)  # 发票总的个数
    # 主体程序，if__else 结构
    if total_invoice_num > 0:

        # 2.获取百度API的token
        token = get_token()

        # 3.创建一个新的文件夹
        out_file_crcreate()

        # 4.创建一个Excel表，用于存处理的结果
        xls, sheet = create_cls()
        doc = fitz.open()  # 在内存创建一个空的pdf

        # 5.从列表取出发票，for遍历
        temp_list = []  # 标记处理好的结果
        for invoice_path in invoice_list:
            try:
                temp = False  # 判定是否是多页正常发票时，用的一个临时变量
                # 5_1.百度识别发票
                res_json = get_invoice_info(invoice_path, token)

                # 5_2.打印一下发票的信息
                print_invoice_info(res_json)

                # 5_3.将这一张发票信息写入Excel表
                body_write(res_json, sheet, doc)

                # 5_4.如果有多页正常的发票文件，添加列表
                multiple_page_invoice_list_add()

            except Exception as e:
                err_do()  # 文件错误处理

            # 5_5.在金额汇总的地方加入公式
            total_money()

        # 清屏一下
        display_copyright()

        # 6.最终的Excel保存，pdf保存
        try:
            save_cls(xls, doc)
        except Exception as e:
            pass

        # 7.文件的转移，原文件夹清除
        do_folder()

        # 8.打印最终处理结果
        print_result()
    else:
        print("")
        print("【待处理发票】文件夹 ---> 空空如也！！！")
        print("")
        print("按任意键，退出~~~")
        ord(msvcrt.getch())
