import fitz  # pip install PyMuPDF

# 对内存中的pdf页面添加水印
def water_mark(doc):
    for a in doc:
        a.insert_text((5, 15), "郑州上控电气技术有限公司", fontname="china-s", fontsize=11)


doc = fitz.open("1.pdf")
water_mark(doc)

doc.save("aaa.pdf")

width, height = fitz.paper_size("a4")  # 获取A4纸的尺寸
page_top = fitz.Rect(0, 0, width, height / 2)  # 上半张