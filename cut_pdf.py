from PyPDF2 import PdfReader, PdfWriter

# ===================== 请修改参数 =====================
pdf_path = r"D:\QAXDownload\dc\bdc.pdf"
# 👇 这里把 - 全部改成 , ！！！
page_ranges = [(1, 50),(51, 100),(101, 150),(151, 200),(201, 250),(251, -1)]  # -1代表最后一页
# ======================================================

reader = PdfReader(pdf_path)
total_pages = len(reader.pages)

for i, (start, end) in enumerate(page_ranges, 1):
    # 处理最后一页
    end = total_pages if end == -1 else end
    writer = PdfWriter()
    
    # 加入指定页码
    for page_num in range(start-1, end):
        writer.add_page(reader.pages[page_num])
    
    # 保存
    with open(f"自定义拆分_{i}_第{start}-{end}页.pdf", "wb") as f:
        writer.write(f)

print("✅ 自定义拆分完成！")