from docx import Document

# Rows记录每行内容,用于判断是否为列表格
Rows = []
# key记录
key = []
doc_path = '个人简历表格.docx'  # 替换为你的Word文档路径
new_doc_path = '个人简历表格1.docx'


def analyze_document(doc_path):
    doc = Document(doc_path)
    tables = doc.tables

    print("文档中的表格数量:", len(tables))

    for i, table in enumerate(tables):
        print(f"\n表格 {i + 1}:")
        print("表格行数:", len(table.rows))
        print("表格列数:", len(table.rows[0].cells))
        # print(table.rows[0].cells[0].text)

        # 打印每一行的标题
        headers = [cell.text.strip() for cell in table.rows[4].cells]
        print("表格标题:", headers)

        # 行循环检查每个单元格的内容，找出空白单元格并标记其标题
        for idy, row in enumerate(table.rows):
            row_data = [cell.text.strip() for cell in row.cells]
            Rows.append(row_data)
            # print(Rows)
            # 判断是否为列表格
            if (row.cells[0].text.strip() == Rows[idy - 1][0] or not row.cells[0].text.strip()) and idy > 0:
                if row.cells[0].text.strip() == Rows[idy - 1][0] and row.cells[0].text.strip():
                    key.append(Rows[idy - 1][0])
                    # 统计出现次数
                    key_num = key.count(Rows[idy - 1][0])
                    # 标题行
                    key_col = idy
                    key_col -= key_num
                    # 排除一个空占多行的情况
                    cell_passed = 0
                    for idx, cell in enumerate(Rows[key_col][1:]):
                        if cell != cell_passed:
                            print(f"{Rows[key_col][0]}的{cell}{key_num}需要填在第 {idy + 1} 行，第 {idx + 2} 列的单元格")
                            table.cell(idy, idx+1).text = "{}的{}{}".format(Rows[key_col][0], cell, key_num)
                            cell_passed =cell
                if not row.cells[0].text.strip():
                    for idx, cell in enumerate(row.cells):
                        print(f"{Rows[0][idx]}{idy}需要填在第 {idy + 1} 行，第 {idx + 1} 列的单元格")
                        table.cell(idy, idx).text = "{}{}".format(Rows[0][idx],idy)

            else:
                for idx, cell in enumerate(row.cells):
                    if not cell.text.strip() and idx < len(headers) and row_data[idx - 1]:
                        print(f"{row_data[idx - 1]} 需要填在第 {idy + 1} 行，第 {idx + 1} 列的单元格")
                        table.cell(idy, idx).text = row_data[idx - 1]

        doc.save(new_doc_path)

        # 清空Rows表格
        Rows.clear()
        # # 列循环检查每个单元格的内容，找出空白单元格并标记其标题检查每个单元格的内容，找出空白单元格并标记其标题
        # for idx, col in enumerate(table.columns):


if __name__ == "__main__":
    # # Rows记录每行内容,用于判断是否为列表格
    # Rows = []
    # # key记录
    # key = []
    # doc_path = '个人简历表格.docx'# 替换为你的Word文档路径
    # new_doc_path = '个人简历表格1.docx'
    analyze_document(doc_path)
