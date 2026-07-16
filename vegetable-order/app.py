from flask import Flask, render_template, request, send_file
from docx import Document
from docx.table import Table
from docx.shared import Pt
from datetime import datetime
import os


app = Flask(__name__)


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATE = os.path.join(
    BASE_DIR,
    "uploads",
    "form2026.docx"
)



# =========================
# 字体设置
# =========================

def set_font(run, size, bold=True):

    run.font.name = "宋体"

    run._element.rPr.rFonts.set(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia",
        "宋体"
    )

    run.font.size = Pt(size)

    run.bold = bold



# =========================
# 获取所有表格
# =========================

def get_all_tables(doc):

    tables = []

    for tbl in doc._body._body.iter(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl"
    ):

        tables.append(
            Table(tbl, doc._body)
        )

    return tables




# =========================
# 安全替换单元格文字
# 不破坏格式
# =========================

def replace_cell_text(cell, text, size):

    paragraph = cell.paragraphs[0]


    # 清除文字，不删除段落格式

    for run in paragraph.runs:
        run.text = ""


    run = paragraph.add_run(text)


    set_font(
        run,
        size,
        True
    )





# =========================
# 首页
# =========================

@app.route("/")
def index():

    return render_template(
        "index.html"
    )




# =========================
# 生成Word
# =========================

@app.route(
    "/generate",
    methods=["POST"]
)
def generate():


    # 客户允许为空

    customer = request.form.get(
        "customer",
        ""
    )


    date = request.form.get(
        "date",
        ""
    )


    if not date:

        date = datetime.now().strftime(
            "%Y年%m月%d日"
        )



    # 获取菜品

    dishes_text = request.form.get(
        "dishes",
        ""
    )


    dishes = [

        x.strip()

        for x in dishes_text.splitlines()

        if x.strip()

    ]



    # 打开模板

    doc = Document(
        TEMPLATE
    )


    tables = get_all_tables(
        doc
    )



    # =========================
    # 写客户和日期
    # 不改变原单元格结构
    # =========================


    if len(tables) > 0:


        header_table = tables[0]


        for row in header_table.rows:


            for cell in row.cells:


                if (
                    "客户" in cell.text
                    or
                    "2026年" in cell.text
                ):


                    replace_cell_text(

                        cell,

                        "客户："
                        +
                        customer
                        +
                        "               "
                        +
                        date,

                        14
                    )



                    break




    # =========================
    # 找菜品列
    # 只修改第二列
    # =========================


    dish_cells = []


    for table in tables[1:]:


        for row in table.rows:


            cells = row.cells


            if len(cells) >= 5:


                number = cells[0].text.strip()



                if number.isdigit():


                    n = int(number)


                    if 1 <= n <= 47:


                        # 第二列=菜品

                        dish_cells.append(
                            cells[1]
                        )




    # =========================
    # 写入菜品
    # 不碰其它列
    # =========================


    for i, cell in enumerate(dish_cells):


        if i < len(dishes):


            replace_cell_text(

                cell,

                dishes[i],

                18

            )




    # =========================
    # 保存
    # =========================


    output_dir = os.path.join(
        BASE_DIR,
        "output"
    )


    os.makedirs(
        output_dir,
        exist_ok=True
    )



    filename = (

        "蔬菜单_"

        +

        datetime.now().strftime(
            "%Y%m%d%H%M%S"
        )

        +

        ".docx"

    )



    output = os.path.join(

        output_dir,

        filename

    )



    doc.save(
        output
    )



    return send_file(

        output,

        as_attachment=True

    )





# =========================
# Railway启动
# =========================

if __name__ == "__main__":


    port = int(

        os.environ.get(
            "PORT",
            5000
        )

    )


    app.run(

        host="0.0.0.0",

        port=port

    )
