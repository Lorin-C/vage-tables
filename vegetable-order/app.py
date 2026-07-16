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
# 设置字体
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
# 获取全部表格
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
# 首页
# =========================

@app.route("/")
def index():

    return render_template(
        "index.html"
    )

# =========================
# 安全修改段落文字
# 保留段落格式
# =========================

def replace_paragraph(paragraph, text, size):

    for run in paragraph.runs:

        run.text = ""


    run = paragraph.add_run(text)

    set_font(
        run,
        size,
        True
    )

# =========================
# 生成Word
# =========================

@app.route(
    "/generate",
    methods=["POST"]
)
def generate():


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



    dishes_text = request.form.get(
        "dishes",
        ""
    )


    dishes = [

        x.strip()

        for x in dishes_text.splitlines()

        if x.strip()

    ]



    doc = Document(
        TEMPLATE
    )


    tables = get_all_tables(
        doc
    )

    # =========================
    # 修改客户和日期
    #
    # 保留：
    # 陈老四蔬菜批发
    #
    # 只修改下一行
    # =========================

    if len(tables) > 0:


        header_table = tables[0]


        for row in header_table.rows:


            for cell in row.cells:


                paragraphs = cell.paragraphs


                for p in paragraphs:


                    text = p.text.strip()



                    # 找客户日期这一行

                    if (
                        "客户：" in text
                        and
                        "2026" in text
                    ):


                        replace_paragraph(

                            p,

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
    # =========================

    dish_cells = []


    for table in tables[1:]:


        for row in table.rows:


            cells = row.cells


            if len(cells) >= 5:


                num = cells[0].text.strip()


                if num.isdigit():


                    n = int(num)


                    if 1 <= n <= 47:


                        # 第二列是菜品

                        dish_cells.append(
                            cells[1]
                        )

    # =========================
    # 写入菜品
    # 只修改文字
    # =========================

    for i, cell in enumerate(dish_cells):


        if i < len(dishes):


            p = cell.paragraphs[0]


            for run in p.runs:

                run.text = ""


            run = p.add_run(
                dishes[i]
            )


            set_font(
                run,
                18,
                True
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
