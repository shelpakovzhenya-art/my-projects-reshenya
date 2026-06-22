from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
DOCX_DIR = ROOT / "reports-docx"
TODAY = "22.06.2026"

TASK_SOURCE = (
    "Задача: проверить проекты из таблицы доступов, найти SEO-проблемы, "
    "исправить безопасные пункты в CMS и подготовить ТЗ на оставшиеся правки."
)

DATA_SOURCE = (
    "Данные: таблица доступов Google Sheets, таблица SEO-материалов/курсов, "
    "публичная проверка сайтов, доступные CMS-админки, результаты проверки через браузер и curl. "
    "Яндекс.Вебмастер напрямую не проверен: был нужен вход в Яндекс ID."
)

COMMON_ACCEPTANCE = [
    "После правок проверить ответы URL через curl или аналогичный инструмент.",
    "Проверить наличие canonical в HTML целевых страниц.",
    "Проверить sitemap.xml и robots.txt после изменений.",
    "После выдачи доступа зайти в Яндекс.Вебмастер и отправить важные страницы на переобход.",
]


REPORTS = [
    {
        "project": "Фильтр-пресс",
        "url": "https://filter-press.ru/",
        "done": [
            "Выполнен вход в Bitrix.",
            "Проверены разделы robots.txt, sitemap.xml, Aspro sitemap, Умный SEO.",
            "Публичный sitemap актуальный по Last-Modified: 16.04.2026.",
            "Правки не вносились: проблема в серверных редиректах и шаблоне canonical.",
        ],
        "tz": [
            ("P1", "Склеить зеркала", "Настроить 301: http -> https, www -> non-www. Основной адрес: https://filter-press.ru/."),
            ("P1", "Закрыть /index.php", "Сделать 301 с /index.php и внутренних index.php на URL без index.php."),
            ("P1", "Добавить canonical", "Вывести rel=canonical на главной, разделах, карточках и посадочных страницах."),
            ("P2", "Обработать URL-параметры", "UTM и служебные параметры должны вести на чистый URL или иметь canonical на чистую страницу."),
        ],
        "facts": [
            "http, https, www и non-www отдают 200 OK.",
            "https://filter-press.ru/index.php отдает 200 OK.",
            "CMS: 1C-Bitrix, шаблон Aspro Allcorp3.",
        ],
    },
    {
        "project": "Грани",
        "url": "https://gs-best.ru/",
        "done": [
            "Выполнен вход в ICMS.",
            "Проверена страница /o-kompanii в структуре сайта.",
            "Проверены SEO-поля /o-kompanii: title и description заполнены и отличаются от главной.",
            "Раздел редиректов ICMS проверен. Проблема http/www/:443 находится на серверном уровне.",
        ],
        "tz": [
            ("P1", "Исправить редиректные цепочки", "Настроить переход сразу на https://gs-best.ru/... без www, без HTTP-шага и без :443 в Location."),
            ("P1", "Добавить canonical", "В шаблоне сайта вывести rel=canonical для чистого URL страницы."),
            ("P2", "Обработать URL-параметры", "Для UTM, sort и других параметров настроить canonical на чистый URL или 301-редирект."),
            ("P2", "Проверить sitemap", "В sitemap должны быть только финальные URL без редиректов."),
        ],
        "facts": [
            "Редиректы проходят через HTTP и https://gs-best.ru:443/.",
            "Страница /o-kompanii имеет отдельные SEO title и description.",
            "CMS: ICMS.",
        ],
    },
    {
        "project": "Аквабурмастер",
        "url": "https://aquaburmaster.ru/",
        "done": [
            "Выполнен вход в ICMS v3.5.46.",
            "Открыт SEO-раздел.",
            "Запущена штатная генерация sitemap.xml. CMS показала успешное сохранение.",
            "Публичная проверка показала, что sitemap.xml не обновился. Правка не засчитана как завершенная.",
            "Проверены SEO-редиректы: правил /uslugi/ -> / и /kontakty/ -> / в списке нет.",
        ],
        "tz": [
            ("P1", "Исправить редиректы со слэшем", "Сделать /uslugi/ -> /uslugi и /kontakty/ -> /kontakty либо выбрать единый формат URL без редиректа на главную."),
            ("P1", "Добавить canonical", "Вывести canonical на главной, разделах, статьях и услугах."),
            ("P2", "Починить генерацию sitemap", "Проверить путь и права записи: публичный sitemap.xml не меняется после генерации в CMS."),
            ("P2", "Обновить PHP", "Сайт раскрывает X-Powered-By: PHP/5.6.40. Версия устаревшая."),
        ],
        "facts": [
            "https://aquaburmaster.ru/uslugi/ редиректит на главную.",
            "https://aquaburmaster.ru/kontakty/ редиректит на главную.",
            "sitemap.xml остался файлом 2018 года с генератором mysitemapgenerator.com.",
        ],
    },
    {
        "project": "Московская застава",
        "url": "https://mzastava.ru/",
        "done": [
            "Проверена доступность /bitrix/admin/.",
            "Вход по данным из таблицы не сработал: форма вернула ошибку неверного логина или пароля.",
            "Повторные попытки входа не выполнялись, чтобы не вызвать блокировку.",
            "Правки не вносились из-за отсутствия рабочего доступа.",
        ],
        "tz": [
            ("P1", "Исправить HTTP в редиректах", "https://mzastava.ru/offers должен сразу вести на https://mzastava.ru/offers/ без промежуточного http."),
            ("P1", "Закрыть /index.php", "Настроить 301 с /index.php на /."),
            ("P1", "Добавить canonical", "Вывести canonical на главной и типовых страницах."),
            ("P2", "Перегенерировать sitemap", "Текущий Last-Modified: 03.08.2022. Нужна пересборка из Bitrix."),
        ],
        "facts": [
            "CMS: 1C-Bitrix.",
            "sitemap.xml: Last-Modified 03.08.2022.",
            "https://mzastava.ru/index.php отдает 200 OK.",
        ],
    },
    {
        "project": "Красная мельница",
        "url": "https://www.kmel.ru/",
        "done": [
            "Выполнен вход в ICMS v2.2.",
            "В редакторе главной страницы заменен переспамленный meta description.",
            "Публичная проверка подтвердила новый meta description на главной.",
        ],
        "tz": [
            ("P1", "Параметры не должны давать 404", "UTM и служебные параметры должны открывать страницу с canonical или редиректиться на чистый URL."),
            ("P1", "Добавить canonical", "Вывести canonical на главной, /about_us, /catalog, /catalog/1-4, /contacts."),
            ("P2", "302 заменить на 301", "Слэш/без слэша должен редиректиться постоянным 301."),
            ("P2", "Убрать HTTP-цепочку", "/index.php должен сразу вести на https://www.kmel.ru/."),
        ],
        "facts": [
            "Новый description: Красная мельница — женская одежда из льна от производителя в Костроме. Каталог льняной одежды оптом и в розницу, натуральные ткани, доставка и условия сотрудничества.",
            "https://www.kmel.ru/?utm_source=test отдает 404.",
            "https://www.kmel.ru/catalog/ редиректит 302 на /catalog.",
        ],
    },
    {
        "project": "Витали",
        "url": "https://vitalikostroma.ru/",
        "done": [
            "Выполнен вход в ICMS v2.2.",
            "Открыт каталог и редактор раздела 39.",
            "В редакторе раздела найдены поля названия, состава, описания, изображений и статуса.",
            "SEO title, meta description и canonical-полей в редакторе не найдено. Контентные правки не выполнялись.",
        ],
        "tz": [
            ("P1", "Склеить слэш/без слэша", "Выбрать формат из sitemap и настроить 301 для /catalog, /catalog/vitalgar, /products/39 и аналогичных URL."),
            ("P1", "Добавить canonical", "В шаблоне главной, каталога и product-разделов вывести canonical на чистый URL."),
            ("P2", "Уникализировать product SEO", "Для /products/39, /products/30, /products/31 нужны уникальные title и description."),
            ("P2", "Оставить один H1", "В шаблоне product-раздела второй H1 заменить на H2/H3 или заголовок блока."),
        ],
        "facts": [
            "/products/39 и /products/39/ отдают 200 OK.",
            "/products/39?p=1 отдает дубль с 200 OK.",
            "Редактор раздела 39 не содержит SEO-полей.",
        ],
    },
]


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_margins(cell, top=80, start=120, bottom=80, end=120):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for m, value in {"top": top, "start": start, "bottom": bottom, "end": end}.items():
        node = tc_mar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_cell_text(cell, text, bold=False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.bold = bold
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    set_cell_margins(cell)


def add_bullets(doc, items):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(item)


def set_table_width(table, widths):
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_w = tbl_pr.first_child_found_in("w:tblW")
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), "9360")
    tbl_ind = tbl_pr.first_child_found_in("w:tblInd")
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:type"), "dxa")
    tbl_ind.set(qn("w:w"), "120")
    grid = tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(width))
        grid.append(col)


def add_tz_table(doc, rows):
    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.style = "Table Grid"
    widths = [900, 2460, 6000]
    set_table_width(table, widths)
    headers = ["Приоритет", "Задача", "Что сделать"]
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_text(cell, header, bold=True)
        set_cell_shading(cell, "E8EEF5")
    for priority, title, action in rows:
        row = table.add_row()
        values = [priority, title, action]
        for i, value in enumerate(values):
            set_cell_text(row.cells[i], value)


def configure_document(doc):
    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.header_distance = Inches(0.492)
    section.footer_distance = Inches(0.492)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.1

    for name, size, color, before, after in [
        ("Heading 1", 16, "2E74B5", 16, 8),
        ("Heading 2", 13, "2E74B5", 12, 6),
        ("Heading 3", 12, "1F4D78", 8, 4),
    ]:
        style = styles[name]
        style.font.name = "Calibri"
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor.from_string(color)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)

    bullet = styles["List Bullet"]
    bullet.font.name = "Calibri"
    bullet.font.size = Pt(11)
    bullet.paragraph_format.space_after = Pt(4)
    bullet.paragraph_format.line_spacing = 1.25


def add_footer(doc):
    footer = doc.sections[0].footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(f"SEO ТЗ, {TODAY}")
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(90, 90, 90)


def build_report(data):
    doc = Document()
    configure_document(doc)
    add_footer(doc)

    title = doc.add_paragraph()
    title.paragraph_format.space_after = Pt(3)
    run = title.add_run(f"SEO ТЗ: {data['project']}")
    run.font.name = "Calibri"
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor.from_string("1F4D78")
    run.bold = True

    meta = doc.add_paragraph()
    meta.add_run("Сайт: ").bold = True
    meta.add_run(data["url"])
    meta.add_run(" | Дата: ").bold = True
    meta.add_run(TODAY)

    doc.add_heading("Задача", level=1)
    doc.add_paragraph(TASK_SOURCE)

    doc.add_heading("Откуда взяты данные", level=1)
    doc.add_paragraph(DATA_SOURCE)

    doc.add_heading("Что сделано", level=1)
    add_bullets(doc, data["done"])

    doc.add_heading("Факты проверки", level=1)
    add_bullets(doc, data["facts"])

    doc.add_heading("ТЗ на доработку", level=1)
    add_tz_table(doc, data["tz"])

    doc.add_heading("Приемка после доработки", level=1)
    add_bullets(doc, COMMON_ACCEPTANCE)

    path = DOCX_DIR / f"SEO ТЗ - {data['project']}.docx"
    doc.save(path)
    return path


def main():
    DOCX_DIR.mkdir(parents=True, exist_ok=True)
    created = [build_report(report) for report in REPORTS]
    for path in created:
        print(path)


if __name__ == "__main__":
    main()
