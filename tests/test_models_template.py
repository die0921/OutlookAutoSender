from src.models.template import Template, BodyPart, DataSource


def test_template_creation():
    template = Template(
        name="测试模板",
        subject="{{主题}}",
        body_type="text",
        body_parts=[
            BodyPart(type="text", content="Hello {{name}}")
        ],
        data_sources={}
    )
    assert template.name == "测试模板"
    assert template.subject == "{{主题}}"
    assert len(template.body_parts) == 1
    assert template.body_parts[0].content == "Hello {{name}}"


def test_body_part_text():
    part = BodyPart(type="text", content="Hello World")
    assert part.type == "text"
    assert part.content == "Hello World"
    assert part.source is None


def test_body_part_excel_range():
    part = BodyPart(
        type="excel_range",
        source="data",
        sheet="汇总表",
        range="A1:E10",
        format="table",
        include_header=True
    )
    assert part.type == "excel_range"
    assert part.sheet == "汇总表"
    assert part.include_header is True


def test_data_source():
    ds = DataSource(name="data", file="report.xlsx")
    assert ds.name == "data"
    assert ds.file == "report.xlsx"


def test_template_with_data_sources():
    template = Template(
        name="月报模板",
        subject="{{月份}}月报",
        body_type="html",
        body_parts=[
            BodyPart(type="text", content="尊敬的用户："),
            BodyPart(type="excel_range", source="data", sheet="Sheet1", range="A1:C10", format="table")
        ],
        data_sources={"data": DataSource(name="data", file="{{数据文件路径}}")}
    )
    assert len(template.body_parts) == 2
    assert "data" in template.data_sources
