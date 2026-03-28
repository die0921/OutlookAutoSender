import pytest
from src.data.template_loader import TemplateLoader


SIMPLE_YAML = """
templates:
  简单模板:
    subject: "{{主题}}"
    body_type: "text"
    body_parts:
      - type: "text"
        content: "{{正文}}"

global_attachments: []
"""

COMPLEX_YAML = """
templates:
  月报模板:
    subject: "{{月份}}月度报告 - {{客户名称}}"
    body_type: "html"
    body_parts:
      - type: "text"
        content: "尊敬的 {{客户名称}}："
      - type: "excel_range"
        source: "data"
        sheet: "汇总表"
        range: "A1:E10"
        format: "table"
        include_header: true
    data_sources:
      data:
        file: "{{数据文件路径}}"

  简单模板:
    subject: "{{主题}}"
    body_type: "text"
    body_parts:
      - type: "text"
        content: "{{正文}}"

global_attachments: []
"""


@pytest.fixture
def simple_templates_file(tmp_path):
    f = tmp_path / "templates.yaml"
    f.write_text(SIMPLE_YAML, encoding='utf-8')
    return str(f)


@pytest.fixture
def complex_templates_file(tmp_path):
    f = tmp_path / "templates.yaml"
    f.write_text(COMPLEX_YAML, encoding='utf-8')
    return str(f)


def test_load_simple_template(simple_templates_file):
    loader = TemplateLoader(simple_templates_file)
    template = loader.load("简单模板")

    assert template.name == "简单模板"
    assert template.subject == "{{主题}}"
    assert template.body_type == "text"
    assert len(template.body_parts) == 1
    assert template.body_parts[0].content == "{{正文}}"


def test_load_all_templates(complex_templates_file):
    loader = TemplateLoader(complex_templates_file)
    templates = loader.load_all()

    assert len(templates) == 2
    assert "月报模板" in templates
    assert "简单模板" in templates


def test_load_complex_template(complex_templates_file):
    loader = TemplateLoader(complex_templates_file)
    template = loader.load("月报模板")

    assert template.subject == "{{月份}}月度报告 - {{客户名称}}"
    assert template.body_type == "html"
    assert len(template.body_parts) == 2
    assert template.body_parts[1].type == "excel_range"
    assert template.body_parts[1].sheet == "汇总表"
    assert template.body_parts[1].include_header is True


def test_load_template_data_sources(complex_templates_file):
    loader = TemplateLoader(complex_templates_file)
    template = loader.load("月报模板")

    assert "data" in template.data_sources
    assert template.data_sources["data"].file == "{{数据文件路径}}"


def test_load_missing_template(simple_templates_file):
    loader = TemplateLoader(simple_templates_file)
    with pytest.raises(KeyError, match="模板不存在"):
        loader.load("不存在的模板")


def test_load_missing_file():
    loader = TemplateLoader("nonexistent.yaml")
    with pytest.raises(FileNotFoundError):
        loader.load_all()
