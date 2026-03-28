import pytest
from src.services.template_engine import TemplateEngine
from src.data.template_loader import TemplateLoader


SIMPLE_YAML = """
templates:
  简单模板:
    subject: "{{主题}}"
    body_type: "text"
    body_parts:
      - type: "text"
        content: "尊敬的 {{客户名称}}，{{正文}}"

  HTML模板:
    subject: "{{月份}}月报"
    body_type: "html"
    body_parts:
      - type: "text"
        content: "尊敬的 {{客户名称}}："
      - type: "text"
        content: "请查收报告。"

global_attachments: []
"""


@pytest.fixture
def template_file(tmp_path):
    f = tmp_path / "templates.yaml"
    f.write_text(SIMPLE_YAML, encoding='utf-8')
    return str(f)


@pytest.fixture
def engine(template_file):
    loader = TemplateLoader(template_file)
    return TemplateEngine(loader)


def test_render_simple_subject(engine):
    subject, _ = engine.render('简单模板', {'主题': '测试邮件', '客户名称': '张三', '正文': '内容'})
    assert subject == '测试邮件'


def test_render_simple_body(engine):
    _, body = engine.render('简单模板', {'主题': '测试', '客户名称': '张三', '正文': '内容'})
    assert '张三' in body
    assert '内容' in body


def test_render_missing_variable_silent(engine):
    """缺少变量时静默处理（不报错）"""
    subject, body = engine.render('简单模板', {'主题': '测试'})
    assert subject == '测试'
    # body should render without crashing even if vars missing


def test_render_html_template(engine):
    subject, body = engine.render('HTML模板', {'月份': '3', '客户名称': '李四'})
    assert subject == '3月报'
    assert '<html>' in body
    assert '李四' in body


def test_get_template_variables(engine):
    variables = engine.get_template_variables('简单模板')
    assert '主题' in variables
    assert '客户名称' in variables
    assert '正文' in variables


def test_render_unknown_template(engine):
    with pytest.raises(KeyError):
        engine.render('不存在的模板', {})
