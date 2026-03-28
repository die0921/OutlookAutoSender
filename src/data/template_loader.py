import yaml
from typing import Dict
from src.models.template import Template, BodyPart, DataSource


class TemplateLoader:
    """从 YAML 文件加载邮件模板"""

    def __init__(self, templates_file: str):
        self.templates_file = templates_file

    def load_all(self) -> Dict[str, Template]:
        """加载所有模板，返回模板名称到 Template 对象的映射"""
        with open(self.templates_file, 'r', encoding='utf-8') as f:
            data = yaml.safe_load(f)

        templates = {}
        for name, tmpl_data in data.get('templates', {}).items():
            templates[name] = self._parse_template(name, tmpl_data)

        return templates

    def load(self, template_name: str) -> Template:
        """加载指定名称的模板"""
        all_templates = self.load_all()
        if template_name not in all_templates:
            raise KeyError(f"模板不存在: {template_name}")
        return all_templates[template_name]

    def _parse_template(self, name: str, data: dict) -> Template:
        """解析模板 YAML 数据为 Template 对象"""
        body_parts = [
            self._parse_body_part(part)
            for part in data.get('body_parts', [])
        ]

        data_sources = {}
        for ds_name, ds_data in data.get('data_sources', {}).items():
            data_sources[ds_name] = DataSource(name=ds_name, file=ds_data['file'])

        return Template(
            name=name,
            subject=data.get('subject', ''),
            body_type=data.get('body_type', 'text'),
            body_parts=body_parts,
            data_sources=data_sources,
        )

    def _parse_body_part(self, data: dict) -> BodyPart:
        """解析正文部分"""
        return BodyPart(
            type=data.get('type', 'text'),
            content=data.get('content'),
            source=data.get('source'),
            sheet=data.get('sheet'),
            range=data.get('range'),
            format=data.get('format'),
            include_header=data.get('include_header'),
        )
