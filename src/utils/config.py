import os
import configparser
import logging
from typing import Dict, Tuple, List

class ConfigManager:
    DEFAULT_CONFIG = {
        'Typography': {
            'body_font': '"仿宋"',
            'body_size': '18',
            'first_line_indent': '12',
            'line_spacing': '1',
            'min_heading_level': '6',
            'title_level_1_font': '"仿宋"',
            'title_level_1_size': '24',
            'title_level_2_font': '"仿宋"',
            'title_level_2_size': '20',
            'title_level_3_font': '"仿宋"',
            'title_level_3_size': '18',
            'title_level_4_font': '"仿宋"',
            'title_level_4_size': '18',
            'title_level_5_font': '"仿宋"',
            'title_level_5_size': '18',
            'title_level_6_font': '"仿宋"',
            'title_level_6_size': '18'
        }
    }

    def __init__(self, config_file: str):
        self.config_file = config_file
        self.config = configparser.ConfigParser()
        self.required_settings = [
            ('Typography', 'body_font'),
            ('Typography', 'body_size'),
            ('Typography', 'first_line_indent'),
            ('Typography', 'line_spacing'),
            ('Typography', 'min_heading_level')
        ]

    def load(self) -> configparser.ConfigParser:
        if not os.path.exists(self.config_file):
            raise FileNotFoundError(f"配置文件不存在: {self.config_file}")
        
        try:
            self.config.read(self.config_file, encoding='utf-8')
            self._validate_config()
            return self.config
        except Exception as e:
            logging.error(f"加载配置文件失败: {str(e)}")
            raise

    def _validate_config(self):
        for section, key in self.required_settings:
            if not self.config.has_option(section, key):
                raise ValueError(f"缺少必要的配置项: [{section}] {key}")

    def create_default_config(self) -> None:
        for section, values in self.DEFAULT_CONFIG.items():
            if not self.config.has_section(section):
                self.config[section] = {}
            for key, value in values.items():
                self.config[section][key] = value

        with open(self.config_file, 'w', encoding='utf-8') as configfile:
            self.config.write(configfile)

    def get_typography_settings(self) -> Dict:
        return dict(self.config['Typography'])

    def get_title_settings(self, level: int) -> Tuple[str, int]:
        font = self.config['Typography'].get(f'title_level_{level}_font', '"仿宋"')
        size = int(self.config['Typography'].get(f'title_level_{level}_size', '18'))
        return font, size