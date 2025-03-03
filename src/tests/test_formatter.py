import unittest
import os
import docx
import sys
from pathlib import Path

# 添加项目根目录到 Python 路径
project_root = Path(__file__).parent.parent.parent
sys.path.append(str(project_root))

from src.main import load_config, format_word

class TestFormatter(unittest.TestCase):
    def setUp(self):
        self.test_dir = Path(__file__).parent
        self.test_config = self.test_dir / 'test_config.ini'
        self.test_input = self.test_dir / 'test_input.docx'
        self.test_output = self.test_dir / 'test_output.docx'
        
        # 创建测试配置文件
        with open(self.test_config, 'w', encoding='utf-8') as f:
            f.write('''[Typography]
body_font = "仿宋"
body_size = 18
first_line_indent = 12
line_spacing = 1
min_heading_level = 6
title_level_1_font = "仿宋"
title_level_1_size = 24''')
            
        # 创建测试文档
        doc = docx.Document()
        doc.add_paragraph('测试文本')
        doc.save(self.test_input)

    def tearDown(self):
        # 清理测试文件
        for file in [self.test_config, self.test_input, self.test_output]:
            if os.path.exists(file):
                os.remove(file)

    def test_load_config(self):
        config = load_config(self.test_config)
        self.assertEqual(config['Typography']['body_font'], '"仿宋"')
        self.assertEqual(config['Typography']['body_size'], '18')

    def test_format_word(self):
        config = load_config(self.test_config)
        format_word(self.test_input, self.test_output, config)
        self.assertTrue(os.path.exists(self.test_output))

if __name__ == '__main__':
    unittest.main()