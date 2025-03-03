import unittest
import os
import tempfile
from src.utils.config import ConfigManager

class TestConfigManager(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.config_file = os.path.join(self.temp_dir, 'test_config.ini')
        self.config_manager = ConfigManager(self.config_file)

    def tearDown(self):
        if os.path.exists(self.config_file):
            os.remove(self.config_file)
        os.rmdir(self.temp_dir)

    def test_create_default_config(self):
        self.config_manager.create_default_config()
        self.assertTrue(os.path.exists(self.config_file))
        
        config = self.config_manager.load()
        self.assertEqual(config['Typography']['body_font'], '"仿宋"')
        self.assertEqual(config['Typography']['body_size'], '18')

    def test_load_invalid_config(self):
        with open(self.config_file, 'w', encoding='utf-8') as f:
            f.write('[Invalid]\nkey=value')
        
        with self.assertRaises(ValueError):
            self.config_manager.load()

    def test_get_typography_settings(self):
        self.config_manager.create_default_config()
        settings = self.config_manager.get_typography_settings()
        
        self.assertIn('body_font', settings)
        self.assertIn('body_size', settings)
        self.assertEqual(settings['body_font'], '"仿宋"')

    def test_get_title_settings(self):
        self.config_manager.create_default_config()
        font, size = self.config_manager.get_title_settings(1)
        
        self.assertEqual(font, '"仿宋"')
        self.assertEqual(size, 24)

    def test_nonexistent_config(self):
        with self.assertRaises(FileNotFoundError):
            self.config_manager.load()

if __name__ == '__main__':
    unittest.main()