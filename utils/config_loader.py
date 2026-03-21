import yaml

class ConfigLoader:
      def __init__(self, path):
            with open(path, encoding="utf-8") as f:
                  self.config = yaml.safe_load(f)
      
      def get_attr(self, key, default=None):
            return self.config.get(key, default)