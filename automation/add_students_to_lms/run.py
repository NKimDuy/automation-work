from modules.init_selenium import InitSelenium
from utils.config_loader import ConfigLoader
import os
import time 

class UpdateStudents:
      def __init__(self):
            config = ConfigLoader(os.path.join(os.path.dirname(__file__), "config.yaml"))
            self.url = f"https://lms.oude.edu.vn/{config.get_attr("SEMESTER")}/course/search.php"

      def run_selenium(self):
            start_selenium = InitSelenium()
            start_selenium.login_selenium(self.url)

test = UpdateStudents()
test.run_selenium()