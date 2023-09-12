import os
import unittest
from odpmkd import *


class MyTest(unittest.TestCase):
    def setUp(self):
        try:
            data_dir = os.environ['DATA_DIR']
            self.data_dir = data_dir
        except KeyError:
            if os.path.exists("../examples"):
                self.data_dir = "../examples"
            elif os.path.exists("examples"):
                self.data_dir = "examples"
        self.output_dir = "/tmp"

    def test(self):
        parser = OdpParser()
        parser.open(os.path.join(self.data_dir, "simple.odp"),
                    os.path.join(self.output_dir, "media"), True, True)
