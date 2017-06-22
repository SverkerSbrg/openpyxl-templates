import unittest
from os.path import dirname

loader = unittest.TestLoader()
start_dir = dirname(__file__)
suite = loader.discover(start_dir)

runner = unittest.TextTestRunner()
runner.run(suite)