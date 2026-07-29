import importlib.util, sys, os
spec = importlib.util.spec_from_file_location("dashboard", os.path.join(os.path.dirname(__file__), "VIVA 1960 Dashboard.py"))
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
app = mod.app
