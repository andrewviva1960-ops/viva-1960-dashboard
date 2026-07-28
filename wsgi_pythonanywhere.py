import sys, os

project_home = os.path.expanduser('~/viva-1960-dashboard')
if project_home not in sys.path:
    sys.path.insert(0, project_home)

from app import app as application
