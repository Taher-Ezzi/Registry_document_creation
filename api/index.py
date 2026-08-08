import os
import sys

# Add parent directory to sys.path so imports work properly
parent_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

from server import app

# Vercel WSGI entry point
app = app
