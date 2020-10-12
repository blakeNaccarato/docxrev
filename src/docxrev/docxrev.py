"""
Microsoft Word review tools (comments, markup, etc.) with Python
"""

import shutil

import win32com.client

# * -------------------------------------------------------------------------------- * #
# * SETUP
# * Get the instance of Word running on this machine. Start it if necessary.

try:
    WORD = win32com.client.gencache.EnsureDispatch("Word.Application")
except AttributeError:
    # We end up here if this cryptic error occurs.
    # https://stackoverflow.com/questions/52889704/python-win32com-excel-com-model-started-generating-errors
    # https://mail.python.org/pipermail/python-win32/2007-August/006147.html
    shutil.rmtree(win32com.__gen_path__)
    WORD = win32com.client.gencache.EnsureDispatch("Word.Application")

# * -------------------------------------------------------------------------------- * #
# * FUNCTIONS * #
