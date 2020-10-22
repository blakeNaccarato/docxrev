# pylint: disable-all

import pathlib
import sys

project = "docxrev"
author = "Blake Naccarato"
copyright = f"2020, {author}"

sys.path.insert(0, str(pathlib.Path(r"..\src").resolve()))

extensions = ["sphinx.ext.napoleon", "sphinx.ext.autodoc"]

exclude_patterns = ["docs", "Thumbs.db", ".DS_Store"]

add_module_names = False
autodoc_member_order = "bysource"
autodoc_typehints = "description"

napoleon_include_init_with_doc = True
napoleon_include_private_with_doc = True
napoleon_include_special_with_doc = True

html_theme = "sphinx_rtd_theme"
