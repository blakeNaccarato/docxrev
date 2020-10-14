"""CLI to search comments."""

import os
import pathlib
import re
from typing import Union

import fire

import docxrev

CWD = os.getcwd()


def main(pattern: str, directory: Union[str, pathlib.Path] = CWD):
    """Print Word document review comments matching a regular expression.

    Search for the regular expression `pattern` in all Word documents in the current
    working directory, or `directory` if supplied.

    Examples
    --------
    Get help for the command.
        py search_commments.py --help

    Search for "word". Quotes are unnecessary.
        py search_comments.py word

    Search for "multiple words". Quotes are necessary.
        py search_commments.py "multiple words"

    Search all files in "C:\\Users\\You\\Desktop". Quotes are unnecessary.
        py search_commments.py word --directory C:\\Users\\You\\Desktop

    Search all files in "C:\\Users\\You\\Desktop\\Space In Path". Quotes are necessary.
        py search_commments.py word --directory "C:\\Users\\You\\Desktop\\Space In Path"

    Parameters
    ----------
    pattern: str
        Regular expression to search for in comments.
    directory: str, optional
        Directory in which to search. Default is current working directory.
    """

    directory = pathlib.Path(directory)
    paths = directory.glob("[!~$]*.docx")

    print()  # make newline between the shell command and the first result

    for path in paths:
        with docxrev.Document(path) as document:
            for comment in document.comments:
                if re.search(pattern, comment):
                    print(
                        document.name,
                        comment.replace("\r", "\t"),
                        sep="\n",
                        end="\n" * 2,
                    )


if __name__ == "__main__":
    fire.Fire(main)
