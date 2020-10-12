def main(pattern: str, directory: str = CWD):
"""Print Word document review comments matching a regular expression.

Search for the regular expression `pattern` in all Word documents in the current working
directory, or `directory` if supplied.

CLI powered by [Python Fire](https://google.github.io/python-fire/guide/).

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
"""

import os
import re
from glob import glob

import fire

import docxrev

_CWD = os.getcwd()


def _search_comments(pattern: str, directory: str = _CWD):
    """Print Word document review comments matching a regular expression.

    Search for the regular expression `pattern` in all Word documents in the current
    working directory, or `directory` if supplied.

    Parameters
    ----------
    pattern: str
        Regular expression to search for in comments.
    directory: str, optional
        Directory in which to search. Default is current working directory.
    """

    if directory is not _CWD:
        os.chdir(directory)

    documents_in_directory = [os.path.abspath(path) for path in glob("[!~$]*.docx")]
    already_open_documents = [document.Name for document in docxrev.WORD.Documents]

    print()  # make newline between the shell command and the first result

    for document in documents_in_directory:

        document_name = os.path.basename(document)

        if document_name in already_open_documents:
            docxrev.WORD.Documents(document).Activate()
        else:
            docxrev.WORD.Documents.Open(document)

        active_doc = docxrev.WORD.ActiveDocument
        comments = [comment.Range.Text for comment in active_doc.Comments]

        for comment in comments:
            if re.search(pattern, comment):
                print(
                    docxrev.WORD.ActiveDocument.Name,
                    comment.replace("\r", "\t"),
                    sep="\n",
                    end="\n" * 2,
                )

        if document_name not in already_open_documents:
            docxrev.WORD.ActiveDocument.Close()


if __name__ == "__main__":
    fire.Fire(main)
