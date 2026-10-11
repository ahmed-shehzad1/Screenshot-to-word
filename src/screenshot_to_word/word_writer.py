"""Microsoft Word automation (Windows + desktop Word only)."""
import os
import tempfile
from pathlib import Path

import win32com.client


class WordAutomationError(RuntimeError):
    pass


class WordWriter:
    def __init__(self):
        if os.name != "nt":
            raise WordAutomationError("Word automation is supported on Windows only.")

    def active_document_name(self) -> str:
        try:
            word = win32com.client.GetActiveObject("Word.Application")
            if word.Documents.Count == 0:
                return "No Word document open"
            return str(word.ActiveDocument.Name)
        except Exception as exc:
            raise WordAutomationError(
                "Microsoft Word is not open, or no accessible document is active."
            ) from exc

    def append_image(self, image) -> str:
        """Append an image to the active document and return its document name."""
        temp_path = None
        try:
            word = win32com.client.GetActiveObject("Word.Application")
            if word.Documents.Count == 0:
                raise WordAutomationError("Open a Word document before inserting screenshots.")

            document = word.ActiveDocument
            document_name = str(document.Name)

            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
                temp_path = temp_file.name
            image.save(temp_path, format="PNG")

            # Insert before Word's final paragraph mark, then add a new paragraph
            # after the picture so the next capture appears underneath it.
            end_pos = max(0, int(document.Content.End) - 1)
            paragraph_range = document.Range(end_pos, end_pos)
            paragraph_range.InsertAfter("\r")

            insert_pos = max(0, int(document.Content.End) - 1)
            insert_range = document.Range(insert_pos, insert_pos)
            document.InlineShapes.AddPicture(
                FileName=str(Path(temp_path).resolve()),
                LinkToFile=False,
                SaveWithDocument=True,
                Range=insert_range,
            )
            document.Range(document.Content.End - 1, document.Content.End - 1).InsertAfter("\r")
            return document_name
        except WordAutomationError:
            raise
        except Exception as exc:
            raise WordAutomationError(f"Could not insert screenshot into Word: {exc}") from exc
        finally:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass
