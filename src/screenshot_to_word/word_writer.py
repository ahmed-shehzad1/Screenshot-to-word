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

    def _word_and_document(self):
        try:
            word = win32com.client.GetActiveObject("Word.Application")
            if word.Documents.Count == 0:
                raise WordAutomationError("Open a Word document before inserting content.")
            return word, word.ActiveDocument
        except WordAutomationError:
            raise
        except Exception as exc:
            raise WordAutomationError(
                "Microsoft Word is not open, or no accessible document is active. Open Word and your document first."
            ) from exc

    def active_document_name(self) -> str:
        _, document = self._word_and_document()
        return str(document.Name)

    def append_image(self, image, max_width_inches: float = 6.3) -> str:
        """Append an embedded image to the active document and return its name."""
        temp_path = None
        try:
            _, document = self._word_and_document()
            document_name = str(document.Name)
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
                temp_path = temp_file.name
            image.save(temp_path, format="PNG")

            end_pos = max(0, int(document.Content.End) - 1)
            document.Range(end_pos, end_pos).InsertAfter("\r")
            insert_pos = max(0, int(document.Content.End) - 1)
            shape = document.InlineShapes.AddPicture(
                FileName=str(Path(temp_path).resolve()),
                LinkToFile=False,
                SaveWithDocument=True,
                Range=document.Range(insert_pos, insert_pos),
            )
            # Word uses points; preserve the image's aspect ratio.
            width_points = max(72.0, float(max_width_inches) * 72.0)
            if float(shape.Width) > width_points:
                shape.LockAspectRatio = True
                shape.Width = width_points
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

    def append_heading(self, heading: str) -> str:
        """Append a formatted heading to the active Word document."""
        heading = heading.strip()
        if not heading:
            raise WordAutomationError("Type a task heading before inserting it.")
        try:
            _, document = self._word_and_document()
            document_name = str(document.Name)
            position = max(0, int(document.Content.End) - 1)
            target = document.Range(position, position)
            target.InsertAfter(heading + "\r")
            # Format the paragraph just added as Word's Heading 2 style.
            start = max(0, int(document.Content.End) - len(heading) - 2)
            heading_range = document.Range(start, start + len(heading))
            heading_range.Style = "Heading 2"
            return document_name
        except WordAutomationError:
            raise
        except Exception as exc:
            raise WordAutomationError(f"Could not insert heading into Word: {exc}") from exc
