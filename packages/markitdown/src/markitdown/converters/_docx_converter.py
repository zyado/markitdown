import sys
import io
from warnings import warn

from typing import BinaryIO, Any

from ._html_converter import HtmlConverter
from ..converter_utils.docx.pre_process import pre_process_docx
from .._base_converter import DocumentConverterResult
from .._stream_info import StreamInfo
from .._exceptions import MissingDependencyException, MISSING_DEPENDENCY_MESSAGE

# Try loading optional (but in this case, required) dependencies
# Save reporting of any exceptions for later
_dependency_exc_info = None
try:
    import mammoth

except ImportError:
    # Preserve the error and stack trace for later
    _dependency_exc_info = sys.exc_info()


ACCEPTED_MIME_TYPE_PREFIXES = [
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
]

ACCEPTED_FILE_EXTENSIONS = [".docx"]

# Mammoth drops page breaks by default; render them as <hr> (Markdown "---")
PAGE_BREAK_STYLE_MAP = "br[type='page'] => hr"


class DocxConverter(HtmlConverter):
    """
    Converts DOCX files to Markdown. Style information (e.g.m headings) and tables are preserved where possible.
    """

    def __init__(self):
        super().__init__()
        self._html_converter = HtmlConverter()

    def accepts(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> bool:
        mimetype = (stream_info.mimetype or "").lower()
        extension = (stream_info.extension or "").lower()

        if extension in ACCEPTED_FILE_EXTENSIONS:
            return True

        for prefix in ACCEPTED_MIME_TYPE_PREFIXES:
            if mimetype.startswith(prefix):
                return True

        return False

    def convert(
        self,
        file_stream: BinaryIO,
        stream_info: StreamInfo,
        **kwargs: Any,  # Options to pass to the converter
    ) -> DocumentConverterResult:
        # Check: the dependencies
        if _dependency_exc_info is not None:
            raise MissingDependencyException(
                MISSING_DEPENDENCY_MESSAGE.format(
                    converter=type(self).__name__,
                    extension=".docx",
                    feature="docx",
                )
            ) from _dependency_exc_info[
                1
            ].with_traceback(  # type: ignore[union-attr]
                _dependency_exc_info[2]
            )

        style_map = kwargs.get("style_map", None)
        # Show tracked insertions/deletions as {++inserted++} / {--deleted--}
        track_changes = kwargs.get("docx_track_changes", True)
        # Show manual page breaks as horizontal rules (---)
        page_breaks = kwargs.get("docx_page_breaks", True)

        if page_breaks:
            # Mammoth uses the first matching rule, so a caller's own style_map wins
            style_map = (style_map or "") + "\n" + PAGE_BREAK_STYLE_MAP

        pre_process_stream = pre_process_docx(file_stream, track_changes=track_changes)
        html = mammoth.convert_to_html(
            pre_process_stream,
            style_map=style_map,
        ).value

        return self._html_converter.convert_string(
            html,
            **kwargs,
        )
