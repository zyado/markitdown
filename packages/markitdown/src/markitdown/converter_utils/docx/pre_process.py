import zipfile
from io import BytesIO
from typing import BinaryIO
from xml.etree import ElementTree as ET

from bs4 import BeautifulSoup, Tag

from .math.omml import OMML_NS, oMath2Latex

MATH_ROOT_TEMPLATE = "".join(
    (
        "<w:document ",
        'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ',
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ',
        'xmlns:o="urn:schemas-microsoft-com:office:office" ',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ',
        'xmlns:v="urn:schemas-microsoft-com:vml" ',
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ',
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ',
        'xmlns:w10="urn:schemas-microsoft-com:office:word" ',
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ',
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ',
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ',
        'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" ',
        'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" ',
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14">',
        "{0}</w:document>",
    )
)


def _convert_omath_to_latex(tag: Tag) -> str:
    """
    Converts an OMML (Office Math Markup Language) tag to LaTeX format.

    Args:
        tag (Tag): A BeautifulSoup Tag object representing the OMML element.

    Returns:
        str: The LaTeX representation of the OMML element.
    """
    # Format the tag into a complete XML document string
    math_root = ET.fromstring(MATH_ROOT_TEMPLATE.format(str(tag)))
    # Find the 'oMath' element within the XML document
    math_element = math_root.find(OMML_NS + "oMath")
    # Convert the 'oMath' element to LaTeX using the oMath2Latex function
    latex = oMath2Latex(math_element).latex
    return latex


def _get_omath_tag_replacement(tag: Tag, block: bool = False) -> Tag:
    """
    Creates a replacement tag for an OMML (Office Math Markup Language) element.

    Args:
        tag (Tag): A BeautifulSoup Tag object representing the "oMath" element.
        block (bool, optional): If True, the LaTeX will be wrapped in double dollar signs for block mode. Defaults to False.

    Returns:
        Tag: A BeautifulSoup Tag object representing the replacement element.
    """
    t_tag = Tag(name="w:t")
    t_tag.string = (
        f"$${_convert_omath_to_latex(tag)}$$"
        if block
        else f"${_convert_omath_to_latex(tag)}$"
    )
    r_tag = Tag(name="w:r")
    r_tag.append(t_tag)
    return r_tag


def _replace_equations(tag: Tag):
    """
    Replaces OMML (Office Math Markup Language) elements with their LaTeX equivalents.

    Args:
        tag (Tag): A BeautifulSoup Tag object representing the OMML element. Could be either "oMathPara" or "oMath".

    Raises:
        ValueError: If the tag is not supported.
    """
    if tag.name == "oMathPara":
        # Create a new paragraph tag
        p_tag = Tag(name="w:p")
        # Replace each 'oMath' child tag with its LaTeX equivalent as block equations
        for child_tag in tag.find_all("oMath"):
            p_tag.append(_get_omath_tag_replacement(child_tag, block=True))
        # Replace the original 'oMathPara' tag with the new paragraph tag
        tag.replace_with(p_tag)
    elif tag.name == "oMath":
        # Replace the 'oMath' tag with its LaTeX equivalent as inline equation
        tag.replace_with(_get_omath_tag_replacement(tag, block=False))
    else:
        raise ValueError(f"Not supported tag: {tag.name}")


def _pre_process_math(content: bytes) -> bytes:
    """
    Pre-processes the math content in a DOCX -> XML file by converting OMML (Office Math Markup Language) elements to LaTeX.
    This preprocessed content can be directly replaced in the DOCX file -> XMLs.

    Args:
        content (bytes): The XML content of the DOCX file as bytes.

    Returns:
        bytes: The processed content with OMML elements replaced by their LaTeX equivalents, encoded as bytes.
    """
    soup = BeautifulSoup(content.decode(), features="xml")
    for tag in soup.find_all("oMathPara"):
        _replace_equations(tag)
    for tag in soup.find_all("oMath"):
        _replace_equations(tag)
    return str(soup).encode()


# CriticMarkup delimiters (https://criticmarkup.com) used to show tracked changes
INSERTION_MARKERS = ("{++", "++}")
DELETION_MARKERS = ("{--", "--}")

# Revision elements and whether they represent inserted (True) or deleted (False) text
_REVISION_TAGS = {"ins": True, "moveTo": True, "del": False, "moveFrom": False}


def _make_text_run(text: str) -> Tag:
    """Creates a plain <w:r><w:t>text</w:t></w:r> run."""
    t_tag = Tag(name="w:t", attrs={"xml:space": "preserve"})
    t_tag.string = text
    r_tag = Tag(name="w:r")
    r_tag.append(t_tag)
    return r_tag


def _mark_revision(tag: Tag, inserted: bool) -> None:
    """
    Wraps the runs of a single tracked change (<w:ins>, <w:del>, <w:moveTo> or <w:moveFrom>)
    in CriticMarkup delimiters, and turns deleted text into ordinary text so it is not dropped.
    """
    runs = tag.find_all("r")
    if not runs:
        # Revision marks with no runs (e.g. an inserted paragraph mark) carry no text
        return

    if not inserted:
        # Deleted text lives in <w:delText>/<w:delInstrText>, which converters skip
        for del_text in tag.find_all("delText"):
            del_text.name = "t"
        for del_instr in tag.find_all("delInstrText"):
            del_instr.name = "instrText"

    if not any(t.get_text() for t in tag.find_all("t")):
        return

    opening, closing = INSERTION_MARKERS if inserted else DELETION_MARKERS
    runs[0].insert_before(_make_text_run(opening))
    runs[-1].insert_after(_make_text_run(closing))

    if not inserted:
        # Converters skip <w:del>/<w:moveFrom> entirely, so lift the runs out of it
        tag.unwrap()


def _pre_process_tracked_changes(soup: BeautifulSoup) -> None:
    """
    Shows tracked changes inline: insertions as {++text++} and deletions as {--text--}.
    Without this step, insertions appear as ordinary text and deletions are silently dropped.
    """
    for name, inserted in _REVISION_TAGS.items():
        for tag in soup.find_all(name):
            # <w:ins>/<w:del> inside run or paragraph properties only flag formatting
            # or paragraph-mark changes, not text, so skip them
            if tag.parent is not None and tag.parent.name in ("rPr", "pPr"):
                continue
            _mark_revision(tag, inserted)


def _pre_process_xml(
    content: bytes, track_changes: bool = True, process_math: bool = True
) -> bytes:
    """Applies all enabled pre-processing steps to one DOCX XML part."""
    soup = BeautifulSoup(content.decode(), features="xml")
    if track_changes:
        _pre_process_tracked_changes(soup)
    if process_math:
        for tag in soup.find_all("oMathPara"):
            _replace_equations(tag)
        for tag in soup.find_all("oMath"):
            _replace_equations(tag)
    return str(soup).encode()


def pre_process_docx(input_docx: BinaryIO, track_changes: bool = True) -> BinaryIO:
    """
    Pre-processes a DOCX file with provided steps.

    The process works by unzipping the DOCX file in memory, transforming specific XML files
    (such as converting OMML elements to LaTeX and marking tracked changes), and then
    zipping everything back into a DOCX file without writing to disk.

    Args:
        input_docx (BinaryIO): A binary input stream representing the DOCX file.
        track_changes (bool, optional): If True, tracked insertions and deletions are shown
            with CriticMarkup ({++inserted++}, {--deleted--}). Defaults to True.

    Returns:
        BinaryIO: A binary output stream representing the processed DOCX file.
    """
    output_docx = BytesIO()
    # The files that need to be pre-processed from .docx
    pre_process_enable_files = [
        "word/document.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
    ]
    with zipfile.ZipFile(input_docx, mode="r") as zip_input:
        files = {name: zip_input.read(name) for name in zip_input.namelist()}
        with zipfile.ZipFile(output_docx, mode="w") as zip_output:
            zip_output.comment = zip_input.comment
            for name, content in files.items():
                if name in pre_process_enable_files:
                    # Try all steps together, then each step alone, so one failing
                    # step (e.g. an unusual equation) doesn't disable the others
                    attempts = [
                        dict(track_changes=track_changes, process_math=True),
                        dict(track_changes=False, process_math=True),
                        dict(track_changes=track_changes, process_math=False),
                    ]
                    updated_content = content
                    for options in attempts:
                        if not (options["track_changes"] or options["process_math"]):
                            continue
                        try:
                            updated_content = _pre_process_xml(content, **options)
                            break
                        except Exception:
                            # If there is an error in processing the content, try the next option
                            continue
                    zip_output.writestr(name, updated_content)
                else:
                    zip_output.writestr(name, content)
    output_docx.seek(0)
    return output_docx
