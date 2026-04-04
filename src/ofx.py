# Copyright (c) 2026 spaceman1313. All rights reserved.
# Use of this source code is governed by the MIT License that can be found in the
# LICENSE file.

"""
This is the start of a light weight OFX handling class module.

The intent is to capture various OFX handling routines spread throughout Pocketsense
into one class.  It will never be a full OFX library, like others available.
"""
from pathlib import Path
import re

class OFX:
    """
    A light weight OFX handling class.

    This class is intended to capture various OFX handling routines spread throughout
    Pocketsense into one class. It will never be a full OFX library, like others
    available.
    """

    def __init__(self, content: str)  -> None:

        self.content = content

    @classmethod
    def load_from_file(cls, filepath: str | Path)  -> "OFX":
        """
        Loads OFX content from a file.

        Args:
            filepath (str|Path): The path to the OFX file to load.
        """

        if isinstance(filepath, str):
            filepath = Path(filepath)

        if filepath.is_file():
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
        else:
            raise FileNotFoundError(f"File not found: {filepath}")

        return cls(content)

    def get_tag_value(self, tag: str) -> str | None:
        """
        Extracts the value of a specified OFX tag from the content.

        Args:
            tag (str): The name of the OFX tag to extract.

        Returns:
            str|None: The value of the specified tag, or None if the tag is not found.
        """

        # RegExp accounts for both <TAG>value</TAG> and <TAG>value formats.
        p = re.compile(rf'<{tag}>(.*?)</{tag}>|<{tag}>([^<]+)',
                             re.DOTALL | re.IGNORECASE)

        # Return the first match found, stripping any leading/trailing whitespace.  If
        # no match is found, return None.
        match = p.search(self.content)
        if match:
            return (match.group(1) or match.group(2)).strip()

        return None

    @staticmethod
    def _header_pattern(header: str) -> re.Pattern:
        """
        Compiles a regular expression pattern for matching a specified OFX header.

        Args:
            header (str): The name of the OFX header to create a pattern for.
        """

        return re.compile(rf'^({header}:\s*)(\S+)', re.IGNORECASE | re.MULTILINE)

    def get_header_value(self, header: str) -> str | None:
        """
        Extracts the value of a specified OFX header from the content.

        Args:
            header (str): The name of the OFX header to extract.

        Returns:
            str|None: The value of the specified header, or None if the header is not
            found.
        """

        # RegExp to look for header.
        p = self._header_pattern(header)

        # Return the first match found.  If no match is found, return None.
        match = p.search(self.content)

        if match:
            return match.group(2)

        return None

    def set_header_value(self, header: str, new_value: str) -> None:
        """
        Sets the value of a specified OFX header in the content.  If the header does
        not exist, a ValueError is raised.

        Args:
            header (str): The name of the OFX header to set.
            value (str): The value to set for the specified header.
        """

        p = self._header_pattern(header)

        # If a match is found, replace the existing value.  If no match is found,
        # add the header and value to the beginning of the content.
        if p.search(self.content):
            self.content = p.sub(rf'\g<1>{new_value}', self.content)
        else:
            raise ValueError(f"Header not found: {header}")
