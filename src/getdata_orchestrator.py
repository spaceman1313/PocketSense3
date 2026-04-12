# Copyright (c) 2026 spaceman1313. All rights reserved.
# Use of this source code is governed by the MIT License that can be found in the
# LICENSE file.

"""
getdata_orchestrator.py - Encapsulates user interactions for getddata.py

Contains the GetDataOrchestrator class, which manages user interactions and decisions
regarding data retrieval and processing for PocketSense3. This class provides methods
to prompt the user for input with default values, and to determine whether to perform
specific actions based on user preferences and interactive mode settings.
"""

from typing import Any

class GetDataOrchestrator:
    """
    Orchestrates the overall data retrieval and processing workflow for PocketSense3.

    This class is responsible for managing user interactions and decisions regarding
    which data to fetch, process, and send to Money. It provides methods to prompt the
    user for input with default values, and to determine whether to perform specific
    actions based on user preferences and interactive mode settings.

    Attributes:
        userdat (list): Contains user configuration and preferences.
        interactive (bool): Indicates whether the program is running in interactive
            mode.
        confirm_individual_scrub (bool): Indicates whether to confirm scrubbing each
            file individually.
        confirm_individual_sendto (bool): Indicates whether to confirm sending each
            file to Money individually.
    """

    def input_default(self,
                      prompt: str, default: object, type_cast: type = str) -> Any:
        """
        Prompts the user for input with a default value.

        When type_cast is set to bool, accepts Y/N, Yes/No, True/False, T/F, 1/0
        (case-insensitive) as valid inputs.  In this case the default should be set to a
        user friendly value (e.g. 'Y' or 'N') rather than a Boolean value.

        Args:
            prompt (str): The prompt message to display to the user.
            default (any): The default value to use if the user provides no input.
            type_cast (type, optional): A function to cast the input to a specific type.
                Defaults to str.

        Returns:
            The value entered by the user, cast to the specified type, or the default
            value if no input is provided.
        """

        # Loop until valid input is received
        valid_input = False
        user_input = None
        while not valid_input:
            try:
                # Get the user input
                user_input = input(f"{prompt} [{default}]: ")

                # Process the special case of Boolean input
                if type_cast == bool:
                    user_input = str(default) if not user_input else user_input
                    if user_input.upper() in ['Y', 'YES', 'TRUE', 'T', '1']:
                        user_input = True
                    elif user_input.upper() in ['N', 'NO', 'FALSE', 'F', '0']:
                        user_input = False
                    else:
                        raise ValueError("Invalid Boolean input")

                # Process all other types
                else:
                    if user_input:
                        user_input = type_cast(user_input)
                    else:
                        user_input = default

                valid_input = True  # Exit the loop if input is valid

            except ValueError:
                # Only error we expect is a ValueError from an invalid type cast, so we
                # can catch that and prompt again
                print(f"Invalid entry. Please enter a value of type "
                      f"{type_cast.__name__}.")

        return user_input

    def __init__(self, userdat) -> None:
        """
        Initializes the GetDataOrchestrator with user data and default settings.

        Args:
            userdat (list): Contains user configuration and preferences.
        """

        self.userdat = userdat
        self.interactive = False
        self.confirm_individual_scrub = False
        self.confirm_individual_sendto = False

    def set_interactive(self) -> None:
        """
        Sets the interactive mode based on user request.
        """

        self.interactive = self.input_default(
            "Run in interactive mode? (y/n)", 'n', bool
        ) if self.userdat.promptStart else False

    def should_fetch_remote_accounts(self) -> bool:
        """
        Returns whether Direct Connect account statements should be downloaded.
        """

        result = self.input_default(
            "\nDownload statements for all online accounts? (y/n)",
            "y" if self.userdat.fetchRemote else "n",
            bool
        ) if self.interactive else self.userdat.fetchRemote

        return bool(result)

    def should_fetch_import_files(self) -> bool:
        """
        Returns whether downloaded account statements should be imported.
        """

        result = self.input_default(
            "\nProcess files in import folder? (y/n)",
            "y" if self.userdat.fetchImport else "n",
            bool
        ) if self.interactive else self.userdat.fetchImport

        return bool(result)

    def should_fetch_quotes(self) -> bool:
        """
        Returns whether Quotes should be downloaded.
        """

        result = self.input_default(
            "\nDownload stock/fund quotes? (y/n)",
            "y" if self.userdat.fetchQuotes else "n",
            bool
        ) if self.interactive else self.userdat.fetchQuotes

        return bool(result)

    def should_open_quotes_html(self, description: str) -> bool:
        """
        Returns whether the quotes HTML file should be opened in the default browser.
        """

        result = self.input_default(
            f"\nOpen {description} in the default browser? (y/n)",
            "y" if self.userdat.showquotehtm else "n",
            bool
        ) if self.interactive else self.userdat.showquotehtm

        return bool(result)

    def should_scrub_files(self) -> bool:
        """
        Returns whether to scrub all the OFX files.
        """

        defval = "y" if self.userdat.scrubOfx else "n"
        if self.interactive:
            userin = (
                self.input_default(
                    "\n\nScrub OFX files? (y/n/c=confirm)",
                    defval,
                    str)
            ).lower()

            if userin == 'c':
                self.confirm_individual_scrub = True
        else:
            userin = defval

        return userin in ["c", "y"]

    def should_scrub_this_file(self, description: str) -> bool:
        """
        Returns whether the specified file should be scrubbed.
        """

        if self.confirm_individual_scrub:
            result = self.input_default(
                f"Scrub file {description}? (y/n)",
                "y",
                bool
            )
            return bool(result)

        return True

    def should_send_to_money(self) -> bool:
        """
        Returns whether to send data to Money.
        """

        defval = "y" if self.userdat.sendToMoney else "n"
        if self.interactive:
            userin = (
                self.input_default(
                    "\n\nSend results to Money? (y/n/c=confirm)",
                    defval,
                    str
                )
            ).lower()

            if userin == 'c':
                self.confirm_individual_sendto = True
        else:
            userin = defval

        return userin in ["c", "y"]

    def should_send_this_file(self, description: str) -> bool:
        """
        Returns whether the specified file should be sent to Money.
        """

        if self.confirm_individual_sendto:
            result = self.input_default(
                f"Send file {description} to Money? (y/n)",
                "y",
                bool
            )
            return bool(result)

        return True
