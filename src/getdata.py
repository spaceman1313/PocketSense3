# Copyright (c) 2026 spaceman1313. All rights reserved.
# Use of this source code is governed by the MIT License that can be found in the
# LICENSE file.

"""
getdata.py - Retrieve statements, stock, and fund data for PocketSense3
This module provides functions to download, import, and process OFX files for financial
accounts, as well as retrieve stock and fund quotes. It supports both interactive and
automated modes, handles account decryption, manages file imports, and sends processed
data to Microsoft Money. The module includes logic to prevent account lockouts, scrub
imported files, combine OFX files, and prompt users for various actions. Logging is
integrated for tracking operations and errors.

Main Features:
- Download OFX statements for Direct Connect accounts
- Import and process OFX files from a designated import directory
- Retrieve and process stock/fund quotes
- Scrub OFX files to ensure compatibility with Microsoft Money
- Combine OFX files if configured
- Send processed files to Microsoft Money
- Interactive and non-interactive operation modes
- Logging of all major actions and errors
- Python version compatibility check

Intended to be called directly from the command line or a batch file.

Pocketsense scripts originally by Robert:
http://sites.google.com/site/pocketsense/

This is an updated version ported to Python3 and modified to better comply with Python
best practices and make future expandability easier.
https://github.com/spaceman1313/PocketSense3
"""


# Define the required minimum Python version.  Make check compatible with Python 2 since
# a lot of users are coming from 2.7.15.  MUST do this before any other imports to
# prevent syntax errors when run on older python versions.
import sys
REQUIRED_MAJOR = 3
REQUIRED_MINOR = 11

if sys.version_info < (REQUIRED_MAJOR, REQUIRED_MINOR):
    # Construct a friendly, informative error message
    # pylint: disable=consider-using-f-string
    error_message = (
        "Pocketsense3 requires Python version %d.%d or higher. You are currently "
        "using Python version %d.%d. Please upgrade your Python installation."
        % (REQUIRED_MAJOR, REQUIRED_MINOR, sys.version_info[0], sys.version_info[1])
        )
    # pylint: enable=consider-using-f-string
    # Use RuntimeError for clarity to non-Python users
    raise RuntimeError(error_message)

# Now import modules as we normally would
# pylint: disable=wrong-import-position, wildcard-import, unused-wildcard-import
import time
import re
import webbrowser
from pathlib import Path

import ofx_online
import quotes
import site_cfg
import scrubber
from getdata_orchestrator import GetDataOrchestrator

from control2 import *
from rlib1 import *
# pylint: enable=wrong-import-position, wildcard-import, unused-wildcard-import

# Setup globals
userdat = site_cfg.site_cfg()
log = create_logger('root', 'getdata.log')
if Debug:
    logging.basicConfig(level=logging.DEBUG)
    log.warning("**DEBUG Enabled: See Control2.py to disable.")
    log.debug('xfrdir = %s', xfrdir.name)


def get_site(ofx: str) -> str:
    """
    Returns the site configuration entry for an OFX file.

    Returns the appropriate site configuration entry for an OFX file based on the FID
    and BANKID values found in the OFX file.  If a matching site is not found, returns
    an empty string.

    Args:
        ofx (str): OFX file content as a string.

    Returns:
        A string containing the site configuration entry name.
    """

    # Get <FID> and <BANKID> values from OFX file, if they exist
    p = re.compile(r'<FID>(.*?)[<\s]', re.IGNORECASE | re.DOTALL)
    r = p.search(ofx)
    fid = r.groups()[0] if r else 'undefined'
    p = re.compile(r'<BANKID>(.*?)[<\s]', re.IGNORECASE | re.DOTALL)
    r = p.search(ofx)
    bankid = r.groups()[0] if r else 'undefined'

    # Try to find a matching site entry based on FID or BANKID.  If a match isn't found,
    # site will be set to the first entry in sites.dat
    site = ""
    sites = userdat.sites
    if fid or bankid:
        for key, value in sites.items():
            # Check for match on FID or BANKID
            if FieldVal(value, 'fid') == fid or FieldVal(value, 'bankid') == bankid:
                site = key
                log.info('Matched import file to site *%s*', key)
                break

    return site


def get_acctid(ofx: str) -> str:
    """
    Returns the <ACCTID> field of an OFX file.

    Returns the <ACCTID> field of an OFX file.  If the field is not found, returns an
    empty string.

    Args:
        ofx (str): OFX file content as a string.

    Returns:
        A str containing the account ID found in the <ACCTID> field of the OFX file, or
        an empty string if the field is not found.
    """

    result = re.search(r"<ACCTID>([0-9]+)", ofx)

    if result is None:
        # No match return blank account
        acctid = ""
    else:
        # Return match
        acctid = result.group(1)

    return acctid


def get_directconnect_ofx_files(acct_array: list) -> tuple[bool, list]:
    """
    Retrieves OFX files for all Direct Connect accounts.

    Retrieves OFX files for all Direct Connect accounts defined in acct_array.
    If an account fails to connect, it is skipped for the remainder of the session
    to help prevent account lockouts.

    Args:
        acct_array (list): List of account information.

    Returns:
        A 2-element tuple containing:
            - bool: Overall status of the download operations.
            - list: List of downloaded OFX files, where each item is a list:
                [sitename, account#, filename, "DirectConnect"].
    """
    # Get download interval, if promptInterval=Yes in sites.dat
    # should move into get ofx function
    dl_interval = userdat.defaultInterval
    if userdat.promptInterval:
        try:
            p = int2(input("Download interval (days) [" + str(dl_interval) + "]: "))
            if p > 0:
                dl_interval = p
            else:
                raise ValueError("Download interval must be a positive integer.")
        except ValueError:
            log.info("Invalid entry. Using defaultInterval=%s", dl_interval)

    log.info("Download interval= %s days", dl_interval)

    # Verify that account info exists
    if len(acct_array) == 0:
        log.info("No accounts have been configured. Run SETUP.PY to add accounts")
        return False, []

    # Track [sitename, username] for failed connections so we don't risk locking an
    # account
    bad_connects = []

    # Initialize output variables
    ofx_list = []
    result = True

    # Process accounts
    for acct in acct_array:
        if [acct[0], acct[3]] not in bad_connects:
            # Carry out the online OFX download
            status, ofx_file = ofx_online.get_dc_OFX(acct, dl_interval)
            if status:
                # Add account and OFX file to output list
                ofx_list.append([acct[0], acct[1], ofx_file, "DirectConnect"])
            else:
                if userdat.skipFailedLogon:
                    # Log and skip any further attempts for this site/user combo
                    bad_connects.append([acct[0], acct[3]])

            # Update overall result status
            result = result and status

    return result, ofx_list


def get_import_ofx_files() -> tuple[bool, list]:
    """
    Retrieves OFX files in the `import` directory

    Retrieves all OFX files in the `import` directory. If a file looks like a valid OFX
    file, it is processed and moved to the `xfrdir` directory.

    Returns:
        A 2-element tuple containing:
            - bool: Overall status of the download operations.
            - list: List of downloaded OFX files.
    """

    log.info('Searching %s for statements to import', importdir)

    # Initialize output variables
    ofx_list = []
    result = True

    # Process all files in import folder [manual user downloaded files].
    # Include anything that looks like a valid ofx file regardless of extension.

    for f in importdir.glob("*.*"):

        # Come up with a friendly display name for logging ("parent/filename.ext")
        in_displayname = str(Path(f.parent.name)/f.name)

        # Read the file
        with open(f, encoding='utf-8') as ifile:
            dat = ifile.read()

        # Only process if it looks like an ofx file
        if validOFX(dat) == '':

            # Try to match file to an entry in sites.dat
            site = get_site(dat)

            # Get the account ID
            account_id = get_acctid(dat)

            # Preserve original file type but save w/ ofx extension
            outname =  f.name + ('' if f.suffix == ".ofx" else ".ofx")

            # Move file to xfrdir
            outpath = xfrdir / outname
            f.rename(outpath)

            # Add to list of OFX files to process
            ofx_list.append([site, account_id, str(outpath), "ImportFile"])

            # Log move
            out_displayname = str(Path(outpath.parent.name)/outpath.name)
            log.info("%s saved to %s", in_displayname, out_displayname)
        else:
            result = False
            log.info("%s does not appear to be a valid OFX file. Skipping.",
                      in_displayname)

    return result, ofx_list


def scrub_files(ofx_list: list, orchestrator: GetDataOrchestrator) -> None:
    """
    Scrubs all OFX files that have been collected.

    Runs the scrubber routine on each OFX file that has been collected.  Option is
    given for the user to scrub one at a time, provided the interactive flag is set.

    Args:
        ofx_list (list): List containing all OFX files to send to Money.
        orchestrator (GetDataOrchestrator): The orchestrator instance for handling user
        interactions.
    """

    # Determine if user wants to scrub all files, and if so, whether they want
    # to confirm each file before scrubbing.
    if orchestrator.should_scrub_files():

        # If doing individual scrubs, warn user about potential for errors if they
        # choose not to scrub.
        if orchestrator.confirm_individual_scrub:
            log.info("Warning not scrubbing files may result in errors when "
                        "importing into MS Money.")

        # Let's scrub them.
        for entry in ofx_list:

            # Get filename and log file being scrubbed
            filename = entry[2]
            # Don't run scrubbers on quote files, they break them.
            if entry[0] == "Stock/Fund Quotes":
                continue

            # Check whether to send each file if user selected 'C'.
            if not orchestrator.should_scrub_this_file(filename):
                continue

            # Don't run scrubbers if we don't have a site match.
            if not entry[0]:
                log.info(
                    "Not matched to a site in sites.dat. Scrubber not run on %s.",
                    filename
                )
                continue

            # Proceed with Scrubbing
            print("")
            log.info("Scrubbing %s", filename)

            with open(filename, 'r', encoding='utf-8') as ifile:
                ofx = ifile.read()

                # Check to see if file has been scrubbed already, and if not scrub
                # it.
                if 'NEWFILEUID:PSIMPORT' not in ofx[:200]:
                    try:
                        scrubber.scrub(filename, userdat.sites[entry[0]])
                    except re.error:
                        log.info("Error running scrubbers on %s",  filename)

            # Set NEWFILEUID:PSIMPORT to flag the file as having already been
            # imported. Don't want to accidentally scrub twice
            with open(filename, 'r', encoding='utf-8') as ifile:
                ofx = ifile.read()

            pattern = re.compile(r'NEWFILEUID:.*')
            ofx2 = pattern.sub('NEWFILEUID:PSIMPORT', ofx)
            if ofx2:
                with open(filename, 'w', encoding='utf-8') as ofile:
                    ofile.write(ofx2)

    else:
        # Don't send to Scrub.
        log.info("OFX files not scrubbed (user selection or sites.dat setting).")
        log.info("Warning not scrubbing files may result in errors when "
                    "importing into MS Money.")


def send_files_to_money(ofx_list: list, quote_file_forced: str,
                        orchestrator: GetDataOrchestrator) -> None:
    """
    Sends all OFX files to Microsoft Money.

    Sends all OFX files in ofxList to Microsoft Money.  If the user has
    enabled the combineOFX option, all files are combined into a single file
    before sending.

    Args:
        ofx_list (list): List containing all OFX files to send to Money.
        quote_file_forced (str): Downloaded ForceQuotes OFX file.
        orchestrator (GetDataOrchestrator): The orchestrator instance for handling user
        interactions.
    """

    # Combine OFX files if option set
    cfile = combineOfx(ofx_list) if (userdat.combineofx and len(ofx_list) > 1) else ""

    # Determine if user wants to send results to Money, and if so, whether they want
    # to verify each file before sending.
    if orchestrator.should_send_to_money():

        # Send to Money
        log.debug('User confirmed upload to Money.')

        # Send ForceQuotes file to Money if defined (NEEDS CLEANUP)
        if quote_file_forced and Path(quote_file_forced).exists():
            if Debug:
                log.debug("Importing ForceQuotes statement: %s", quote_file_forced)
            run_file(quote_file_forced)  # Force transactions for MoneyUK
            input(
                "ForceQuote statement loaded.  Accept in Money and press <Enter> "
                "to continue."
            )

        # Send individual file(s) or combined file to Money
        print("")
        log.info("Sending statement(s) to Money...")

        if cfile and not orchestrator.confirm_individual_sendto:
            # Send combined file
            log.info("Importing combined OFX file: %s", cfile)
            run_file(cfile)

        else:
            # Send individual files, prompting for verification if user selected 'V'
            for ofxfile in ofx_list:
                # Upload each file one at a time, verify if requested
                if orchestrator.should_send_this_file(ofxfile[2]):
                    log.info("Importing %s", ofxfile[2])
                    run_file(ofxfile[2])

                # Slight delay, to force load order in Money
                time.sleep(0.5)

    else:
        # Don't send to Money
        log.info("Results not sent to Money (user selection or sites.dat setting).")


def main():
    """
    Main function when getdata.py is run directly

    Main function to orchestrate the retrieval of OFX files and quotes, and sending them
    to Money.
    """

    # Print version information at the start of the run
    print('')
    log.info("%s, Ver: %s", AboutTitle, AboutVersion)

    # Get account info
    # acct_array = [['SiteName', 'Account#', 'AcctType', 'UserName', 'PassWord'], ...]
    pwkey, _, acct_array = get_cfg()

    if len(acct_array) > 0 and pwkey != '':
        # if accounts are encrypted... decrypt them
        pwkey = decrypt_pw(pwkey)
        acct_array = acctDecrypt(acct_array, pwkey)

    # Delete old data files
    for old_file in xfrdir.glob("*.ofx"):
        old_file.unlink(missing_ok=True)

    # Determine if running in interactive mode.  If so, prompt user at each step.  If
    # not, use settings defined in sites.dat
    orchestrator = GetDataOrchestrator(userdat)
    orchestrator.set_interactive()

    # Create process Queue in the right order
    work_queue = ['Accts', 'importFiles']
    if userdat.savetickersfirst:
        work_queue.insert(0,'Quotes')
    else:
        work_queue.append('Quotes')

    # Initialize variables to track status and results across operations
    status = True   # Overall status across all operations
                    # True if all operations are successful
    ofx_list = []   # List of all OFX files to send to Money.
                    # Each item is a list: [SiteName, Account#, OFX filename]
    quote_file, quote_file_forced, html_quote_file = "", "", ""

    # Get and process files in the order defined in Queue
    for q_entry in work_queue:

        # Get OFX files for Direct Connect accounts
        if q_entry == 'Accts' and orchestrator.should_fetch_remote_accounts():
            log.info("\n--- Downloading statements for all accounts ---")
            accts_status, new_list = get_directconnect_ofx_files(acct_array)
            ofx_list.extend(new_list)
            status = status and accts_status

        # Get OFX files from import folder
        if q_entry == 'importFiles' and orchestrator.should_fetch_import_files():
            print("")
            log.info("--- Processing OFX files in import folder ---")
            import_status, new_list = get_import_ofx_files()
            ofx_list.extend(new_list)
            status = status and import_status

        # Get stock/fund quotes
        if q_entry == 'Quotes' and orchestrator.should_fetch_quotes():
            print("")
            log.info("--- Downloading stock/fund quotes ---")
            quote_status, quote_file, quote_file_forced, html_quote_file = (
                quotes.getQuotes())
            if quote_status:
                new_list = ['Stock/Fund Quotes', '', quote_file, "Quote"]
                ofx_list.append(new_list)
            status = status and quote_status

    # Scrub all OFX files and send to Money.
    if len(ofx_list) > 0:

        # Scrub ofx files, the interactive prompts get dealt inside the function.
        print("")
        log.info('--- Scrubbing files ---')
        scrub_files(ofx_list, orchestrator)

        # Send to money, the interactive prompts get dealt with inside the function.
        print("")
        log.info('--- Sending files to MS Money ---')
        send_files_to_money(ofx_list, quote_file_forced, orchestrator)

    else:
        log_message = (
            "No OFX files were fetched or created. Verify network connection "
            "and/or input folders.")
        log.warning(log_message)

        # Set status to False to prevent auto-close of command window
        status = False

    # Display quotes.htm if downloaded and user wants to see it.
    if html_quote_file and orchestrator.should_open_quotes_html(html_quote_file.name):
        log.debug("Full quotes file name: %s", html_quote_file)
        log.info("Opening %s in browser.", html_quote_file.name)
        webbrowser.open(html_quote_file.as_uri())  # don't wait for browser close

    # Keep window open if needed, complete execution
    if not status:
        log_message = (
            "One or more errors were detected during the processing of OFX files. "
            "Review display and log to identify the problem.")
        log.warning(log_message)

    if Debug or orchestrator.interactive or userdat.promptEnd or not status:
        input("\nPress <Enter> to close the window...")

    print("")
    log.info(
        '-----------------------------------------------------------------------------')


if __name__ == "__main__":
    main()
