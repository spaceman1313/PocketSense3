# GetData.py
# http://sites.google.com/site/pocketsense/
# retrieve statements, stock and fund data
# Intial version: rlc: Feb-2010

# History
# ---------
# 11-Mar-2010*rlc
#   - Added "interactive" mode
#   - Download all statements and quotes before beginning upload to Money
#   - Allow stock quotes to be sent to Money before statements (option defined in sites.dat)
# 09-May-2010*rlc
#   - Download files in the order that they will be sent to Money so that file timestamps are in the same order
#   - Send data to Money using the os.system() call rather than os.startfile(), as this seems
#     to help force the order when sending files to Money (FIFO)
#   - Added logic to catch failed connections and server timeouts
#   - Added "About" title and version to start
# 05-Sep-2010*rlc
#   - Updated to support spaces in SiteName values in sites.dat
#   - Don't auto-close command window if any error is detected during download operations
# 04-Jan-2011*rlc
#   - Display quotes.htm after download if "ShowQuoteHTM: Yes" defined in sites.dat
#   - Ask to display quotes.htm after download if "ShowQuoteHTM: Yes" defined in sites.dat (overrides ShowQuoteHTM)
# 18-Jan-2011*rlc
#   - Added 0.5 s delay between "file starts", which sends an OFX file to Money
# 23Aug2012*rlc
#   - Added user option to change default download interval at runtime
#   - Added support for combineOFX
# 28Aug2013*rlc
#   - Added support for forceQuotes option
# 21Oct2013*rlc
#   - Modified forceQuote option to prompt for statement accept in Money before continuing
# 25Feb2014*rlc
#   - Bug fix for forceQuote option when the quote feature isn't being used
# 14May2018*rlc
#   - If an a connection fails for a specific user/pw combo, don't try other accounts during the session
#     Added to help prevent accounts getting locked when a user changes their password, has multiple
#     accounts at the institution, but forgot to update their account settings in Setup.
#16Sep2019*rlc
#   - Add support for ofx import from ./import subfolder.  any file present in ./import will be inspected,
#     and if it looks like a valid OFX file, will be processed the same as a downloaded statement (scrubbed, etc.)
#19Jun2023*rlc
#   - add logging
#14Dec2023*cgn
#   - Update to python3


# Define the required minimum Python version.  Make check compatible with Python 2 since
# a lot of users are coming from 2.7.15.  Need to do this before any other imports to
# prevent syntax errors on older versions.
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
    raise RuntimeError(error_message) # Use RuntimeError for clarity to non-Python users

# Now import modules as we normally would
# pylint: disable=wrong-import-position
import os
import glob
import time
import re

import ofx_online, quotes, site_cfg, scrubber
from control2 import *
from rlib1 import *
# pylint: enable=wrong-import-position

#startup
print('')
userdat = site_cfg.site_cfg()
log = create_logger('root', 'getdata.log')
if Debug:
    logging.basicConfig(level=logging.DEBUG)
    log.warning("**DEBUG Enabled: See Control2.py to disable.")
    log.debug('xfrdir = %s' % xfrdir)


def getSite(ofx: str):

    # find matching site entry for ofx
    # matches on FID or BANKID value found in ofx and in sites list

    #get fid value from ofx
    site = None
    p = re.compile(r'<FID>(.*?)[<\s]',re.IGNORECASE | re.DOTALL)
    r = p.search(ofx)
    fid = r.groups()[0] if r else 'undefined'
    p = re.compile(r'<BANKID>(.*?)[<\s]',re.IGNORECASE | re.DOTALL)
    r = p.search(ofx)
    bankid = r.groups()[0] if r else 'undefined'
    sites = userdat.sites
    if fid or bankid:
        for s in sites:
            if not site: site=sites[s]   #defaults to first site found, if matching fid/bankid not found
            thisFid    = FieldVal(sites[s], 'fid')
            thisBankid = FieldVal(sites[s], 'bankid')
            if thisFid == fid or thisBankid == bankid:
                site = sites[s]
                log.info('Matched import file to site *%s*' % s)
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
        dl_interval (int): Download interval in days.

    Returns:
        A 2-element tuple containing:
            - bool: Overall status of the download operations.
            - list: List of downloaded OFX files.
    """
    #get download interval, if promptInterval=Yes in sites.dat
    # should move into get ofx function
    dl_interval = userdat.defaultInterval
    if userdat.promptInterval:
        try:
            p = int2(input("Download interval (days) [" + str(dl_interval) + "]: "))
            if p>0:
                dl_interval = p
            else:
                raise ValueError("Download interval must be a positive integer.")
        except ValueError:
            log.info("Invalid entry. Using defaultInterval=%s", dl_interval)

    log.info("ownload interval= %s days", dl_interval)

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
                ofx_list.append([acct[0], acct[1], ofx_file])
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
    for f in glob.glob(importdir+'*.*'):

        # Get the parts of the filename
        fname = os.path.basename(f)             # full base filename.extension
        bext  = os.path.splitext(fname)[1]      # file extension

        # Read the file
        with open(f) as ifile:
            dat = ifile.read()

        # Only process if it looks like an ofx file
        if validOFX(dat) == '':
            log.info("Importing %s", fname)

            # Scrub file if it hasn't already been imported (and hence, scrubbed)
            if 'NEWFILEUID:PSIMPORT' not in dat[:200]:
                try:
                    site = getSite(dat)
                    scrubber.scrub(f, site)
                except re.error:
                    log.info("No site defined for %s in sites.dat: skipping scrub "
                             "routines",  fname)

            # Set NEWFILEUID:PSIMPORT to flag the file as having already been imported
            # Don't want to accidentally scrub twice
            with open(f, 'r', encoding='utf-8') as ifile:
                ofx = ifile.read()
            pattern = re.compile(r'NEWFILEUID:.*')
            ofx2 = pattern.sub('NEWFILEUID:PSIMPORT', ofx)
            if ofx2:
                with open(f, 'w') as ofile:
                    ofile.write(ofx2)

            account_id = get_acctid(ofx2)

            # Preserve original file type but save w/ ofx extension and move to xfrdir
            outname = xfrdir+fname + ('' if bext=='.ofx' else '.ofx')
            os.rename(f, outname)
            ofx_list.append(['import file', account_id, outname])
            log.info('%s saved to %s', fname, outname)

    return result, ofx_list


def send_files_to_money(ofx_list: list, quote_file_forced: str, interactive_flag: bool):
    """
    Sends all OFX files to Microsoft Money.

    Sends all OFX files in ofxList to Microsoft Money.  If the user has
    enabled the combineOFX option, all files are combined into a single file
    before sending.

    Args:
        ofx_list (list): List containting all OFX files to send to Money.
        quote_file_forced (str): Downloaded ForceQuotes OFX file.
        interactive_flag (bool): Flag indicating whether to prompt the user for
            interactive input.
    """

    # Combine OFX files if option set
    cfile = combineOfx(ofx_list) if (userdat.combineofx and len(ofx_list) > 1) else ""

    # Determine if user wants to send results to Money, and if so, whether they want
    # to verify each file before sending.
    if interactive_flag:
        userin = str(input_default(
            "\nSend Results to Money? (y/n/v=Verify)", 'y', str))[:1].upper()
    else:
        userin = 'Y' if userdat.sendToMoney else 'N'

    # Process user selection
    match userin:
        case 'N':
            # Don't send to Money
            log.info("Results not sent to Money (user selection or sites.dat setting).")

        case 'Y' | 'V':
            # Send to Money
            log.debug('User confirmed upload to Money.')

            # Send ForceQuotes file to Money if defined (NEEDS CLEANUP)
            if glob.glob(quote_file_forced):
                if Debug:
                    log.debug("Importing ForceQuotes statement: %s", quote_file_forced)
                run_file(quote_file_forced)  #force transactions for MoneyUK
                input(
                    "ForceQuote statement loaded.  Accept in Money and press <Enter> "
                    "to continue."
                )

            # Send individual file(s) or combined file to Money
            print("")
            log.info("Sending statement(s) to Money...")

            if cfile and userin != 'V':
                # Send combined file
                log.info("Importing combined OFX file: %s", cfile)
                run_file(cfile)

            else:
                # Send individual fils, prompting for verification if user selected 'V'
                for ofxfile in ofx_list:
                    # Upload each file one at a time, verify if requested
                    if (input_default(
                        f"Upload {ofxfile[0]} : {ofxfile[1]}? (y/n)", 'y', bool)
                        if userin == 'V' else True
                    ):

                        log.info("Importing %s", ofxfile[2])
                        run_file(ofxfile[2])

                    # Slight delay, to force load order in Money
                    time.sleep(0.5)

        case _:
            # Invalid selection
            log.info('Invalid selection.  Results not sent to Money.')


def input_default(prompt: str, default: object, type_cast: type=str) -> object:
    """
    Prompts the user for input with a default value.

    When type_cast is set to bool, accepts Y/N, Yes/No, True/False, T/F, 1/0
    (case-insensitive) as valid inputs.  In this case the default should be set to a user
    friendly value (e.g. 'Y' or 'N') rather than a boolean value.

    Args:
        prompt (str): The prompt message to display to the user.
        default (any): The default value to use if the user provides no input.
        type_cast (type, optional): A function to cast the input to a specific type.
            Defaults to str.

    Returns:
        The value entered by the user, cast to the specified type, or the default value
        if no input is provided.
    """

    # Loop until valid(ish) input is received
    user_input = ""
    while not user_input:
        try:
            # Get the user input
            user_input = input(f"{prompt} [{default}]: ")

            # Process the special case of boolean input
            if type_cast == bool:
                user_input = str(default) if not user_input else user_input
                if user_input.upper() in ['Y', 'YES', 'TRUE', 'T', '1']:
                    user_input = True
                elif user_input.upper() in ['N', 'NO', 'FALSE', 'F', '0']:
                    user_input = False
                else:
                    raise ValueError("Invalid boolean input")
                break

            # Process all other types
            else:
                user_input = default if not type_cast(user_input) else user_input

        except ValueError:
            # Only error we expect is a ValueError from an invalid type cast, so we can
            # catch that and prompt again
            print(f"Invalid entry. Please enter a value of type {type_cast.__name__}.")

    return user_input


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
    pwkey, getquotes, acct_array = get_cfg()
    #ToDo: userdat.fetchQuotes: = getquotes

    if len(acct_array) > 0 and pwkey != '':
        #if accounts are encrypted... decrypt them
        pwkey=decrypt_pw(pwkey)
        acct_array = acctDecrypt(acct_array, pwkey)

    #delete old data files
    ofxfiles = xfrdir+'*.ofx'
    if glob.glob(ofxfiles):
        os.system("del "+ofxfiles)

    # Determine if running in interactive mode.  If so, prompt user at each step.  If
    # not, use settings defined in sites.dat
    interactive_flag = bool(
        input_default("Run in interactive mode? (y/n)", 'n', bool)
        if userdat.promptStart else False
    )

    # Create process Queue in the right order
    work_queue = ['Accts', 'importFiles']
    if userdat.savetickersfirst:
        work_queue.insert(0,'Quotes')
    else:
        work_queue.append('Quotes')

    # Initialize variables to track status and results across operations
    status = True   # Overall status across all operations
                    # True if all operations are succesful
    ofx_list = []   # List of all OFX files to send to Money.
                    # Each item is a list: [SiteName, Account#, OFX filename]
    quote_file, quote_file_forced, html_quote_file = "", "", ""

    # Get and process files in the order defined in Queue
    for q_entry in work_queue:

        # Get OFX files for Direct Connect accounts
        if q_entry == 'Accts':
            if (input_default(
                "\nDownload statements for all online accounts? (y/n)", 'n', bool)
                if interactive_flag else userdat.fetchRemote
            ):

                log.info("\n--- Downloading statements for all accounts ---")
                accts_status, new_list = get_directconnect_ofx_files(acct_array)
                ofx_list.extend(new_list)
                status = status and accts_status

        # Get OFX files from import folder
        if q_entry == 'importFiles':
            if (input_default(
                "\nProcess files in import folder? (y/n)", 'n', bool)
                if interactive_flag else userdat.fetchImport
            ):

                print("")
                log.info("--- Processing OFX files in import folder ---")
                import_status, new_list = get_import_ofx_files()
                ofx_list.extend(new_list)
                status = status and import_status

        # Get stock/fund quotes
        if q_entry == 'Quotes':
            if (input_default(
                "\nDownload stock/fund quotes? (y/n)", 'n', bool)
                if interactive_flag else userdat.fetchQuotes
            ):

                print("")
                log.info("--- Downloading stock/fund quotes ---")
                quote_status, quote_file, quote_file_forced, html_quote_file = (
                    quotes.getQuotes())
                if quote_status:
                    new_list = ['Stock/Fund Quotes','',quote_file]
                    ofx_list.append(new_list)
                status = status and quote_status

    print("")
    log.info('--- OFX file fetching completed. ---')

    # Process all OFX files and send to Money.
    if len(ofx_list) > 0:
        # Send to money, the interactive prompts get dealt with inside the function
        send_files_to_money(ofx_list, quote_file_forced, interactive_flag)

    else:
        log_message = (
            "No OFX files were fetched or created. Verify network connection "
            "and/or input folders.")
        log.warning(log_message)

        # Set status to False to prevent auto-close of command window
        status = False

    # Display quotes.htm if downloaded and user wants to see it.
    htm_filename_root = os.path.basename(html_quote_file) if html_quote_file else ""

    if html_quote_file and (input_default(
        f"\nOpen <{htm_filename_root}> in the default browser? (y/n)", 'n', bool)
        if (interactive_flag or userdat.askquotehtm) else userdat.showquotehtm
    ):

        log.debug("Full quotes file name: %s", html_quote_file)
        log.info("Opening %s in browser.", htm_filename_root)
        os.startfile(html_quote_file)  #don't wait for browser close

    # Keep window open if needed, complete execution
    if not status:
        log_message = (
            "One or more errors were detected during the processing of OFX files. "
            "Review display and log to identify the problem.")
        log.warning(log_message)

    if Debug or interactive_flag or userdat.promptEnd or not status:
        input("\nPress <Enter> to close the window...")

    print("")
    log.info(
        '-----------------------------------------------------------------------------')


if __name__=="__main__":
    main()
