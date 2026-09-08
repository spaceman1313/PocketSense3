# *PocketSense3*
Python3 implementation of *PocketSense* OFX handling scripts

This is a Python3 implementation of the PocketSense scripts developed by Robert and
found here:
https://sites.google.com/site/pocketsense/home/msmoneyfixp1

In addition to having ported the scripts to Python 3, the `Getdata.py` main module has
been substantially redone. `Getdata.py` is now PEP8 compliant and has numerous
additional comments allowing the user to better understand what is going on.  In
addition the main workflow now offers finer control (via interactive user input or
`sites.dat` settings) to allow the user to control whether to:
- Download OFX statements from direct connect servers (many of us no longer have any
  financial institutions that support these)
- Import OFX statements that have been downloaded via the web or created from CSV
  files
- Download quotes from Yahoo!
- Process OFX files through the scrubbers
- Automatically send the OFX file(s) to Microsoft Money

My long term intent is to continue refactor and modularize some of the code to make it
easier to implement additional changes as well as use modern debugging tools.  In the
development pipeline is the incorporation of other available CSV to OFX scripts.  This
will allow users to continue using OFX functionality in Microsoft Money and other legacy
accounting systems even as more banks discontinue the use of OFX.

Please be aware that as of March 2026 I no longer have access to any banks with
DirectConnection and thus cannot test this functionality any more. As such it is
impossible for me to test many of the different permutations users may have.  I rely on
you to provide me with feedback and bugs.  Please  enter a GitHub issue (preferred) or
email me if you encounter a problem.

I can be reached at pocketsense3 at the usual g email system.

### Requirements:
I believe the scripts should be compatible with Python 3.11 and higher, however I have
only tested them with Python 3.14 (miniconda distribution). __You will also need to
install the `requests` package in your Python installation.__

### Installation:
Follow the instructions in the original PocketSense website, but use Python 3.11 or
higher:
https://sites.google.com/site/pocketsense/home/msmoneyfixp1/p2

### Transferring from PocketSense for Python 2:
You should be able to copy your `sites.dat`, `connect.key`, and `ofx_config.cfg` files
and use the new scripts without any issues. I strongly recommend you unencrypt your
`ofx_config.cfg` file by running the *PocketSense2* setup utility before encrypting it
again using the *PocketSense3* setup utility.
> [!TIP]
> Please retain copies of your original configuration files. *PocketSense3* will update
> the files with a format that may not be backwards compatible.

You should also take a look at the new settings in the `sites.template` file
(`fetchRemote`, `fetchImport`, `fetchQuotes`, `scrubOFX`, and `sendToMoney`) and copy
those to your own `sites.dat` file. These settings enable finer control over what
operations PocketSense3 carries out when running `getdata.py`.

`Setup.py` can no longer be used to control the setting that enables/disables quote
downloads.  That should now be done via the `fetchQuotes` setting in `sites.dat`.

__One final note.__  `control2.py` has a new debug setting called `BlockSendToMoney`.  As
its name implies, this setting will block sending the OFX file(s) to Microsoft Money at
the very last step in the process, thus overriding the `sites.dat` setting
`sendToMoney`. This setting is purposely set to True due to the beta nature of these
scripts.  When you feel comfortable with this new version, set that setting to False to
allow *PocketSense3* to work as intended.

> [!IMPORTANT]
> *PocketSense* (both 2 and 3) use DES encryption, if desired, to locally store
> passwords. This is probably good enough for storing passwords on your own computer,
> but generally the DES algorithm is no longer considered the best for encrypting
> sensitive data. The encryption algorithm used to connect to a financial's institution
> OFX server is the same HTTPS as on your web browser, so that is not a concern.