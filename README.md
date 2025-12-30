#PocketSense3
Python3 implementation of PocketSense OFX handling scripts

This is a Python3 implementation of the PocketSense scripts developed by Robert and found
here:
https://sites.google.com/site/pocketsense/home/msmoneyfixp1

The initial 3.0.beta release is a pure translation of the original scripts.  My long term
intent is to refactor and modularize some of the code to make it easier to implement
additional changes.  In the development pipeline is the incorporation of other available
CSV to OFX scripts.  This will allow users to continue using OFX functionality in
Microsoft Money and other legacy accounting systems even as more banks discontinue the
use of OFX.

Please be aware that as of December 2025, the only two banks to which I have access with
DirectConnection are Fidelity and NetBenefits (which are really the same bank).  As such
it is impossible for me to test many of the different permutations users may have.  I
rely on you to provide me with feedback and bugs.  Please email me or enter a GitHub
issue if you encounter a problem.

I can be reached at pocketsense3 at the usual g email system.

Requirements:
I believe the scripts should be compatible with Python 3.10 and higher, however I have
only tested them with Python 3.14 (miniconda distribution).  You will also need to
install the `requests` package in your Python installation.

Installation:
Follow the instructions in the original PocketSense website, but use Python 3.10 or
higher:
https://sites.google.com/site/pocketsense/home/msmoneyfixp1/p2

Transferring from PocketSense for Python 2: You should be able to copy your `sites.dat`,
`connect.key`, and `ofx_config.cfg` files and use the new scripts without any issues.  I
have not yet tested with an encrypted configuration file. Please retain copies of your
original configuration files.  PocketSense3 will update the files with a format that is
not backwards compatible.