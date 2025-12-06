# control2.py
# http://sites.google.com/site/pocketsense/
# contains some common configuration data and modules for the ofx pkg
# Initial version: rlc: Feb-2010
#
#04-Jan-2010*rlc
#   - Added DefaultAppID and DefaultAppVer

# 27Jul2013: rlc
#   - Added locale support
# 03Sep2014: rlc
#   - xfrdir is now platform independent
# 06Mar2017: rlc
#   - Set DefaultAppVer = 2400 (Quicken 2015)
#   - Moved utility functions to rlib1 module
# 20Jun2023
#   - Add logging
# 03Dec2023 cgn
#   - Update to python3
#------------------------------------------------------------------------------------

#---MODULES---
import os

Debug = False
#Debug = True           #debug mode = enable only when testing and delete log files after using.

SendtoMoney = False     #enable sending downloaded files to MS Money

#logging
logFileEnable = True
logFileLimit  = 10       #logfile size limit (MB)

AboutTitle    = 'PocketSense OFX Download Python Scripts for Python 3'
AboutVersion  = '3.0.beta'
AboutSource   = 'http://sites.google.com/site/pocketsense'
AboutName     = 'Robert & Carlos'

#xfrdir = temp directory for statement downloads.  Platform independent
xfrdir    = os.path.join(os.path.curdir,"xfr") + os.sep
importdir = os.path.join(os.path.curdir,"import") + os.sep
cfgFile  = 'ofx_config.cfg'    #user account settings (can be encrypted)

DefaultAppID  = 'QWIN'
DefaultAppVer = '2700'





