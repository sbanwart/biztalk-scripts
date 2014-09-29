biztalk-scripts
===============

A PowerShell script for creating and configuring a set of hosts, host instances and send/receive handlers. This is a pure PowerShell implementation with no dependencies outside of PowerShell and BizTalk.

Usage
-----

Configure the options at the top of the file for your environment and then run the script. This script works best if this script is run prior to configuring any BizTalk ports.

Prerequisites
-------------

* PowerShell 2.0+

Known Limitations
-----------------

* Currently the script has only been tested with creating in-process hosts.
* Error handling is basically non-existent

License
-------

Apache 2.0

