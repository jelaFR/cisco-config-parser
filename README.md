# cisco-config-parser
This script parse Cisco configuration files and output to an excel file
This in initially posted by Jean-Eudes LAYRISSE on 01/12/2018

Change directory variable to the path where configuration files are located
More, you can change the out_file variable to specify the path to the output file.

For now, this script is tested with Catalyst platforms (2960,3750,4500,6500) and on Nexus platform (Nexus3000)
This should work on other Cisco models as this is based on the CiscoConfParse library.

In order to work, you'll need to install xlsxwriter and ciscoconfparse
