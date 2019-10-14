# cisco-config-parser
This script parse Cisco configuration files and output to an excel file

For now, this script is tested with Catalyst platforms (2960,3750,4500,6500) and on Nexus platform (Nexus3000)
This should work on other Cisco models as this is based on the CiscoConfParse library.

The excel file will return these parameters (one worksheet by hostname):
Interface	: The name of the interface
mode : currently configured mode (or DTP)
description	: the description on the interface
authentication	: is the interface authenticated
etherchannel_id	: ID of port-channel
access_vlan	: access vlan configured on the port
voice_vlan	: voice vlan configured on the port or globally configured using network-profile
trunk_vlan	: trunk vlan configured
trunk_native	: native vlan on trunk
previous_config : a list with the full previous configuration
