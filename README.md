mac-arp-diff.py

This python script is designed to be run from within a SecureCRT window when
connected to a Cisco distribution switch. It obtains the MAC address table and
the ARP table, creates a spreadsheet named for the switch, writes the pertinent
information to two different tabs and creates a conditional formatting that
highlights all of the addresses in the MAC table that do not have a corresponding
address in the ARP table. This is helpful to locate devices that have been configured
for a different vlan than the one they are attached to. It also helpfully references
a list of IEEE OUIs and will tell you the manufacturer of each MAC.

In SecureCRT, you can create a button to launch the script.
