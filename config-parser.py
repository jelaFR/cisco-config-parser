# Import libraries
import os
from xlsxwriter import Workbook as Wb
from ciscoconfparse import CiscoConfParse
from tqdm import tqdm

def convert_list_to_xlswriter_headers(list_name):
    """This function converts a simple list in XLSXwriter headers
       Header format is [{'header': 'name1'}, {'header': 'name2'}, {'header': 'name3}]

    Args:
        list_name (list): List to be converted in headers
    """
    headers = list()
    for item in list_name:
        item_transformed = {'header': f'{item}'}
        headers.append(item_transformed)

    return headers


def search_file_in_directory(in_directory,filter=""):
    """
    This function search all files inside a directory with a specific item in filter
    :param in_directory: Root directory where files are located
    :param filter: (optionnal) Return only files containing this value
    :return: List containing files with full path
    """
    file_list = list()
    for root,subdir,files in os.walk(in_directory):
        for file in files:
            if filter.lower() in file.lower():
                file_list.append(os.path.join(root,file))
    return file_list


def sh_run_to_dict(filename):
    """
    This script parse "show running-config" file looking for parameters
    This is working on Cisco Catalyst and Nexus switches
    :param filename: Filename where output of "show running-config" is located
    :return: A dictionary containing list of switches and ifaces parameters
    """

    ##
    # Try to automatically determine the platform
    ##
    try:
        parse = CiscoConfParse (filename, factory=True, syntax='nxos')
        if len(parse.re_search_children("^interface \S+Ethernet")) != 0:
            parse = CiscoConfParse (filename, factory=True, syntax='ios')
    except ValueError:
        parse = CiscoConfParse (filename, factory=True, syntax='ios')


    ##
    # Get infos from the top of the config file
    ##

    # Hostname (used to check if this is a valid Cisco configuration file)
    hostname = parse.find_lines("hostname")
    if len(hostname) == 0: # it seems that it is not an Cisco config file
        return
    hostname = hostname[0].strip("hostname ")
    HOSTNAME_LIST.append(hostname.lower())
    SWITCHES_PARAMS[hostname] = dict()

    # Vlan List
    vlan_list = parse.re_search_children("^vlan \d+$")
    for vlan_object in vlan_list:
        vlan_id = str(vlan_object.text.strip("vlan "))
        vlan_name = vlan_object.re_search_children("name")
        if len(vlan_name) != 0:
            vlan_name = vlan_object.re_search_children("name")[0].text.strip(" name")
        else:
            vlan_name = ""
        # Create a table for later use in Excel
        if vlan_id not in GLOBAL_VLAN:
            GLOBAL_VLAN[vlan_id] = dict()
            GLOBAL_VLAN[vlan_id]["vlan_id"] = vlan_id
            GLOBAL_VLAN[vlan_id]["vlan_name"] = vlan_name
            GLOBAL_VLAN[vlan_id]["switches"] = list()
        if hostname not in GLOBAL_VLAN[vlan_id]["switches"]:
            GLOBAL_VLAN[vlan_id]["switches"].append(hostname.lower())

    # Voice Vlan
    global_voice_vlan = parse.find_objects_w_parents("^network-policy profile","voice vlan")
    if global_voice_vlan:
        global_voice_vlan = global_voice_vlan[0].text.split()[2]
    
    # Iface list
    all_ifaces = parse.find_objects ("^interface")
    for iface_param in all_ifaces:
        # Populate ifaces params in dictionnary for each iface
        if iface_param.is_ethernet_intf or iface_param.re_search_children ("channel-group"):

            # Get interface name and initiate dictionary
            iface_name = iface_param.text.strip("interface ")
            SWITCHES_PARAMS[hostname][iface_name] = dict()
            SWITCHES_PARAMS[hostname][iface_name]["hostname"] = hostname
            SWITCHES_PARAMS[hostname][iface_name]["iface_name"] = iface_name

            # Get interface mode (access,trunk or dynamic)
            iface_mode = iface_param.re_search_children ("switchport mode")
            if iface_mode:
                iface_mode = iface_mode[0].text.strip(" switchport mode ")
            else:
                iface_mode = 'dynamic'  # VAR iface_mode (mode not configured)
            SWITCHES_PARAMS[hostname][iface_name]["mode"] = iface_mode

            # Get interface description
            iface_desc = iface_param.re_search_children ("description")
            if iface_desc:
                iface_desc = iface_desc[0].text.strip("  description ")  # VAR iface_desc
            else:
                iface_desc = ''  # VAR iface_desc (iface without description)
            SWITCHES_PARAMS[hostname][iface_name]["description"] = iface_desc

            # Check if interface has authentication configured
            iface_auth = iface_param.re_search_children ("authentication port-control auto")
            if iface_auth:
                SWITCHES_PARAMS[hostname][iface_name]["authentication"] = 'yes'
            else:
                SWITCHES_PARAMS[hostname][iface_name]["authentication"] = ''

            # Check if iface is part of port-channel
            if iface_param.re_search_children ("channel-group"):
                iface_etherchannel = iface_param.portchannel_number
                SWITCHES_PARAMS[hostname][iface_name]["etherchannel_id"] = iface_etherchannel
            else:
                SWITCHES_PARAMS[hostname][iface_name]["etherchannel_id"] = ""

            # Collect access vlan configured
            iface_access_vlan = iface_param.re_search_children ("switchport access vlan")
            if iface_access_vlan:
                iface_access_vlan = iface_access_vlan[0].text.strip("  switchport access vlan ")
            else:
                iface_access_vlan = ''
            SWITCHES_PARAMS[hostname][iface_name]["access_vlan"] = iface_access_vlan

            # Collect voice vlan configured (if any)
            iface_voice_vlan = iface_param.re_search_children ("switchport voice vlan")
            if iface_voice_vlan:
                iface_voice_vlan = iface_voice_vlan[0].text.strip("  switchport voice vlan ")
            elif global_voice_vlan:
                iface_voice_vlan = global_voice_vlan
            else:
                iface_voice_vlan = ''
            SWITCHES_PARAMS[hostname][iface_name]["voice_vlan"] = iface_voice_vlan

            # Collect allowed vlans trunk informations
            iface_trunk_vlan = iface_param.re_search_children ("switchport trunk allowed vlan")
            if iface_trunk_vlan:
                iface_trunk_vlan = iface_trunk_vlan[0].text.strip("  switchport trunk allowed vlan ")
                if iface_param.re_search_children ("switchport trunk allowed vlan add"):
                    len_iface_trunk_add = len(iface_param.re_search_children ("switchport trunk allowed vlan add"))
                    for index in range(len_iface_trunk_add):
                        iface_trunk_vlan_add = iface_param.re_search_children ("switchport trunk allowed vlan add")[index]\
                            .text.strip("  switchport trunk allowed vlan add ")
                        iface_trunk_vlan = iface_trunk_vlan+","+iface_trunk_vlan_add
            else:
                iface_trunk_vlan = ''
            iface_trunk_vlan = str (iface_trunk_vlan)
            SWITCHES_PARAMS[hostname][iface_name]["trunk_vlan"] = iface_trunk_vlan

            # Collect native vlan trunk informations
            iface_trunk_native = iface_param.re_search_children ("switchport trunk native vlan")
            if iface_trunk_native:
                iface_trunk_native = iface_trunk_native[0].text.strip("  switchport trunk native vlan ")
            else:
                iface_trunk_native = ''
            SWITCHES_PARAMS[hostname][iface_name]["trunk_native"] = iface_trunk_native

            # Collect speed informations
            iface_speed = iface_param.re_search_children ("speed")
            if iface_speed:
                iface_speed = iface_speed[0].text.strip("  speed ")
            else:
                iface_speed = 'auto'
            SWITCHES_PARAMS[hostname][iface_name]["iface_speed"] = iface_speed

            # Collect duplex informations
            iface_duplex = iface_param.re_search_children ("duplex")
            if iface_duplex:
                iface_duplex = iface_duplex[0].text.strip("  duplex ")
            else:
                iface_duplex = 'auto'
            SWITCHES_PARAMS[hostname][iface_name]["iface_duplex"] = iface_duplex
    
    #SVI List
    all_routed_ifaces = parse.find_objects_w_child(parentspec=r"^interface Vlan", childspec=r"ip address")
    for iface_param in all_routed_ifaces:
        vlan_id = iface_param.text.strip("interface Vlan")
        ip_address_cidr = f"{iface_param.ip_addr}/{iface_param.ipv4_masklength}"
        GLOBAL_SVI[(hostname.lower(),vlan_id)] = ip_address_cidr


def dict_to_xlsx(out_file):
    """
    This function parse a dictionary and create excel files based on IT.
    The first level of parameters will become filenames, the second tabs in excel, the third lines
    :param out_file: excel file
    """
    # Check if file exists and remove previous one
    if os.path.isfile(out_file):
        try:
            os.remove(out_file)
        except PermissionError:
            while True:
                try:
                    os.remove(out_file)
                    break
                except PermissionError:
                    print(f"[ERROR] {os.path.basename(out_file)} is actually opened !!")
                    input(f"Please close it and press enter.")

    # Create an excel workbook
    workbook = Wb(out_file)
    worksheet_iface = workbook.add_worksheet("ifaces")
    worksheet_iface.freeze_panes(1, 1)
    worksheet_vlans = workbook.add_worksheet("vlan_list")
    worksheet_vlans.freeze_panes(1, 1)

    # Create specific format
    cell_green = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'green'})
    cell_red = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'red'})

    # Get SWITCH table params
    iface_total = 0
    iface_global_params = list()
    for switch_name in SWITCHES_PARAMS:
        for iface_name in SWITCHES_PARAMS[switch_name]:
            iface_total += 1
            if iface_total == 1:
                ifaces_headers = list(SWITCHES_PARAMS[switch_name][iface_name].keys())
            iface_global_params.append(list(SWITCHES_PARAMS[switch_name][iface_name].values()))
    ifaces_headers = convert_list_to_xlswriter_headers(ifaces_headers)

    # Get VLAN table params
    HOSTNAME_LIST.sort() # Order hosname alphetically
    vlan_headers = list(GLOBAL_VLAN.get("1","").keys())
    vlan_headers.pop()
    vlan_headers = vlan_headers + HOSTNAME_LIST
    vlan_headers = convert_list_to_xlswriter_headers(vlan_headers)
    vlan_global_params = list()
    """ Parse all vlans in order to create an output list
        containing VLAN_ID, VLAN_NAME, SW1, SW2, SW3 with values "Yes"
        if vlan is created on this switch or "No" if it's not
    """
    for vlan_id in GLOBAL_VLAN:
        vlan_params_list = list()
        for vlan_params in GLOBAL_VLAN[vlan_id]:
            if vlan_params == "vlan_id":
                vlan_params_list.append(GLOBAL_VLAN[vlan_id]["vlan_id"])
                for sw_hostname in HOSTNAME_LIST:
                    # Vlan exist on this switch
                    if sw_hostname in GLOBAL_VLAN[vlan_id]["switches"]:
                        vlan_params_list.append(GLOBAL_SVI.get((sw_hostname, vlan_id),"Yes"))
                    else: # Vlan does not exist
                        vlan_params_list.append("No")
            elif vlan_params == "vlan_name":
                vlan_params_list.insert(1, GLOBAL_VLAN[vlan_id]["vlan_name"])
        vlan_global_params.append(vlan_params_list)

    # Add tables inside WorkSheets
    worksheet_iface.add_table(0, 0, iface_total, len(ifaces_headers) -1, \
        {'name': 'iface_list', 'data': iface_global_params, 'columns': ifaces_headers})
    worksheet_vlans.add_table(0, 0, len(GLOBAL_VLAN), len(vlan_headers) -1, \
        {'name': 'vlan_list', 'first_column': True, 'data': vlan_global_params, 'columns': vlan_headers})


    # Add conditional formatting inside worksheet Vlans
    worksheet_vlans.conditional_format(0, 0, iface_total, len(ifaces_headers) -1, {'type': 'text', 'criteria': 'containing',\
        'value': 'Yes', 'format': cell_green})
    worksheet_vlans.conditional_format(0, 0, iface_total, len(ifaces_headers) -1, {'type': 'text', 'criteria': 'containing',\
        'value': 'No', 'format': cell_red})


    # Save and exit
    workbook.close()


if __name__ == "__main__":
    try:
        welcome_msg = "Create Excel file from CISCO cfg with one line per interface"
        welcome_len = len(welcome_msg) + 4
        print("*"*welcome_len)
        print("*",welcome_msg,"*")
        print("*"*welcome_len)
        # Vars
        while True:
            in_dir = input("Please type path where files are located:")
            if os.path.isdir(in_dir):
                break
        out_file = os.path.join(in_dir,"output.xlsx")

        # Define global vars
        SWITCHES_PARAMS = dict()
        GLOBAL_VLAN = dict()
        GLOBAL_SVI = dict()
        HOSTNAME_LIST = list()

        # Main
        file_list = search_file_in_directory(in_dir,filter="config")
        for filename in tqdm(file_list,desc="Analysing cisco Cfg file"):
            sh_run_to_dict (filename)
        dict_to_xlsx(out_file)

        end_msg = "Script is now finished, out file is named: "+os.path.basename(out_file)
        end_len = len(end_msg) + 4
        print("*"*end_len)
        print("*",end_msg,"*")
        print("*"*end_len)

    # Exception handling
    except KeyboardInterrupt:
        os.system("cls")
        print("\n[KeyInterrupt] Exiting as requested")
        exit(0)
