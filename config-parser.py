import os
import xlsxwriter
from ciscoconfparse import CiscoConfParse
from tqdm import tqdm

def search_file_in_directory(directory,filter=""):
    """
    This function search all files inside a directory with a specific item in filter
    :param directory: Root directory where files are located
    :param filter: (optionnal) Return only files containing this value
    :return: List containing files with full path
    """
    file_list = list()
    for root,subdir,files in os.walk(directory):
        for file in files:
            if filter.lower() in file.lower():
                file_list.append(os.path.join(root,file))
    return file_list


def sh_run_to_dict(filename, output_dictionary, platform="ios"):
    """
    This script parse "show running-config" file looking for parameters
    This is working on Cisco Catalyst and Nexus switches
    :param output_dictionary: Dictionary name where parameters are stored
    :param filename: Filename where output of "show running-config" is located
    :param platform: "ios" or "Catalyst" for regular Catalyst IOS; "nxos" or "nexus" for Cisco Nexus platform
    :return: A dictionary containing list of switches and ifaces parameters
    """
    #Check platform type if specified
    if (platform.lower() == "catalyst") or (platform.lower() == "ios"):
        parse = CiscoConfParse (filename, factory=True, syntax='ios')
    elif (platform.lower() == "nexus") or (platform.lower() == "nxos"):
        parse = CiscoConfParse (filename, factory=True, syntax='nxos')
    else:
        print("Error: please provide a valid platform name")
        return
    hostname = parse.find_lines ("hostname")
    if len(hostname) == 0:
        # it seemds that it is not an Cisco config file
        return
    hostname = hostname[0].replace ('hostname ', '')
    output_dictionary[hostname] = {}
    #Get global voice vlan informations
    global_voice_vlan = parse.find_objects_w_parents("^network-policy profile","voice vlan")
    if global_voice_vlan:
        global_voice_vlan = global_voice_vlan[0].text.split()[2]
    # Get interfaces list and parse it
    all_ifaces = parse.find_objects ("^interface")

    for iface_param in all_ifaces:
        # Populate ifaces params in dictionnary for each iface
        if iface_param.is_ethernet_intf or iface_param.re_search_children ("channel-group"):

            # Get interface name and initiate dictionary
            iface_name = iface_param.text.replace ("interface ", "")
            output_dictionary[hostname][iface_name] = {}

            # Get interface mode (access,trunk or dynamic)
            iface_mode = iface_param.re_search_children ("switchport mode")
            if iface_mode:
                iface_mode = iface_mode[0].text.replace (" switchport mode ", "")
            else:
                iface_mode = 'dynamic'  # VAR iface_mode (mode not configured)
            output_dictionary[hostname][iface_name]["mode"] = iface_mode

            # Get interface description
            iface_desc = iface_param.re_search_children ("description")
            if iface_desc:
                iface_desc = iface_desc[0].text.replace ("  description ", "")  # VAR iface_desc
            else:
                iface_desc = ''  # VAR iface_desc (iface without description)
            output_dictionary[hostname][iface_name]["description"] = iface_desc

            # Check if interface has authentication configured
            iface_auth = iface_param.re_search_children ("authentication port-control auto")
            if iface_auth:
                output_dictionary[hostname][iface_name]["authentication"] = 'yes'
            else:
                output_dictionary[hostname][iface_name]["authentication"] = ''

            # Check if iface is part of port-channel
            if iface_param.re_search_children ("channel-group"):
                iface_etherchannel = iface_param.portchannel_number
                output_dictionary[hostname][iface_name]["etherchannel_id"] = iface_etherchannel
            else:
                output_dictionary[hostname][iface_name]["etherchannel_id"] = ""

            # Collect access vlan configured
            iface_access_vlan = iface_param.re_search_children ("switchport access vlan")
            if iface_access_vlan:
                iface_access_vlan = iface_access_vlan[0].text.replace ("  switchport access vlan ", "")
            else:
                iface_access_vlan = ''
            output_dictionary[hostname][iface_name]["access_vlan"] = iface_access_vlan

            # Collect voice vlan configured (if any)
            iface_voice_vlan = iface_param.re_search_children ("switchport voice vlan")
            if iface_voice_vlan:
                iface_voice_vlan = iface_voice_vlan[0].text.replace ("  switchport voice vlan ", "")
            elif global_voice_vlan:
                iface_voice_vlan = global_voice_vlan
            else:
                iface_voice_vlan = ''
            output_dictionary[hostname][iface_name]["voice_vlan"] = iface_voice_vlan

            # Collect allowed vlans trunk informations
            iface_trunk_vlan = iface_param.re_search_children ("switchport trunk allowed vlan")
            if iface_trunk_vlan:
                iface_trunk_vlan = iface_trunk_vlan[0].text.replace ("  switchport trunk allowed vlan ", "")
                if iface_param.re_search_children ("switchport trunk allowed vlan add"):
                    len_iface_trunk_add = len(iface_param.re_search_children ("switchport trunk allowed vlan add"))
                    for index in range(len_iface_trunk_add):
                        iface_trunk_vlan_add = iface_param.re_search_children ("switchport trunk allowed vlan add")[index]\
                            .text.replace("  switchport trunk allowed vlan add ","")
                        iface_trunk_vlan = iface_trunk_vlan+","+iface_trunk_vlan_add
            else:
                iface_trunk_vlan = ''
            iface_trunk_vlan = str (iface_trunk_vlan)
            output_dictionary[hostname][iface_name]["trunk_vlan"] = iface_trunk_vlan

            # Collect native vlan trunk informations
            iface_trunk_native = iface_param.re_search_children ("switchport trunk native vlan")
            if iface_trunk_native:
                iface_trunk_native = iface_trunk_native[0].text.replace ('  switchport trunk native vlan ', '')
            else:
                iface_trunk_native = ''
            output_dictionary[hostname][iface_name]["trunk_native"] = iface_trunk_native

            # Collect speed informations
            iface_speed = iface_param.re_search_children ("speed")
            if iface_speed:
                iface_speed = iface_speed[0].text.replace ('  speed ', '')
            else:
                iface_speed = 'auto'
            output_dictionary[hostname][iface_name]["iface_speed"] = iface_speed

            # Collect duplex informations
            iface_duplex = iface_param.re_search_children ("duplex")
            if iface_duplex:
                iface_duplex = iface_duplex[0].text.replace ('  duplex ', '')
            else:
                iface_duplex = 'auto'
            output_dictionary[hostname][iface_name]["iface_duplex"] = iface_duplex

            # Collect full previous configuration for rollback
            #previous_config = iface_param.re_search_children ('')
            #previous_config_list = []
            #for param in previous_config:
            #    previous_config_list.append (param.text)
            #output_dictionary[hostname][iface_name]["previous_config"] = previous_config_list

    # Get vlan list and
    return


def dict_to_xlsx(dictionary,out_file):
    """
    This function parse a dictionary and create excel files based on IT.
    The first level of parameters will become filenames, the second tabs in excel, the third lines
    :param dictionary: the name of the dictionary
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
    workbook = xlsxwriter.Workbook(out_file)
    worksheet = workbook.add_worksheet("Analyse")
    # Create excel worksheet with keys
    row = 1
    for switch_name in dictionary:
        for switch_iface in dictionary[switch_name]:
            col = 2
            for iface_param_name in dictionary[switch_name][switch_iface]:
                iface_param_value = dictionary[switch_name][switch_iface][iface_param_name]
                if row == 1: # We are writing headers
                    if col == 2: # We are writing hostname + iface headers
                        worksheet.write(row-1, col-2, "Hostname")
                        worksheet.write(row-1, col-1, "Iface")
                        worksheet.set_column(row-1, col-2, 30)
                        worksheet.set_column(row-1, col-1, 30)
                    worksheet.write(row-1,col,iface_param_name)
                    worksheet.set_column(row-1, col, 30)
                # Once header are written, proceed normal behavior
                if col == 2:
                    worksheet.write(row, col-2, switch_name)
                    worksheet.write(row, col-1, switch_iface)       
                worksheet.write(row, col, str(iface_param_value))
                col += 1
            row += 1
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
            directory = input("Please type path where files are located:")
            if os.path.isdir(directory):
                break
        switch_params = dict()
        out_file = os.path.join(directory,"output.xlsx")

        # Main
        file_list = search_file_in_directory(directory,filter="config")
        for file in tqdm(file_list,desc="Analysing cisco Cfg file"):
            sh_run_to_dict (file,switch_params,"nxos")
        dict_to_xlsx(switch_params,out_file)

        end_msg = "Script is now finished, out file is named: "+os.path.basename(out_file)
        end_len = len(end_msg) + 4
        print("*"*end_len)
        print("*",end_msg,"*")
        print("*"*end_len)
    except KeyboardInterrupt:
        os.system("cls")
        print("\n[KeyInterrupt] Exiting as requested")
        exit(0)
