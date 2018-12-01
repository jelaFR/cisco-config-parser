def search_file_in_directory(directory,filter=""):
    """
    This function search all files inside a directory with a specific item in filter
    :param directory: Root directory where files are located
    :param filter: (optionnal) Return only files containing this value
    :return: List containing files with full path
    """
    import os
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
    from ciscoconfparse import CiscoConfParse
    #Check platform type if specified
    if (platform.lower() == "catalyst") or (platform.lower() == "ios"):
        parse = CiscoConfParse (filename, factory=True, syntax='ios')
    elif (platform.lower() == "nexus") or (platform.lower() == "nxos"):
        parse = CiscoConfParse (filename, factory=True, syntax='nxos')
    else:
        print("Error: please provide a valid platform name")
        return
    hostname = parse.find_lines ("hostname")
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
        if iface_param.is_ethernet_intf or iface_param.is_portchannel:

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
                iface_desc = iface_desc[0].text.replace (" description ", "")  # VAR iface_desc
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
            if iface_param.in_portchannel:
                iface_etherchannel = iface_param.portchannel_number
                output_dictionary[hostname][iface_name]["etherchannel_id"] = iface_etherchannel
            else:
                output_dictionary[hostname][iface_name]["etherchannel_id"] = ""

            # Collect access vlan configured
            iface_access_vlan = iface_param.re_search_children ("switchport access vlan")
            if iface_access_vlan:
                iface_access_vlan = iface_access_vlan[0].text.replace (" switchport access vlan ", "")
            else:
                iface_access_vlan = ''
            output_dictionary[hostname][iface_name]["access_vlan"] = iface_access_vlan

            # Collect voice vlan configured (if any)
            iface_voice_vlan = iface_param.re_search_children ("switchport voice vlan")
            if iface_voice_vlan:
                iface_voice_vlan = iface_voice_vlan[0].text.replace (" switchport voice vlan ", "")
            elif global_voice_vlan:
                iface_voice_vlan = global_voice_vlan
            else:
                iface_voice_vlan = ''
            output_dictionary[hostname][iface_name]["voice_vlan"] = iface_voice_vlan

            # Collect allowed vlans trunk informations
            iface_trunk_vlan = iface_param.re_search_children ("switchport trunk allowed vlan")
            if iface_trunk_vlan:
                iface_trunk_vlan = iface_trunk_vlan[0].text.replace (" switchport trunk allowed vlan ", "")
                if iface_param.re_search_children ("switchport trunk allowed vlan add"):
                    len_iface_trunk_add = len(iface_param.re_search_children ("switchport trunk allowed vlan add"))
                    for index in range(len_iface_trunk_add):
                        iface_trunk_vlan_add = iface_param.re_search_children ("switchport trunk allowed vlan add")[index]\
                            .text.replace(" switchport trunk allowed vlan add ","")
                        iface_trunk_vlan = iface_trunk_vlan+","+iface_trunk_vlan_add
            else:
                iface_trunk_vlan = ''
            iface_trunk_vlan = str (iface_trunk_vlan)
            output_dictionary[hostname][iface_name]["trunk_vlan"] = iface_trunk_vlan

            # Collect native vlan trunk informations
            iface_trunk_native = iface_param.re_search_children ("switchport trunk native vlan")
            if iface_trunk_native:
                iface_trunk_native = iface_trunk_native[0].text.replace (' switchport trunk native vlan ', '')
            else:
                iface_trunk_native = ''
            output_dictionary[hostname][iface_name]["trunk_native"] = iface_trunk_native

            # Collect full previous configuration for rollback
            previous_config = iface_param.re_search_children ('')
            previous_config_list = []
            for param in previous_config:
                previous_config_list.append (param.text)
            output_dictionary[hostname][iface_name]["previous_config"] = previous_config_list
    return


def dict_to_xlsx(dictionary,out_file):
    """
    This function parse a dictionary and create excel files based on IT.
    The first level of parameters will become filenames, the second tabs in excel, the third lines
    :param dictionary: the name of the dictionary
    :param out_file: excel file
    """
    import os
    import xlsxwriter
    # Check if file exists and remove previous one
    if os.path.isfile(out_file):
        os.remove(out_file)
    # Create an excel workbook
    workbook = xlsxwriter.Workbook(out_file)
    # Create excel worksheet with keys
    for key in dictionary.keys():
        #Add one worksheet per key (hostname)
        worksheet = workbook.add_worksheet(key)
        # Start from the first cell
        row = 0
        col = 0
        for item in dictionary[key].keys():
            worksheet.write(row, col, "Interface")
            for subitem in dictionary[key][item].keys():
                col += 1
                worksheet.write(row,col,subitem)
            row += 1
            break
        # Parse dictionary
        col = 0
        for item in dictionary[key]:
            worksheet.write(row, col, item)
            worksheet.set_column(row,col,30)
            for subitem in dictionary[key][item].values():
                col += 1
                worksheet.write(row, col, str(subitem))
                worksheet.set_column(row, col, 30)
            row += 1
            col = 0

    workbook.close()


# Vars
directory = r"PATH_WHERE_FILES_ARE_LOCATED"
dict = {}
out_file = "output.xlsx"

# Main
file_list = search_file_in_directory(directory,filter="confg")
for file in file_list:
    sh_run_to_dict (file,dict,"ios")
dict_to_xlsx(dict,out_file)
