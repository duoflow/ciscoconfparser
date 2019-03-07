import re
import openpyxl
from openpyxl.styles import Alignment
from datetime import datetime
import logging


class ConfigParser:
    def __init__(self):
        self.debugmode = 1
        self.timestamp = str(datetime.now().strftime('%m%d%Y-%H%M%S'))
        logging.basicConfig(filename='log/' + self.timestamp + '_ciscoparserlog.log', filemode='w',
                            format='%(asctime)s: %(name)s - %(levelname)s - %(message)s', datefmt='%d%m%Y %H:%M:%S',
                            level=logging.DEBUG)
        self.logger = logging.getLogger('cplogger')
        if self.debugmode:
            self.logger.debug('Class created')

    # determite interface type
    def defineinterfacetype(self, configstrings):
        if self.debugmode:
            print('defineinterfacetype(): started')
        regex_Static_L3 = r"(^ip.*\saddress\s\d)"
        pattern_Static_L3 = re.compile(regex_Static_L3, re.IGNORECASE)
        #
        regex_IPoE_L3 = r"(^initiator\sdhcp)|(^initiator.unclassified.ip-address.)"
        pattern_IPoE_L3 = re.compile(regex_IPoE_L3, re.IGNORECASE)
        #
        regex_PPPoE = r"pppoe.enable.group."
        pattern_PPPoE = re.compile(regex_PPPoE, re.IGNORECASE)

        regex_QinQ = r".*qinq.*"
        pattern_QinQ = re.compile(regex_QinQ, re.IGNORECASE)
        #
        regex_Dot1q = r"(.*dot1q.*)"
        pattern_dot1q = re.compile(regex_Dot1q, re.IGNORECASE)
        #
        # make lookup to determine interface type
        interfacetype = {'type': '', 'qinq': ''}
        i = 1
        for conf_string in configstrings:
             # determine interface type
            if pattern_Static_L3.match(conf_string):
                if self.debugmode:
                    print('f()(%s): _defineInterfaceType L3-static' % i)
                interfacetype['type'] = "L3-static"
                break
            elif pattern_IPoE_L3.match(conf_string):
                if self.debugmode:
                    print('f()(%s): _defineInterfaceType IPoE' % i)
                interfacetype['type'] = "IPoE"
                break
            elif pattern_PPPoE.match(conf_string):
                if self.debugmode:
                    print('f()(%s): _defineInterfaceType PPPoE' % i)
                interfacetype['type'] = "PPPoE"
                break
            else:
                interfacetype['type'] = "L3-static"
            i += 1
        # vlan tagging type
        i = 0
        for conf_string in configstrings:
            # defile vlan encapsulation type
            if pattern_QinQ.match(conf_string):
                if self.debugmode:
                    print('defineinterfacetype():(%s): _defineInterfaceType vlan-tagging: qinq' % i)
                interfacetype['qinq'] = 'qinq'
                break
            elif pattern_dot1q.match(conf_string):
                if self.debugmode:
                    print('defineinterfacetype():(%s): _defineInterfaceType vlan-tagging: dot1q' % i)
                interfacetype['qinq'] = 'dot1q'
                break
            else:
                interfacetype['qinq'] = 'unknown'
            i += 1
        # if interface is not defined return null
        return interfacetype

    # parsing of L3 static interfaces dot1q and qinq
    def parseintl3static(self, vlantagging, intconfigstrings):
        int_params = {'vlan': [],
                      'description': [],
                      'vrf': [],
                      'ipaddr': [],
                      'shutdown': []
                      }
        # what to find
        regexpdict = {'vlan': r"(encapsulation\sdot1q\s)(\d{1,4})",
                      'description': r"(description\s)(.*)",
                      'vrf': r"(ip\svrf\sforwarding\s)(.*)",
                      'ipaddr': r"(ip\saddress\s)(\d{1,3}.\d{1,3}.\d{1,3}.\d{1,3}.*)",
                      'shutdown': r"^shutdown"
                      }
        if vlantagging == 'qinq':
            regexpdict['vlan'] = r"(^encapsulation.dot1q.)(\d{1,4})(.second-dot1q.)(.*$)"
        # start global cycle
        for key in regexpdict:
            regexp_pattern = re.compile(regexpdict[key], re.IGNORECASE)
            # start inner cycle
            duplicate_vlan = 0
            duplicate_description = 0
            duplicate_vrf = 0
            i = 0
            for conf_string in intconfigstrings:
                mlist = regexp_pattern.findall(conf_string)
                if mlist:
                    if self.debugmode:
                        print("parseintl3static(): (%s) conf: %s" % (i, conf_string))
                        print("parseintl3static(): (%s) mlist :%s" % (i, mlist))
                        print('parseintl3static(): (%s) dupl: vlan-%s, descr-%s, vrf-%s' % (i, duplicate_vlan,
                                                                                            duplicate_description,
                                                                        duplicate_vrf))
                    if vlantagging == 'dot1q' and key == 'vlan' and duplicate_vlan == 0:
                        int_params[key].append(mlist[0][1])
                        duplicate_vlan = 1
                        if self.debugmode:
                            print('parseintl3static(): (%s) vlan id was added: %s' % (i, mlist[0][1]))
                    elif vlantagging == 'qinq' and key == 'vlan':
                        int_params[key].append([mlist[0][1],mlist[0][3]])
                        duplicate_vlan = 1
                        if self.debugmode:
                            print('parseintl3static(): (%s) add vlan: pe-vid:%s ce-vid:%s' % (i, mlist[0][1],
                                                                                              mlist[0][3]))
                    elif key == 'description' and duplicate_description == 0:
                        int_params[key].append(mlist[0][1])
                        duplicate_description = 1
                        if self.debugmode:
                            print('parseintl3static(): (%s) add description: %s' % (i, mlist[0][1]))
                    elif key == 'vrf' and duplicate_vrf == 0:
                        int_params[key].append(mlist[0][1])
                        duplicate_vrf = 1
                    # IP address
                    elif key == 'ipaddr':
                        ip = re.sub(r'(\d{1,3})(\s)(\d{1,2})','\g<1>/\g<3>', mlist[0][1])
                        int_params[key].append(ip)
                        if self.debugmode:
                            print('parseintl3static(): (%s) add ipaddr: %s' % (i, ip))
                    # shutdown
                    elif key == 'shutdown':
                        int_params[key].append('Adm shutdown')
                        if self.debugmode:
                            print('parseintl3static(): (%s) add shutdown: %s' % (i, 'shutdown'))
                i += 1
        # check int params
        for key in int_params:
            if len(int_params[key]) == 0:
                int_params[key] = ['none']
        # return values
        return int_params
        # end of the f() parseintl3static

    # parsing of L3 static interfaces dot1q and qinq
    def parseintpppoe(self, vlantagging, intconfigstrings):
        # regexp dictionary
        int_params = {'vlan': [],
                      'description': [],
                      'pppoegroup': [],
                      'ipaddr': [],
                      'vrf': [],
                      'accessgroup': [],
                      'ipunnumbered': [],
                      'servicepolicy': [],
                      'shutdown': []
                      }
        regexpdict = {'vlan': r"(^encapsulation.dot1q.)(\d{1,4})",
                      'description': r"(description\s)(.*)",
                      'pppoegroup': r"(^pppoe.enable.group.)(.*)",
                      'ipaddr': r"(^ip.*address.)(.*)",
                      'vrf': r"(ip.vrf.forwarding\s)(.*)",
                      'accessgroup': r"(^ip.access-group.)(.*)",
                      'ipunnumbered': r"(^ip.unnumbered.)(.*)",
                      'servicepolicy': r"(^service-policy.type.)(.*)",
                      'shutdown': r"^shutdown"
                      }
        # what to find
        if vlantagging == 'qinq':
            regexpdict['vlan'] = r"(^encapsulation.dot1q.)(\d{1,4})(.second-dot1q.)(.*)"
        # start global cycle
        for key in regexpdict:
            regexp_pattern = re.compile(regexpdict[key], re.IGNORECASE)
            # start inner cycle
            i = 0
            duplflag = {'description': 0,
                        'vlan': 0,
                        'pppoegroup': 0,
                        'ipaddr': 0,
                        'vrf': 0,
                        'accessgroup': 0
                        }
            # parse conf strings
            for conf_string in intconfigstrings:
                # match all entries
                mlist = regexp_pattern.findall(conf_string)
                if mlist:
                    if self.debugmode:
                        print("parseintpppoe(): (%s) int conf: %s" % (i, conf_string))
                        print("parseintpppoe(): (%s) mlist :%s" % (i, mlist))
                    if vlantagging == 'dot1q' and key == 'vlan' and duplflag[key] == 0:
                        int_params[key].append(mlist[0][1])
                        duplflag[key] = 1
                        if self.debugmode:
                            print('parseintpppoe(): (%s) add vlan: %s' % (i, mlist[0][1]))
                    elif vlantagging == 'qinq' and key == 'vlan' and duplflag[key] == 0:
                        int_params['vlan'].append([mlist[0][3], mlist[0][1]])
                        duplflag[key] = 1
                        if self.debugmode:
                            print('parseintpppoe(): (%s) add vlan: pe-vid:%s ce-vid:%s' % (i, mlist[0][3], mlist[0][1]))
                    elif key == 'description' and duplflag[key] == 0:
                        int_params[key].append(mlist[0][1])
                        duplflag[key] = 1
                        if self.debugmode:
                            print('parseintpppoe(): (%s) add description: %s' % (i, mlist[0][1]))
                    elif key == 'pppoegroup' and duplflag[key] == 0:
                        int_params[key].append(mlist[0][1])
                        duplflag[key] = 1
                        if self.debugmode:
                            print('parseintpppoe(): (%s) add pppoegroup: %s' % (i, mlist[0][1]))
                    elif key == 'ipaddr':
                        int_params[key].append(mlist[0][1])
                        duplflag[key] = 1
                        if self.debugmode:
                            print('parseintpppoe(): (%s) add ipaddr: %s' % (i, mlist[0][1]))
                    elif key == 'vrf' and duplflag[key] == 0:
                        int_params[key].append(mlist[0][1])
                        duplflag[key] = 1
                        if self.debugmode:
                            print('parseintpppoe(): (%s) add vrf: %s' % (i, mlist[0][1]))
                    elif key == 'accessgroup':
                        int_params[key].append(mlist[0][1])
                        if self.debugmode:
                            print('parseintpppoe(): (%s) add %s: %s' % (i, key, mlist[0][1]))
                    elif key == 'ipunnumbered' and duplflag[key] == 0:
                        int_params[key].append(mlist[0][1])
                        duplflag[key] = 1
                        if self.debugmode:
                            print('parseintpppoe(): (%s) add %s: %s' % (i, key, mlist[0][1]))
                    elif key == 'servicepolicy' and duplflag[key] == 0:
                        int_params[key].append(mlist[0][1])
                        duplflag[key] = 1
                        if self.debugmode:
                            print('parseintpppoe(): (%s) add %s: %s' % (i, key, mlist[0][1]))
                    elif key == 'shutdown' and duplflag[key] == 0:
                        int_params[key].append('Adm shutdown')
                        duplflag[key] = 1
                        if self.debugmode:
                            print('parseintpppoe(): (%s) add %s: %s' % (i, key, mlist[0][1]))
                i += 1
        # check int params
        for key in int_params:
            if len(int_params[key]) == 0:
                int_params[key] = ['none']
        # return values
        return int_params
        # end of the function

    # parsing of IPoE interfaces dot1q and qinq
    def parseintipoe(self, vlantagging, intconfigstrings):
        int_params = {'description': [],
                      'vlan': [],
                      'vrf': [],
                      'ipaddr': [],
                      'ipunnumbered': [],
                      'dhcprelayinfo': [],
                      'accessgroup': [],
                      'servicepolicy': [],
                      'ipsubstype': [],
                      'initiator': [],
                      'shutdown': []
                      }
        # what to find
        regexpdict = {'description': r"(description.)(.*)",
                      'vlan': r"(^encapsulation.dot1q.)(\d{1,4})",
                      'vrf': r"(ip.vrf.forwarding\s)(.*)",
                      'ipaddr': r"(^ip.address.)(\d{1,3}.\d{1,3}.\d{1,3}.\d{1,3}.*)",
                      'ipunnumbered': r"(^ip.unnumbered.)(.*)",
                      'dhcprelayinfo': r"(^ip.dhcp.relay.information.)(.*)",
                      'accessgroup': r"(^ip.access-group.)(.*)",
                      'servicepolicy': r"(^service-policy.type.)(.*)",
                      'ipsubstype': r"(^ip.subscriber.)(.*)",
                      'initiator': r"(^initiator.)(.*)",
                      'shutdown': r"^shutdown"
                      }
        if vlantagging == 'qinq':
            regexpdict['vlan'] = r"(^encapsulation\sdot1q\s)(\d{1,4})"
        # start global cycle
        for key in regexpdict:
            regexp_pattern = re.compile(regexpdict[key], re.IGNORECASE)
            # start inner cycle
            duplflag = {'description': 0,
                        'vlan': 0,
                        'vrf': 0,
                        'ipaddr': 0,
                        'ipunnumbered': 0,
                        'dhcprelayinfo': 0,
                        'accessgroup': 0,
                        'servicepolicy': 0,
                        'ipsubstype': 0,
                        'initiator': 0,
                        'shutdown': 0
                        }
            i = 0
            for conf_string in intconfigstrings:
                mlist = regexp_pattern.findall(conf_string)
                if mlist:
                    if self.debugmode:
                        print("parseintipoe(): (%s) conf: %s" % (i, conf_string))
                        print("parseintipoe(): (%s) mlist :%s" % (i, mlist))
                    # vlan dot1q
                    if vlantagging == 'dot1q' and key == 'vlan' and duplflag['vlan'] == 0:
                        int_params[key].append(mlist[0][1])
                        duplflag['vlan'] = 1
                        if self.debugmode:
                            print('parseintipoe(): (%s) add vlan: %s' % (i, mlist[0]))
                    # vlan qinq
                    elif vlantagging == 'qinq' and key == 'vlan':
                        int_params['vlan'].append([mlist[0][3], mlist[0][1]])
                        duplflag['vlan'] = 1
                        if self.debugmode:
                            print('parseintipoe(): (%s) add vlan: pe-vid:%s ce-vid:%s' % (i, mlist[0][3], mlist[0][1]))
                    # vlan-range qinq
                    elif vlantagging == 'qinq' and key == 'vlan-range':
                        int_params['vlan'].append([mlist[0][3], "%s-%s" % (mlist[0][1], mlist[0][2])])
                        duplflag['vlan'] = 1
                        if self.debugmode:
                            print('parseintipoe(): (%s) add vlan: pe-vid:%s ce-vid:%s' % (i, mlist[0][3], [(mlist[0][1], mlist[0][2])]))
                    # description
                    elif key == 'description' and duplflag['description'] == 0:
                        int_params[key].append(mlist[0][1])
                        duplflag['description'] = 1
                        if self.debugmode:
                            print('parseintipoe(): (%s) add description: %s' % (i, mlist[0][1]))
                    # vrf
                    elif key == 'vrf' and duplflag['vrf'] == 0:
                        int_params[key].append(mlist[0][1])
                        duplflag['vrf'] = 1
                        if self.debugmode:
                            print('parseintipoe(): (%s) add vrf: %s' % (i, mlist[0][1]))
                    # ipaddr
                    elif key == 'ipaddr':
                        int_params[key].append(mlist[0][1])
                        duplflag['ipaddr'] = 1
                        if self.debugmode:
                            print('parseintipoe(): (%s) add ipaddr: %s' % (i, mlist[0][1]))
                    # ipunnumbered
                    elif key == 'ipunnumbered' and duplflag['ipunnumbered'] == 0:
                        int_params[key].append(mlist[0][1])
                        duplflag['ipunnumbered'] = 1
                        if self.debugmode:
                            print('parseintipoe(): (%s) add ipunnumbered: %s' % (i, mlist[0][1]))
                    # dhcprelayinfo
                    elif key == 'dhcprelayinfo':
                        int_params[key].append(mlist[0][1])
                        duplflag['dhcprelayinfo'] = 1
                        if self.debugmode:
                            print('parseintipoe(): (%s) add dhcprelayinfo: %s' % (i, mlist[0][1]))
                    # accessgroup
                    elif key == 'accessgroup':
                        int_params[key].append(mlist[0][1])
                        duplflag['accessgroup'] = 1
                        if self.debugmode:
                            print('parseintipoe(): (%s) add accessgroup: %s' % (i, mlist[0][1]))
                    # servicepolicy
                    elif key == 'servicepolicy' and duplflag['servicepolicy'] == 0:
                        int_params[key].append(mlist[0][1])
                        duplflag['servicepolicy'] = 1
                        if self.debugmode:
                            print('parseintipoe(): (%s) add servicepolicy: %s' % (i, mlist[0][1]))
                    # ipsubstype
                    elif key == 'ipsubstype' and duplflag['ipsubstype'] == 0:
                        int_params[key].append(mlist[0][1])
                        duplflag['ipsubstype'] = 1
                        if self.debugmode:
                            print('parseintipoe(): (%s) add ipsubstype: %s' % (i, mlist[0][1]))
                    # initiator
                    elif key == 'initiator' and duplflag['initiator'] == 0:
                        int_params[key].append(mlist[0][1])
                        duplflag['initiator'] = 1
                        if self.debugmode:
                            print('parseintipoe(): (%s) add initiator: %s' % (i, mlist[0][1]))
                    # shutdown
                    elif key == 'shutdown' and duplflag['shutdown'] == 0:
                        int_params[key].append('Adm shutdown')
                        duplflag['shutdown'] = 1
                        if self.debugmode:
                            print('parseintipoe(): (%s) add shutdown: %s' % (i, mlist[0]))
                i += 1
        # check int params
        for key in int_params:
            if len(int_params[key]) == 0:
                int_params[key] = ['none']
        # return values
        return int_params
        # end of the function

    # determite if string includes interface number
    def findinterfaceinstring(self, configline):
        regex_interface = r'(^interface.*)'
        pattern_interface = re.compile(regex_interface)
        intnumber = pattern_interface.findall(configline)
        if len(intnumber):
            return intnumber[0]
        else:
            return 0

    # interface parameter collector
    def collectintparameters(self, configlines, position=0):
        # start process
        interface_params = []
        position += 1
        while position < len(configlines):
            interface = self.findinterfaceinstring(configlines[position])
            if interface == 0:
                param = re.sub(r"(\t|\s{2,4}|^\s{1,3}|\n)", "", configlines[position])
                if len(param) > 2:
                    interface_params.append(param)
                position += 1
            else:
                break
        return {'position': position-1, 'params': interface_params}

    # create interface dictionary
    def createinterfacedictionary(self, configlines):
        # start process
        i = 0
        # create interface dictionary
        int_dictionary = {}
        # start to populate int dictionary
        numconfiglines = len(configlines)
        while i < numconfiglines:
            interface = self.findinterfaceinstring(configlines[i])
            # if it is an interface
            if interface:
                # define interface parameters
                # print("Start to collect parameters. Step (%s)"%i)
                result = self.collectintparameters(configlines, i)
                interface_params = result['params']
                # set interface parameters
                int_dictionary[interface] = interface_params
                i = result['position']
            # increase counter
            i = i + 1
        return int_dictionary

    # write header to excel file
    def writeExcelHeader(self, wsheet):
        wsheet['A1'] = "#"
        wsheet['B1'] = "Interface number"
        wsheet['C1'] = "Sub-If number"
        wsheet['D1'].alignment = Alignment(wrapText=True)
        wsheet['D1'] = "Interface type\nL3-static, IPoE, PPPoE"
        wsheet['E1'] = "Int description"
        wsheet['F1'] = "VRF"
        wsheet['G1'] = "IP Address"
        wsheet['H1'] = "IP unnumbered interface"
        wsheet['I1'] = "Dot1Q VID"
        wsheet['J1'] = "QinQ PE-VID"
        wsheet['K1'] = "QinQ CE-VID"
        wsheet['L1'].alignment = Alignment(wrapText=True)
        wsheet['L1'] = "DHCP Relay"
        wsheet['M1'].alignment = Alignment(wrapText=True)
        wsheet['M1'] = "Access group"
        wsheet['N1'].alignment = Alignment(wrapText=True)
        wsheet['N1'] = "Service policy"
        wsheet['O1'].alignment = Alignment(wrapText=True)
        wsheet['O1'] = "Subscriber type"
        wsheet['P1'].alignment = Alignment(wrapText=True)
        wsheet['P1'] = "Auth initiator"
        wsheet['Q1'].alignment = Alignment(wrapText=True)
        wsheet['Q1'] = "PPPoe enable group"
        wsheet['R1'] = "Interface state"

    # write to excel L3 static
    def writeexcell3static(self, wsheet, str_index, int_type, int_name, int_params):
        #
        if self.debugmode:
            print('f() writeexcell3static started')
        #
        regexp_int = r"(interface\s)(.*)"
        regexp_pattern = re.compile(regexp_int, re.IGNORECASE)
        int_main_name = regexp_pattern.findall(int_name)[0][1]
        # int_main_name = ''
        # for L3 static Dot1Q interface
        if int_type['type'] == "L3-static" and int_type['qinq'] != 'qinq':
            int_params = self.parseintl3static(int_type['qinq'], int_params)
            ipinfo = ''
            for ip in int_params['ipaddr']:
                ipinfo += "\n" + ip
            # define excell cells
            # data string index
            wsheet['A' + str(str_index)] = str_index-1
            # main interface name
            wsheet['B' + str(str_index)] = int_main_name
            # sub interface name
            wsheet['C' + str(str_index)] = int_name
            # interface type
            wsheet['D' + str(str_index)] = "%s %s" % (int_type['type'], int_type['qinq'])
            # interface description
            wsheet['E' + str(str_index)] = int_params['description'][0]
            # VRF
            wsheet['F' + str(str_index)] = int_params['vrf'][0]
            # IP addresses
            wsheet['G' + str(str_index)].alignment = Alignment(wrapText=True)
            wsheet['G' + str(str_index)] = ipinfo
            # IP unnumbered interface
            wsheet['H' + str(str_index)] = '―'
            # Dot1Q VID
            wsheet['I' + str(str_index)] = int_params['vlan'][0]
            # QinQ interface pe-vid
            wsheet['J' + str(str_index)] = '―'
            # QinQ interface ce-vid
            wsheet['K' + str(str_index)] = '―'
            # DHCP Relay
            wsheet['L' + str(str_index)] = '―'
            # Access group
            wsheet['M' + str(str_index)] = '―'
            # Service policy
            wsheet['N' + str(str_index)] = '―'
            # Subscriber type
            wsheet['O' + str(str_index)] = '―'
            # Auth initiator
            wsheet['P' + str(str_index)] = '―'
            # PPPoE enable group
            wsheet['Q' + str(str_index)] = '―'
            # Shutdown state
            wsheet['R' + str(str_index)] = int_params['shutdown'][0]
        # for L3 static QinQ interface
        elif int_type['type'] == "L3-static" and int_type['qinq'] == 'qinq':
            int_params = self.parseintl3static(int_type['qinq'], int_params)
            ipinfo = ''
            for ip in int_params['ipaddr']:
                ipinfo += "\n"
                ipinfo += ip
            # define excell cells
            for vid in int_params['vlan']:
                # define excell cells
                # data string index
                wsheet['A' + str(str_index)] = str_index - 1
                # main interface name
                wsheet['B' + str(str_index)] = int_main_name
                # sub interface name
                wsheet['C' + str(str_index)] = int_name
                # interface type
                wsheet['D' + str(str_index)] = "%s %s" % (int_type['type'], int_type['qinq'])
                # interface description
                wsheet['E' + str(str_index)] = int_params['description'][0]
                # VRF
                wsheet['F' + str(str_index)] = int_params['vrf'][0]
                # IP addresses
                wsheet['G' + str(str_index)].alignment = Alignment(wrapText=True)
                wsheet['G' + str(str_index)] = ipinfo
                # IP unnumbered interface
                wsheet['H' + str(str_index)] = '―'
                # Dot1Q VID
                wsheet['I' + str(str_index)] = int_params['vlan'][0]
                # QinQ interface pe-vid
                wsheet['J' + str(str_index)] = vid[0]
                # QinQ interface ce-vid
                wsheet['K' + str(str_index)] = vid[1]
                # DHCP Relay
                wsheet['L' + str(str_index)] = '―'
                # Access group
                wsheet['M' + str(str_index)] = '―'
                # Service policy
                wsheet['N' + str(str_index)] = '―'
                # Subscriber type
                wsheet['O' + str(str_index)] = '―'
                # Auth initiator
                wsheet['P' + str(str_index)] = '―'
                # PPPoE enable group
                wsheet['Q' + str(str_index)] = '―'
                # Shutdown state
                wsheet['R' + str(str_index)] = int_params['shutdown'][0]
                # increment string number
                str_index += 1
        # return excel string number
        return str_index

    # write to excel L3 static
    def writeexcelpppoe(self, wsheet, str_index, int_type, int_name, int_params):
        if self.debugmode:
            print('writeexcelpppoe() started')
        #
        regexp_int = r"(interface\s)(.*)(\..*)"
        regexp_pattern = re.compile(regexp_int, re.IGNORECASE)
        int_main_name = regexp_pattern.findall(int_name)[0][1]
        # for L3 static Dot1Q interface
        if int_type['type'] == "PPPoE" and int_type['qinq'] != 'qinq':
            int_params = self.parseintpppoe(int_type['qinq'], int_params)
            if self.debugmode == 1:
                print('writeexcelpppoe(): int params for %s:' % int_name)
                print(int_params)
            # ipaddress information
            ipinfo = ''
            for ip in int_params['ipaddr']:
                ipinfo += "\n" + ip
            # accessgroup information
            accessgroup = ''
            for line in int_params['accessgroup']:
                accessgroup += line + '\n'
            # define excell cells
            # data string index
            wsheet['A' + str(str_index)] = str_index-1
            # main interface name
            wsheet['B' + str(str_index)] = int_main_name
            # sub interface name
            wsheet['C' + str(str_index)] = int_name
            # interface type
            wsheet['D' + str(str_index)] = "%s %s" % (int_type['type'], int_type['qinq'])
            # interface description
            wsheet['E' + str(str_index)] = int_params['description'][0]
            # VRF
            wsheet['F' + str(str_index)] = int_params['vrf'][0]
            # IP addresses
            wsheet['G' + str(str_index)].alignment = Alignment(wrapText=True)
            wsheet['G' + str(str_index)] = ipinfo
            # IP unnumbered interface
            wsheet['H' + str(str_index)] = int_params['ipunnumbered'][0]
            # Dot1Q VID
            wsheet['I' + str(str_index)] = int_params['vlan'][0]
            # QinQ interface pe-vid
            wsheet['J' + str(str_index)] = '―'
            # QinQ interface ce-vid
            wsheet['K' + str(str_index)] = '―'
            # DHCP Relay
            wsheet['L' + str(str_index)] = '―'
            # Access group
            wsheet['M' + str(str_index)] = accessgroup
            # Service policy
            wsheet['N' + str(str_index)] = int_params['servicepolicy'][0]
            # Subscriber type
            wsheet['O' + str(str_index)] = '―'
            # Auth initiator
            wsheet['P' + str(str_index)] = '―'
            # PPPoE enable group
            wsheet['Q' + str(str_index)] = int_params['pppoegroup'][0]
            # Shutdown state
            wsheet['R' + str(str_index)] = int_params['shutdown'][0]
        # for PPPoE QinQ interface
        elif int_type['type'] == "PPPoE" and int_type['qinq'] == 'qinq':
            int_params = self.parseintpppoe(int_type['qinq'], int_params)
            # ipinfo
            ipinfo = ''
            for ip in int_params['ipaddr']:
                ipinfo += "\n"
                ipinfo += ip
            # accessgroup information
            accessgroup = ''
            for line in int_params['accessgroup']:
                accessgroup += line + '\n'
            # define excell cells
            for vid in int_params['vlan']:
                # define excell cells
                # data string index
                wsheet['A' + str(str_index)] = str_index - 1
                # main interface name
                wsheet['B' + str(str_index)] = int_main_name
                # sub interface name
                wsheet['C' + str(str_index)] = int_name
                # interface type
                wsheet['D' + str(str_index)] = "%s %s" % (int_type['type'], int_type['qinq'])
                # interface description
                wsheet['E' + str(str_index)] = int_params['description'][0]
                # VRF
                wsheet['F' + str(str_index)] = int_params['vrf'][0]
                # IP addresses
                wsheet['G' + str(str_index)].alignment = Alignment(wrapText=True)
                wsheet['G' + str(str_index)] = ipinfo
                # IP unnumbered interface
                wsheet['H' + str(str_index)] = int_params['ipunnumbered'][0]
                # Dot1Q VID
                wsheet['I' + str(str_index)] = '―'
                # QinQ interface pe-vid
                wsheet['J' + str(str_index)] = vid[0]
                # QinQ interface ce-vid
                wsheet['K' + str(str_index)] = vid[1]
                # DHCP Relay
                wsheet['L' + str(str_index)] = '―'
                # Access group
                wsheet['M' + str(str_index)] = accessgroup
                # Service policy
                wsheet['N' + str(str_index)] = int_params['servicepolicy'][0]
                # Subscriber type
                wsheet['O' + str(str_index)] = '―'
                # Auth initiator
                wsheet['P' + str(str_index)] = '―'
                # PPPoE enable group
                wsheet['Q' + str(str_index)] = int_params['pppoegroup'][0]
                # Shutdown state
                wsheet['R' + str(str_index)] = int_params['shutdown'][0]
                # increment string number
                str_index += 1
        # return excel string number
        return str_index

    # write to excel IPoE static
    def writeexcelipoe(self, wsheet, str_index, int_type, int_name, int_params):
        if self.debugmode:
            print('writeexcelipoe() started')
        #
        regexp_int = r"(interface\s)(.*)(\..*)"
        regexp_pattern = re.compile(regexp_int, re.IGNORECASE)
        int_main_name = regexp_pattern.findall(int_name)[0][1]
        # int_main_name = ''
        # for L3 static Dot1Q interface
        if int_type['type'] == "IPoE" and int_type['qinq'] != 'qinq':
            int_params = self.parseintipoe(int_type['qinq'], int_params)
            if self.debugmode == 1:
                print('writeexcelipoe(): int params for %s:' % int_name)
                print(int_params)
            # ipaddress information
            ipinfo = ''
            for ip in int_params['ipaddr']:
                ipinfo += "\n" + ip
            # dhcprelayinfo information
            dhcprelayinfo = ''
            for line in int_params['dhcprelayinfo']:
                dhcprelayinfo += line + '\n'
            # accessgroup information
            accessgroup = ''
            for line in int_params['accessgroup']:
                accessgroup += line + '\n'
            # define excell cells
            # data string index
            wsheet['A' + str(str_index)] = str_index-1
            # main interface name
            wsheet['B' + str(str_index)] = int_main_name
            # sub interface name
            wsheet['C' + str(str_index)] = int_name
            # interface type
            wsheet['D' + str(str_index)] = "%s %s" % (int_type['type'], int_type['qinq'])
            # interface description
            wsheet['E' + str(str_index)] = int_params['description'][0]
            # VRF
            wsheet['F' + str(str_index)] = int_params['vrf'][0]
            # IP addresses
            wsheet['G' + str(str_index)].alignment = Alignment(wrapText=True)
            wsheet['G' + str(str_index)] = ipinfo
            # IP unnumbered interface
            wsheet['H' + str(str_index)] = int_params['ipunnumbered'][0]
            # Dot1Q VID
            wsheet['I' + str(str_index)] = int_params['vlan'][0]
            # QinQ interface pe-vid
            wsheet['J' + str(str_index)] = '―'
            # QinQ interface ce-vid
            wsheet['K' + str(str_index)] = '―'
            # DHCP Relay
            wsheet['L' + str(str_index)] = dhcprelayinfo
            # Access group
            wsheet['M' + str(str_index)] = accessgroup
            # Service policy
            wsheet['N' + str(str_index)] = int_params['servicepolicy'][0]
            # Subscriber type
            wsheet['O' + str(str_index)] = int_params['ipsubstype'][0]
            # Auth initiator
            wsheet['P' + str(str_index)] = int_params['initiator'][0]
            # PPPoE enable group
            wsheet['Q' + str(str_index)] = '―'
            # Shutdown state
            wsheet['R' + str(str_index)] = int_params['shutdown'][0]
        # for L3 static QinQ interface
        elif int_type['type'] == "IPoE" and int_type['qinq'] == 'qinq':
            int_params = self.parseintl3static(int_type['qinq'], int_params)
            ipinfo = ''
            for ip in int_params['ipaddr']:
                ipinfo += "\n"
                ipinfo += ip
            # dhcprelayinfo information
            dhcprelayinfo = ''
            for line in int_params['dhcprelayinfo']:
                dhcprelayinfo += line + '\n'
            # accessgroup information
            accessgroup = ''
            for line in int_params['accessgroup']:
                accessgroup += line + '\n'
            # define excell cells
            for vid in int_params['vlan']:
                # define excell cells
                # data string index
                wsheet['A' + str(str_index)] = str_index - 1
                # main interface name
                wsheet['B' + str(str_index)] = int_main_name
                # sub interface name
                wsheet['C' + str(str_index)] = int_name
                # interface type
                wsheet['D' + str(str_index)] = "%s %s" % (int_type['type'], int_type['qinq'])
                # interface description
                wsheet['E' + str(str_index)] = int_params['description'][0]
                # VRF
                wsheet['F' + str(str_index)] = int_params['vrf'][0]
                # IP addresses
                wsheet['G' + str(str_index)].alignment = Alignment(wrapText=True)
                wsheet['G' + str(str_index)] = ipinfo
                # IP unnumbered interface
                wsheet['H' + str(str_index)] = int_params['ipunnumbered'][0]
                # Dot1Q VID
                wsheet['I' + str(str_index)] = '―'
                # QinQ interface pe-vid
                wsheet['J' + str(str_index)] = vid[0]
                # QinQ interface ce-vid
                wsheet['K' + str(str_index)] = vid[1]
                # DHCP Relay
                wsheet['L' + str(str_index)] = dhcprelayinfo
                # Access group
                wsheet['M' + str(str_index)] = accessgroup
                # Service policy
                wsheet['N' + str(str_index)] = int_params['servicepolicy'][0]
                # Subscriber type
                wsheet['O' + str(str_index)] = int_params['ipsubstype'][0]
                # Auth initiator
                wsheet['P' + str(str_index)] = int_params['initiator'][0]
                # PPPoE enable group
                wsheet['Q' + str(str_index)] = '―'
                # Shutdown state
                wsheet['R' + str(str_index)] = int_params['shutdown'][0]
                # increment string number
                str_index += 1
        # return excel string number
        return str_index

    # f() writetofile - write excel file from RAM do file system
    def writetofile(self, interfaces):
        if self.debugmode != 0:
            print('writetofile() started')
        # create excel file
        wb = openpyxl.Workbook()
        worksheet = wb.active
        worksheet.title = 'interfaces'
        # worksheet header
        self.writeExcelHeader(worksheet)
        # worksheet's strings
        str_num = 2
        for key in interfaces:
            # determine interface type
            int_type = self.defineinterfacetype(interfaces[key])
            if self.debugmode != 0:
                print("main cycle (iter:%s) " % (str_num - 1))
                print(key, interfaces[key])
                print('interface type: %s' % int_type)
            #
            if int_type['type'] == "L3-static":
                if self.debugmode != 0:
                    print("main cycle (iter:%s) L3-static int:%s" % ((str_num - 1), key))
                str_num = self.writeexcell3static(worksheet, str_num, int_type, key, interfaces[key])
                str_num += 1
            elif int_type['type'] == "IPoE":
                if self.debugmode != 0:
                    print("main cycle (iter:%s) IPoE int:%s:" % ((str_num - 1), key))
                str_num = self.writeexcelipoe(worksheet, str_num, int_type, key, interfaces[key])
                str_num += 1
            elif int_type['type'] == "PPPoE":
                if self.debugmode != 0:
                    print("main cycle (iter:%s) PPPoE int:%s:" % ((str_num - 1), key))
                str_num = self.writeexcelpppoe(worksheet, str_num, int_type, key, interfaces[key])
                str_num += 1
        # save information to file
        try:
            wb.save("%s_results.xlsx" % self.timestamp)
            print('Write to file. Lines (%s)' % worksheet.max_row)
        except IOError:
            print("Error: I can\'t create excel file")
        # end of f() writetofile
#
# start script
#
config = 0
try:
    filename1 = 'device.cfg/b-i-1.int.cfg'
    filename2 = 'device.cfg/b-i-2.int.cfg'
    filename3 = 'device.cfg/pppoe.cfg'
    f = open(filename3)
except IOError:
    print("Error: I can\'t find file or read data")
else:
    configlines = f.readlines()
    f.close()

# find all interfaces
if configlines != 0:
    parser = ConfigParser()
    interfaces = parser.createinterfacedictionary(configlines)
    parser.writetofile(interfaces)

