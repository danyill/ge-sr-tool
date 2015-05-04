#!/usr/bin/python

# NB: As per Debian Python policy #!/usr/bin/env python2 is not used here.

"""
ge-sr-tool.py

A tool to browse GE SR760 relay files to extract parameter information
intended for bulk processing.

Usage defined by running with option -h.

This tool can be run from the IDLE prompt using the main def.

Thoughtful ideas most welcome. 

Installation instructions (for Python *2.7.9*):
 - pip install tablib
 - or if behind a proxy server: pip install --proxy="user:password@server:port" packagename
 - within Transpower: pip install --proxy="transpower\mulhollandd:password@tptpxy001.transpower.co.nz:8080" tablib    

TODO: 
 - sorting options on display and dump output?    
 - sort out guessing of Transpower standard design version
 - Providing an interface which gives coherence between different firmware versions
 - Processing newer relays as just csv files? Might be less error prone
"""

__author__ = "Daniel Mulholland"
__copyright__ = "Copyright 2015, Daniel Mulholland"
__credits__ = ['Stuart Sim for input information', "Kenneth Reitz https://github.com/kennethreitz/tablib"]
__license__ = "GPL"
__version__ = '0.05'
__maintainer__ = "Daniel Mulholland"
__hosted__ = "https://github.com/danyill/ge-sr-tool"
__email__ = "dan.mulholland@gmail.com"

# update this line if using from Python interactive
#__file__ = r'W:\Education\Current\pytooldev\ge-sr-tool'

import sys
import os
import argparse
import glob
import re
import tablib
import csv

SR760_EXTENSION = '760'
INPUT_DATA_FILE = 'input_information_sr760.csv'
OUTPUT_FILE_NAME = "output"

# not sure the backslashes are required or correct here
INPUT_SPLIT_EXPRESSION = r'\"([\w :\.#\(\)\/\+\-\%~]*)"'
BASE_PATH = os.path.dirname(os.path.realpath(__file__))
SR760_SETPOINT_EXPR_START = '^' + '0x' 
SR760_SETPOINT_EXPR_END = '=' + '(' + r'[\w :+/\\()!,.\-_\\*]*' + ')'
REGEX_NEW_RELAY_PARAMETER_EXTRACT_START = r'^(DATA|750PC_DATA),' 
REGEX_NEW_RELAY_PARAMETER_EXTRACT_END = r',([\w :+/\\()!,.\-_\\*]*),([\w :+/\\\(\)!,.\-_\\*]*)\r\n'
# ^(DATA|750PC_DATA),Underfrequency 2 Pickup(Setpoints),([\w :+/\\()!,.\-_\\*]*),([\w :+/\\()!,.\-_\\*]*)\r\n
# this probably needs to be expanded

LEGACY_NEW_FILESIZE = 30000 # this is kind of dumb but probably effective

MY_SEPARATOR = "~"
OUTPUT_HEADERS = ['Filename','Setting Name','Value']

def main(arg=None):
    parser = argparse.ArgumentParser(
        description='Process individual or multiple ' +SR760_EXTENSION+ ' files and produce summary'\
            ' of results as a csv or xls file.',
        epilog='Enjoy. Bug reports and feature requests welcome. Feel free to build a GUI :-)',
        prefix_chars='-/')

    parser.add_argument('-o', choices=['csv','xlsx'],
                        help='Produce output as either comma separated values (csv) or as'\
                        ' a Micro$oft Excel .xlsx spreadsheet. If no output provided then'\
                        ' output is to the screen.')
    
    parser.add_argument('-n', '--new', action="store_true",
                       help='New relay types firmware version 7.44')                       
    
    parser.add_argument('-l', '--legacy', action="store_true",
                       help='Old relay types firmware version 4.??')     
                       
    parser.add_argument('path', metavar='PATH|FILE', nargs=1, 
                       help='Go recursively go through path PATH. Redundant if FILE'\
                       ' with extension .'+SR760_EXTENSION+' is used. When recursively' \
                       'called, only searches for files with:' + SR760_EXTENSION \
                       + '. Globbing is allowed with the * and ? characters.'\
                       )

    parser.add_argument('-s', '--screen', action="store_true",
                       help='Show output to screen')

    # Not implemented yet
    #parser.add_argument('-d', '--design', action="store_true",
    #                   help='Attempt to determine Transpower standard design version and' \
    #                   ' include this information in output')
                       
    parser.add_argument('settings', metavar='C~S', type=str, nargs='+',
                       help='Settings in the form of C~S where C is the category' \
                       ' and S is the setting name. '
                       ' C and S are specified differently depending on whether the ' \
                       ' relay is legacy or new.'
                       ' ' \
                       ' When using the -l or --legacy flags:' \
                       ' Use information in \'input_information_sr760.csv\' using '\
                       ' columns 3 and 5.'\
                       ' ' \
                       ' When using the -n or --new flags:' \
                       ' Information is specified differently, \'TEST_NEW_SR760_FILE.760\'' \
                       ' Setpoint Group [1234]~Column 2  or' \
                       ' Device Information~Column 2 for settings prior to grouped settings.' \
                       ' ' \
                       ' Examples included with the code.'\
                       )
                       

    parser.add_argument('-v', '--version', action='version', version='%(prog)s ' + __version__)

    if arg == None:
        args = parser.parse_args()
    else:
        args = parser.parse_args(arg.split())
    
    # read in list of files
    files_to_do = return_file_paths(args.path, SR760_EXTENSION)
    
    # sort out the reference data for addresses to parameter matching
    lookup = {}
    with open(INPUT_DATA_FILE, mode='r') as csvfile:
        ref_d = csv.DictReader(csvfile)        
        for row in ref_d:
            key = row.pop('SR750/760 - V401') + MY_SEPARATOR + row.pop('Info Name')
            # we will assume no duplicates and ensure input data correct
            #if key in result:
            # implement your duplicate row handling here
            #    pass
            lookup[key] = row        
    
    if files_to_do != []:
        process_760_files(files_to_do, args, lookup)
    else:
        print('Found nothing to do for path: ' + args.path[0])
        sys.exit()
        os.system("Pause")
    
def return_file_paths(args_path, file_extension):
    paths_to_work_on = []
    for p in args_path:
        p = p.translate(None, ",\"")
        if not os.path.isabs(p):
            paths_to_work_on +=  glob.glob(os.path.join(BASE_PATH,p))
        else:
            paths_to_work_on += glob.glob(p)
            
    files_to_do = []
    # make a list of files to iterate over
    if paths_to_work_on != None:
        for p_or_f in paths_to_work_on:
            if os.path.isfile(p_or_f) == True:
                # add file to the list
                print os.path.normpath(p_or_f)
                files_to_do.append(os.path.normpath(p_or_f))
            elif os.path.isdir(p_or_f) == True:
                # walk about see what we can find
                files_to_do = walkabout(p_or_f, file_extension)
    return files_to_do        

def walkabout(p_or_f, file_extension):
    """ searches through a path p_or_f, picking up all files with EXTN
    returns these in an array.
    """
    return_files = []
    for root, dirs, files in os.walk(p_or_f, topdown=False):
        for name in files:
            if (os.path.basename(name)[-3:]).upper() == file_extension:
                return_files.append(os.path.join(root,name))
    return return_files
    
def process_760_files(files_to_do, args, reference_data):
    parameter_info = []
        
    for filename in files_to_do:      
        filesize = os.path.getsize(filename)
        extracted_data = None
        
        if args.legacy == True:
            if filesize < LEGACY_NEW_FILESIZE:
                extracted_data = extract_parameters_legacy(filename, args.settings, reference_data)
        elif args.new == True:
            if filesize >= LEGACY_NEW_FILESIZE:
                extracted_data = extract_parameters_new(filename, args.settings, reference_data)
        else:
            print 'You must specify legacy or new. Refer the help'
            sys.exit()
        
        if extracted_data != None:
            parameter_info += extracted_data

    # for exporting to Excel or CSV
    data = tablib.Dataset()    
    for k in parameter_info:
        data.append(k)
    data.headers = OUTPUT_HEADERS

    # don't overwrite existing file
    name = OUTPUT_FILE_NAME 
    if args.o == 'csv' or args.o == 'xlsx': 
        # this is stupid and klunky
        while os.path.exists(name + '.csv') or os.path.exists(name + '.xlsx'):
            name += '_'        

    # write data
    if args.o == None:
        pass
    elif args.o == 'csv':
        with open(name + '.csv','wb') as output:
            output.write(data.csv)
    elif args.o == 'xlsx':
        with open(name + '.xlsx','wb') as output:
            output.write(data.xlsx)

    if args.screen == True:
        display_info(parameter_info)

def extract_parameters_legacy(filename, s_parameters, reference_data):
    parameter_info=[]
    # read data
    with open(filename,'r') as f:
        data = f.read()

    # get input arguments in a sane format
    s_parameters = re.findall(INPUT_SPLIT_EXPRESSION,' '.join(s_parameters))
    # ^0x'17AA=([\w :+/\\()!,.\-_\\*]*)
    # [SETPOINT GROUP 1](.|\n)*^0x17AA=([\w :+/\\()!,.\-_\\*]*)
    result = None
    for sp in s_parameters:
        # there is only device information and setpoint data
        grouping = sp.split(MY_SEPARATOR)

        #^\[SETPOINT GROUP 1\]\r\n
        #val = re.search(r'^\[' \
        #    + grouping[0].upper() \
        #    + r'\]\r\n', data, flags=re.MULTILINE)
        # ^^ why is this not working?
        val = re.search(r'^\[' \
            + grouping[0].upper() \
            + r'\]$', data, flags=re.MULTILINE)
        if val is not None:  
            if sp.split(MY_SEPARATOR)[0] == 'Device Information':
                result = re.search('^' + \
                            reference_data[sp]['ADDRESS'] + \
                             SR760_SETPOINT_EXPR_END, \
                             data[val.end():], flags=re.MULTILINE).group(1)
            else:
                pattern = SR760_SETPOINT_EXPR_START + \
                            reference_data[sp]['ADDRESS'] + \
                             SR760_SETPOINT_EXPR_END
                result = re.search(pattern, data[val.end():], 
                    flags=re.MULTILINE).group(1)
        else:
            print "Grouping term not found"
        
        if result <> None:
            filename = os.path.basename(filename)
            parameter_info.append([filename, sp, result])
        else:
            print "Not matched:" + grouping[1]            
            
    return parameter_info

def extract_parameters_new(filename, s_parameters, reference_data):
    parameter_info=[]
    
    # read data
    with open(filename,'rb') as f:
        data = f.read()

    # get input arguments in a sane format
    s_parameters = re.findall(INPUT_SPLIT_EXPRESSION,' '.join(s_parameters))

    result = None
    for sp in s_parameters:
        # there is only device information and setpoint data
        grouping = sp.split(MY_SEPARATOR)
        val = re.search(r'^\[' \
            + grouping[0].upper() \
            + r'\]\r\n', data, flags=re.MULTILINE)
        if val is not None:  
            if sp.split(MY_SEPARATOR)[0] == 'Device Information':
                # match (DATA|750PC_DATA),sp[1],([\w :+/\\()!,.\-_\\*]*),([\w :+/\\()!,.\-_\\*]*)\r\n
                # ^(DATA|750PC_DATA),Underfrequency 2 Pickup\(Setpoints\),([\w :+/\\()!,.\-_\\*]*),([\w :+/\\()!,.\-_\\*]*)\r\n
                result = re.search(REGEX_NEW_RELAY_PARAMETER_EXTRACT_START + \
                             re.escape(grouping[1]) + \
                             REGEX_NEW_RELAY_PARAMETER_EXTRACT_END, \
                             data, flags=re.MULTILINE)
            else:
                pattern = REGEX_NEW_RELAY_PARAMETER_EXTRACT_START + \
                            re.escape(grouping[1]) + \
                             REGEX_NEW_RELAY_PARAMETER_EXTRACT_END
                result = re.search(pattern, data[val.end():], 
                    flags=re.MULTILINE)
        else:
            pass
            #print "Grouping term not found"
        
        if result <> None:
            filename = os.path.basename(filename)
            parameter_info.append([filename, sp, result.group(3)])
        else:
            print "Not matched:" + grouping[1]
    return parameter_info 
    
def display_info(parameter_info):
    lengths = []
    # first pass to determine column widths:
    for line in parameter_info:
        for index,element in enumerate(line):
            try:
                lengths[index] = max(lengths[index], len(element))
            except IndexError:
                lengths.append(len(element))
    
    parameter_info.insert(0, OUTPUT_HEADERS)
    # now display in columns            
    for line in parameter_info:
        display_line = '' 
        for index,element in enumerate(line):
            display_line += element.ljust(lengths[index]+2,' ')
        print display_line
   
if __name__ == '__main__':   
    
    if len(sys.argv) == 1 :
        
        """
        To select the parameters choose from input_information_sr760.csv
            Column 1~Column 5
        """
        
        # Uncomment this for help
        #main('--help')
        
        # Legacy application
        main(r'-o xlsx --legacy -s in \
        "Device Information~Version" \
        "Setpoint Group 1~Underfrequency 2 Pickup" \
        "Setpoint Group 1~Underfrequency 2 Pickup" \
        "Setpoint Group 2~Underfrequency 2 Pickup" \
        "Setpoint Group 2~Underfrequency 2 Pickup" \
        "Setpoint Group 3~Underfrequency 2 Pickup" \
        "Setpoint Group 3~Underfrequency 2 Pickup" \
        "Setpoint Group 4~Underfrequency 2 Pickup" \
        "Setpoint Group 4~Underfrequency 2 Pickup" \
        ') 

        # Also legacy application
        main(r'-o xlsx --legacy -s -o xlsx "W:\Education\Current\Stationware_Dump\20150430_Stationware_Settings_Applied" --new \
        "Device Information~Version" \
        "Setpoint Group 1~Underfrequency 1 Function" \
        "Setpoint Group 1~Underfrequency 1 Relays" \
        "Setpoint Group 1~Underfrequency 1 Pickup" \
        "Setpoint Group 1~Underfrequency 1 Pickup" \
        "Setpoint Group 1~Underfrequency 1 Delay" \
        "Setpoint Group 1~Underfrequency 1 Minimum Operating Voltage" \
        "Setpoint Group 1~Underfrequency 1 Minimum Operating Current" \
        "Setpoint Group 1~Underfrequency 2 Function" \
        "Setpoint Group 1~Underfrequency 2 Relays" \
        "Setpoint Group 1~Underfrequency 2 Pickup"')
        
        """
        New data format and name of settings has changed but is still human readable
        get the names of parameters you are interested by looking at a typical file
        
        DATA,Underfrequency 1 Function,1793,1,1,1,0,6048,0 (Disabled)
        DATA,Underfrequency 1: Relay 3,1794,1,1,1,3,6049,0 (Do Not Operate)
        DATA,Underfrequency 1: Relay 4,1795,1,1,1,4,6049,0 (Do Not Operate)
        DATA,Underfrequency 1: Relay 5,1796,1,1,1,5,6049,0 (Do Not Operate)
        DATA,Underfrequency 1: Relay 6,1797,1,1,1,6,6049,0 (Do Not Operate)
        DATA,Underfrequency 1: Relay 7,1798,1,1,1,7,6049,0 (Do Not Operate)
        DATA,Underfrequency 1 Pickup(Setpoints),1799,1,1,1,0,6050,59.00 Hz
        DATA,Underfrequency 1 Delay,1800,1,1,1,0,6051,2.00 s
        DATA,Underfrequency 1 Minimum Operating Voltage,1801,1,1,1,0,6052,0.70 x VT
        DATA,Underfrequency 1 Minimum Operating Current
        """

        """
        To select the parameters choose from info/TEST_NEW_SR760_FILE.760
            Setpoint Group [1234]~Column 2  or
            Device Information~Column 2 for settings prior to grouped settings
            
            Watch out, the folks at General Electric cannot spell :-(
        """
       
        # main(r'-o xlsx in --new \
        
        # Newer relay application
        main(r'-o xlsx -s "W:\Education\Current\Stationware_Dump\20150430_Stationware_Settings_Applied" --new \
                "Setpoint Group 1~Underfrequency 1 Function" \
                "Setpoint Group 1~Underfrequency 1: Relay 3" \
                "Setpoint Group 1~Underfrequency 1: Relay 4" \
                "Setpoint Group 1~Underfrequency 1: Relay 5" \
                "Setpoint Group 1~Underfrequency 1: Relay 6" \
                "Setpoint Group 1~Underfrequency 1: Relay 7" \
                "Setpoint Group 1~Underfrequency 1 Pickup(Setpoints)" \
                "Setpoint Group 1~Underfrequency 1 Delay" \
                "Setpoint Group 1~Underfrequency 1 Minimum Operating Voltage" \
                "Setpoint Group 1~Underfrequency 1 Minimum Operating Current" \
                "Setpoint Group 1~Underfrequency 2 Function" \
                "Setpoint Group 1~Underfrequency 2: Relay 3" \
                "Setpoint Group 1~Underfrequency 2: Relay 4" \
                "Setpoint Group 1~Underfrequency 2: Relay 5" \
                "Setpoint Group 1~Underfrequency 2: Relay 6" \
                "Setpoint Group 1~Underfrequency 2: Relay 7" \
                "Setpoint Group 1~Underfrequency 2 Pickup(Setpoints)" \
                "Setpoint Group 1~Underfrequency 2 Delay" \
                "Setpoint Group 1~Underfrequency 2 Minimum Opeating Voltage" \
                "Setpoint Group 1~Underfrequency 2 Minimum Operating Current" \
                ')
        
    else:
        main()
    os.system("Pause")
        
