#!/usr/bin/env python3

import sys
import os
import pprint
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime

#### Variables #####################################################################
#Timestamp
today = datetime.now()
timestamp = today.strftime("%Y%m%d_%H.%M")

#Working Directory. The directory the script lives in
workingdir = sys.path[0]

# Log directory and file creation
logdir = os.path.join(workingdir, "logs")
if not os.path.exists(logdir):
    os.makedirs(logdir)
logfile = os.path.join(logdir, timestamp)

# Read the excel file
xlfile = r"C:\Users\jholtz\Downloads\Usernames first group to kim.xlsx"
xlread = pd.read_excel(xlfile)

# Test XML file
#xml_file = r"C:\temp\PrsnlExport2020-06-10@1202(00001).xml"
####################################################################################

#### Functions #####################################################################
#Creates a prnl element and appends it to the prsnl_list
def add_prsnl(status, person_id, username, lastname, firstname, middle_name, position, begin_date, physician_ind, data_status, end_date):
    # Adding subtags inside our PRSNL tag
    element_prsnl = ET.Element('PRSNL')
    # Add subtag under the "PRSNL" 
    element_prsnl_status = ET.SubElement(element_prsnl, 'STATUS')
    element_prsnl_external_id = ET.SubElement(element_prsnl, 'EXTERNAL_ID')
    if person_id != 'nan':
        element_prsnl_person_id = ET.SubElement(element_prsnl, 'PERSON_ID')
        element_prsnl_person_id.text = person_id
    element_prsnl_username = ET.SubElement(element_prsnl, 'USERNAME')
    element_prsnl_lastname = ET.SubElement(element_prsnl, 'LAST_NAME')
    element_prsnl_firstname = ET.SubElement(element_prsnl, 'FIRST_NAME')
    if middle_name != 'nan':
        element_prsnl_middle_name = ET.SubElement(element_prsnl, 'MIDDLE_NAME')
        element_prsnl_middle_name.text = middle_name
    element_prsnl_position = ET.SubElement(element_prsnl, 'POSITION')
    element_prsnl_begin_date = ET.SubElement(element_prsnl, 'BEG_DT_TM')
    element_prsnl_physician_ind = ET.SubElement(element_prsnl, 'PHYSICIAN_IND')
    element_prsnl_data_status = ET.SubElement(element_prsnl, 'DATA_STATUS')
    element_prsnl_end_date = ET.SubElement(element_prsnl, 'END_DT_TM')
    element_prsnl_active_ind = ET.SubElement(element_prsnl, 'ACTIVE_IND')

    # Add the text to the subelements in PRSNL (Place the variables from the excel file here)
    element_prsnl_status.text = status
    element_prsnl_external_id.text = ""
    element_prsnl_username.text = username
    element_prsnl_lastname.text = lastname
    element_prsnl_firstname.text = firstname    
    element_prsnl_position.text = position
    element_prsnl_begin_date.text = begin_date
    element_prsnl_physician_ind.text = physician_ind
    element_prsnl_data_status.text = data_status
    element_prsnl_end_date.text = end_date
    element_prsnl_active_ind.text = '1'
    # Append the prsl list
    element_prsnl_list.append(element_prsnl)
    ## Org Group list creation
    element_prsnl_org_group_list = ET.SubElement(element_prsnl, 'ORG_GROUP_LIST')
    #### Org Group List ###############################################################
    ## Individual org creation within the list (add some kind of variable list later on)
    add_org_group(group_id='54463', group_type='SECURITY', group_name='Physician Network Services', group_list=element_prsnl_org_group_list)
    add_org_group(group_id='54465', group_type='SECURITY', group_name='Texas Tech Physicians', group_list=element_prsnl_org_group_list)
    add_org_group(group_id='1629881', group_type='SECURITY', group_name='UMC', group_list=element_prsnl_org_group_list)

# Creates an org and adds it to the org group
def add_org_group(group_id, group_type, group_name, group_list):
    element_prsnl_org_group_list_org_group = ET.Element('ORG_GROUP')
    element_prsnl_org_group_list_org_group_org_group_id = ET.SubElement(element_prsnl_org_group_list_org_group, 'ORG_GROUP_ID')
    element_prsnl_org_group_list_org_group_org_group_type = ET.SubElement(element_prsnl_org_group_list_org_group, 'ORG_GROUP_TYPE')
    element_prsnl_org_group_list_org_group_org_group_name = ET.SubElement(element_prsnl_org_group_list_org_group, 'ORG_GROUP_NAME')
    #element_prsnl_org_group_list_org_group_org_group_delete_ind = ET.SubElement(element_prsnl_org_group_list_org_group, 'ORG_GROUP_DELETE_IND')

    # Add the text to the subelements in PRSNL (Place the variables from the excel file here)
    element_prsnl_org_group_list_org_group_org_group_id.text = group_id
    element_prsnl_org_group_list_org_group_org_group_type.text = group_type
    element_prsnl_org_group_list_org_group_org_group_name.text = group_name
    #element_prsnl_org_group_list_org_group_org_group_delete_ind.text = '0'

    # Append per itteration incase there are multiples
    group_list.append(element_prsnl_org_group_list_org_group)

###################################################################################

# Setup the xml file header/type
with open(r'C:\temp\PrsnlImport.xml', 'w') as f:
    f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
f.close()

# Setup the parent (root) tag onto which other tages would be created
data = ET.Element('DOMAIN_LIST')
# Add domain attributes to the domain list
data.set('hostname', 'HB115')
data.set('ImportAvailable', '1')
data.text = ''
# Adding subtags inside our root tag
element_prsnl_domain = ET.SubElement(data, 'PRSNL_DOMAIN')
# Add attributes to the prsnl_domain
element_prsnl_domain.set('hnam_imp_key', 'PRSNL')
# Setup prsnl_list tree
element_prsnl_list = ET.SubElement(element_prsnl_domain, 'PRSNL_LIST')

#### PRSNL Section ################################################################
for i in range(len(xlread)):
    status = '0'
    username = str(xlread.username[i])
    lastname = str(xlread.lastname[i])
    firstname =str( xlread.firstname[i])
    middle_name = str(xlread.middle_name[i])
    position = str(xlread.position[i])
    begin_date = str(xlread.begin_date[i])
    end_date = str(xlread.end_date[i])
    physician_ind = '1'
    data_status = 'Auth (Verified)'
    person_id = str(xlread.person_id[i]).split('.')[0]
    add_prsnl(status=status, username=username, person_id=person_id, lastname=lastname, firstname=firstname, middle_name=middle_name, position=position, begin_date=begin_date, physician_ind=physician_ind, data_status=data_status, end_date=end_date)


# Convert the data to a string
b_xml = ET.tostring(data, short_empty_elements=False)

# Write to a file with the "write/binary" setting
with open(r'C:\temp\PrsnlImport.xml', 'ab') as f:
    f.write(b_xml)
f.close()
