#*- coding: utf-8 -*-
import re
import os
import configparser
import requests
import pymarc
import sys
import json
import time

#Creates the folder if it does not already exist
def check_or_create_dir(path):
    if not os.path.exists(path):
        if not os.path.exists(os.path.dirname(path)):
            check_or_create_dir(os.path.dirname(path))
        os.mkdir(path)

#Folio API configuration file class.
class configuration:
    def __init__(self):
        self.config_folder = os.path.join(os.path.join(os.path.join(os.path.join('C:\\Users', os.getlogin()), 'AppData'), 'Local'), 'FOLIO-api')
        check_or_create_dir(self.config_folder)
        if not os.path.isfile(os.path.join(self.config_folder, 'okapi_manager.ini')):
            self.initial_config()
            self.add_section('data')
            self.config['data']['okapi_url'] = ''
            self.config['data']['tenant'] = ''
            self.config['data']['username'] = ''
            self.config['data']['password'] = ''
            self.config['data']['okapi_token'] = ''
            self.write_config_file()
        else:
            self.initial_config()
    #Create Folio API configuration file.
    def initial_config(self):
        self.config = configparser.ConfigParser()
        self.config_file = os.path.join(self.config_folder, 'okapi_manager.ini')
        self.config_data = {}
        self.read_config_file()
    def read_config_file(self):
        check_or_create_dir(self.config_folder)
        # create a blank file if none exists
        if not os.path.isfile(self.config_file):
            self.write_config_file()
        self.config.read(self.config_file)
    #Updates the configuration file with the current configuration.
    def write_config_file(self):
        with open(self.config_file, 'w') as config_out:
            self.config.write(config_out)
    #Check to see if a section with a given name exist.
    def section_exist(self, section):
        if section in self.config:
            return True
        return False
    #Create a new section if one with a given name does not exist.
    def add_section(self, section):
        if not self.section_exist(section):
            self.config[section] = {}
            self.write_config_file()
    

config = configuration()

def get_okapi_url_from_user():
    okapi_url = input('Okapi url:  ')
    return okapi_url

def get_tenant_from_user():
    tenant = input('Tenant:  ')
    return tenant

def get_username_from_user():
    username = input('Username:  ')
    return username

def get_password_from_user():
    password = input('Password:  ')
    return password

def save_password():
    save = input('Save password (yes/no)?  ')
    save_boolean = False
    if re.match('(ye?s?)', save.lower()):
        save_boolean = True
    return save_boolean

def get_uuid_from_user():
    uuid = input('UUID:  ')
    return uuid

def get_token():
    return config.config['data']['okapi_token']

#Gets new token if one does not exist or existing token is invalid.
def validate_token():
    url = config.config['data']['okapi_url'] + '/instance-storage/instances?limit=0'
    headers = {'Accept' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    req = requests.get(url, headers=headers)
    test_token = req.text
    if test_token == 'Invalid token' or test_token == 'Token missing, access requires permission: inventory-storage.instances.collection.get':
        url = config.config['data']['okapi_url'] + '/authn/login'
        data = '{\"tenant\" : \"' + config.config['data']['tenant'] + '\", \"username\" : \"' + config.config['data']['username'] + '\", \"password\" : \"' + config.config['data']['password'] + '\"}'
        headers = {'Content-type' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant']}
        req = requests.post(url, data=data, headers=headers)
        config.config['data']['okapi_token'] = req.json()['okapiToken']
        config.write_config_file()

#Gets auth information
def auth(okapi_url=None, tenant=None, username=None, password=None, change_okapi_url=False, change_tenant=False, change_username=False, change_password=False, refresh_token=False):
    if okapi_url is not None and tenant is not None and username is not None and password is not None:
        if okapi_url is not None:
            config.config['data']['okapi_url'] = okapi_url
        if tenant is not None:
            config.config['data']['tenant'] = tenant
        if username is not None:
            config.config['data']['username'] = username
        if password is not None:
            config.config['data']['password'] = password
        url = config.config['data']['okapi_url'] + '/authn/login'
        data = '{\"tenant\" : \"' + config.config['data']['tenant'] + '\", \"username\" : \"' + config.config['data']['username'] + '\", \"password\" : \"' + config.config['data']['password'] + '\"}'
        headers = {'Content-type' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant']}
        req = requests.post(url, data=data, headers=headers)
        config.config['data']['okapi_token'] = req.json()['okapiToken']
    elif change_okapi_url or change_tenant or change_tenant or change_username or change_password or refresh_token or'tenant' not in config.config['data'] or 'username' not in config.config['data'] or 'password' not in config.config['data'] or 'okapi_token' not in config.config['data']:
        if change_okapi_url or 'okapi_url' not in config.config['data']:
            config.config['data']['okapi_url'] = get_okapi_url_from_user()
        url = config.config['data']['okapi_url'] + '/authn/login'
        if change_tenant or 'tenant' not in config.config['data']:
            config.config['data']['tenant'] = get_tenant_from_user()
        if change_username or 'username' not in config.config['data']:
            config.config['data']['username'] = get_username_from_user()
        save_pass = False
        if change_password or 'password' not in config.config['data']:
            config.config['data']['password'] = get_password_from_user()
            save_pass = save_password()
        data = '{\"tenant\" : \"' + config.config['data']['tenant'] + '\", \"username\" : \"' + config.config['data']['username'] + '\", \"password\" : \"' + config.config['data']['password'] + '\"}'
        headers = {'Content-type' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant']}
        if refresh_token or 'okapi_token' not in config.config['data']:
            req = requests.post(url, data=data, headers=headers)
            config.config['data']['okapi_token'] = req.json()['okapiToken']
        if not save_pass:
            config.config['data'].pop('password')
        config.write_config_file()
    validate_token()

#Checks auth on start up.
#auth()

validate_token()

#Replaces slash with space.
def replace_slash(text):
    while re.match('([^\\\]*)(\\\)(.*$)', text):
        text = re.match('([^\\\]*)(\\\)(.*$)', text).group(1) + ' ' + re.match('([^\\\]*)(\\\)(.*$)', text).group(3)
    return text

#Converts list to string.
def list_to_string(data):
    output = ''
    if type(data) is list:
        for item in data:
            output = output + item
    else:
        output = data
    return output

#Print completion status to screen.
def print_status(current, last):
    i = current / last
    i = "{:.1%}".format(i)
    sys.stdout.write('\r{0} complete.  Record {1} of {2}'.format(i, current, last))
    sys.stdout.flush()

#Converts Folio 006 into Marc field.
def format_006(data):
    formatted_006 = ''
    #Books
    if data['Type'] in ['t', 'a']:
        formatted_006 = list_to_string(data['Type']) + list_to_string(data['Ills']) + list_to_string(data['Audn']) + list_to_string(data['Form']) + list_to_string(data['Cont']) + list_to_string(data['GPub']) + list_to_string(data['Conf']) + list_to_string(data['Fest']) + list_to_string(data['Indx']) + ' ' + list_to_string(data['LitF']) + list_to_string(data['Biog'])
    #Continuing Resources
    elif data['Type'] == 's':
        formatted_006 = list_to_string(data['Type']) + list_to_string(data['Freq']) + list_to_string(data['Regl']) + ' ' + list_to_string(data['SrTp']) + list_to_string(data['Orig']) + list_to_string(data['Form']) + list_to_string(data['EntW']) + list_to_string(data['Cont']) + list_to_string(data['GPub']) + list_to_string(data['Conf']) + ' ' + ' ' + ' ' + list_to_string(data['Alph']) + list_to_string(data['S/L'])
    #Computer Files
    elif data['Type'] == 'm':
        formatted_006 = list_to_string(data['Type']) + ' ' + ' ' + ' ' + ' ' + list_to_string(data['Audn']) + list_to_string(data['Form']) + ' ' + ' ' + list_to_string(data['File']) + ' ' + list_to_string(data['GPub']) + ' ' + ' ' + ' ' + ' ' + ' ' + ' '
    #Maps
    elif data['Type'] in ['e', 'f']:
        formatted_006 = list_to_string(data['Type']) + list_to_string(data['Relf']) + list_to_string(data['Proj']) + ' ' + list_to_string(data['CrTp']) + ' ' + ' ' + list_to_string(data['GPub']) + list_to_string(data['Form']) + ' ' + list_to_string(data['Indx']) + ' ' + list_to_string(data['SpFm'])
    #Mixed Materials
    elif data['Type'] == 'p':
        formatted_006 = list_to_string(data['Type']) + ' ' + ' ' + ' ' + ' ' + ' ' + list_to_string(data['Form']) + ' ' + ' ' + ' ' + ' ' + ' ' + ' ' + ' ' + ' ' + ' ' + ' ' + ' '
    #Sound Recordings
    elif data['Type'] in ['i', 'j']:
        formatted_006 = list_to_string(data['Type']) + list_to_string(data['Comp']) + list_to_string(data['FMus']) + list_to_string(data['Part']) + list_to_string(data['Audn']) + list_to_string(data['Form']) + list_to_string(data['AccM']) + list_to_string(data['LTxt']) + ' ' + list_to_string(data['TrAr']) + ' '
    #Scores
    elif data['Type'] in ['c', 'd']:
        formatted_006 = list_to_string(data['Type']) + list_to_string(data['Comp']) + list_to_string(data['FMus']) + list_to_string(data['Part']) + list_to_string(data['Audn']) + list_to_string(data['Form']) + list_to_string(data['AccM']) + list_to_string(data['LTxt']) + ' ' + list_to_string(data['TrAr']) + ' '
    #Visual Materials
    elif data['Type'] in ['g', 'k', 'o', 'r']:
        formatted_006 = list_to_string(data['Type']) + list_to_string(data['Time']) + ' ' + list_to_string(data['Audn']) + ' ' + ' ' + ' ' + ' ' + ' ' + list_to_string(data['GPub']) + list_to_string(data['Form']) + ' ' + ' ' + ' ' + list_to_string(data['TMat']) + list_to_string(data['Tech'])
    formatted_006 = replace_slash(formatted_006)
    return formatted_006


#Converts Folio 007 into Marc field.
def format_007(data, uuid):
    formatted_007 = ''
    #Electronic Resource
    if data['$categoryName'].lower() == 'electronic resource':
        try:
            formatted_007 = list_to_string(data['Category']) + list_to_string(data['SMD']) + ' ' + list_to_string(data['Color']) + list_to_string(data['Dimensions']) + list_to_string(data['Sound']) + list_to_string(data['Image bit depth']) + list_to_string(data['File formats']) + list_to_string(data['Quality assurance target(s)']) + list_to_string(data['Antecedent/ Source']) + list_to_string(data['Level of compression']) + list_to_string(data['Reformatting quality'])
        except:
            formatted_007 = list_to_string(data['Category']) + list_to_string(data['SMD']) + ' ' + list_to_string(data['Color']) + list_to_string(data['Dimensions']) + list_to_string(data['Sound'])
    #Globe
    elif data['$categoryName'].lower() == 'globe':
        formatted_007 = list_to_string(data['Category']) + list_to_string(data['SMD']) + ' ' + list_to_string(data['Color']) + list_to_string(data['Physical medium']) + list_to_string(data['Type of reproduction'])
    #Kit
    elif data['$categoryName'].lower() == 'kit':
        formatted_007 = list_to_string(data['Category']) + list_to_string(data['SMD'])
    #Map
    elif data['$categoryName'].lower() == 'map':
        formatted_007 = list_to_string(data['Category']) + list_to_string(data['SMD']) + ' ' + list_to_string(data['Color']) + list_to_string(data['Physical medium']) + list_to_string(data['Type of reproduction']) + list_to_string(data['Production/reproduction details']) + list_to_string(data['Positive/negative aspect'])
    #Microform
    elif data['$categoryName'].lower() == 'microform':
        formatted_007 = list_to_string(data['Category']) + list_to_string(data['SMD']) + ' ' + list_to_string(data['Positive/negative aspect']) + list_to_string(data['Dimensions']) + list_to_string(data['Reduction ratio range/Reduction ratio']) + list_to_string(data['Color']) + list_to_string(data['Emulsion on film']) + list_to_string(data['Generation']) + list_to_string(data['Base of film'])
    #Motion Picture
    elif data['$categoryName'].lower() == 'motion picture':
        formatted_007 = list_to_string(data['Category']) + list_to_string(data['SMD']) + ' ' + list_to_string(data['Color']) + list_to_string(data['Motion picture presentation format']) + list_to_string(data['Sound on medium or separate']) + list_to_string(data['Medium for sound']) + list_to_string(data['Dimensions']) + list_to_string(data['Configuration of playback channels']) + list_to_string(data['Production elements']) + list_to_string(data['Positive/Negative aspect']) + list_to_string(data['Generation']) + list_to_string(data['Base of film']) + list_to_string(data['Refined categories of color']) + list_to_string(data['Kind of color stock or print']) + list_to_string(data['Deterioration stage']) + list_to_string(data['Completeness']) + list_to_string(data['Film inspection date'])
    #Nonprojected Graphic
    elif data['$categoryName'].lower() == 'nonprojected graphic':
        formatted_007 = list_to_string(data['Category']) + list_to_string(data['SMD']) + ' ' + list_to_string(data['Color']) + list_to_string(data['Primary support material']) + list_to_string(data['Secondary support material'])
    #Notated Music
    elif data['$categoryName'].lower() == 'notated music':
        formatted_007 = list_to_string(data['Category']) + list_to_string(data['SMD'])
    #Projected Graphic
    elif data['$categoryName'].lower() == 'projected graphic':
        formatted_007 = list_to_string(data['Category']) + list_to_string(data['SMD']) + ' ' + list_to_string(data['Color']) + list_to_string(data['Color']) + list_to_string(data['Base of emulsion']) + list_to_string(data['Sound on medium or separate']) + list_to_string(data['Medium for sound']) + list_to_string(data['Dimensions']) + list_to_string(data['Secondary support material'])
    #Remote-sensing Image
    elif data['$categoryName'].lower() == 'remote-sensing image':
        formatted_007 = list_to_string(data['Category']) + list_to_string(data['SMD']) + ' ' + list_to_string(data['Altitude of sensor']) + list_to_string(data['Attitude of sensor']) + list_to_string(data['Cloud cover']) + list_to_string(data['Platform construction type']) + list_to_string(data['Platform use category']) + list_to_string(data['Sensor type']) + list_to_string(data['Data type'])
    #Sound Recording
    elif data['$categoryName'].lower() == 'sound recording':
        formatted_007 = list_to_string(data['Category']) + list_to_string(data['SMD']) + ' ' + list_to_string(data['Speed']) + list_to_string(data['Configuration of playback channels']) + list_to_string(data['Groove width/ groove pitch']) + list_to_string(data['Dimensions']) + list_to_string(data['Tape width']) + list_to_string(data['Tape configuration']) + list_to_string(data['Kind of disc, cylinder, or tape']) + list_to_string(data['Kind of material']) + list_to_string(data['Kind of cutting']) + list_to_string(data['Special playback characteristics']) + list_to_string(data['Capture and storage technique'])
    #Tactile Material
    elif data['$categoryName'].lower() == 'tactile material':
        formatted_007 = list_to_string(data['Category']) + list_to_string(data['SMD']) + ' ' + list_to_string(data['Class of braille writing']) + list_to_string(data['Level of contraction']) + list_to_string(data['Braille music format']) + list_to_string(data['Special physical characteristics'])
    #Text
    elif data['$categoryName'].lower() == 'text':
        formatted_007 = list_to_string(data['Category']) + list_to_string(data['SMD'])
    #Unspecified
    elif data['$categoryName'].lower() == 'unspecified':
        formatted_007 = list_to_string(data['Category']) + list_to_string(data['SMD'])
    #Videorecording
    elif data['$categoryName'].lower() == 'videorecording':
        formatted_007 = list_to_string(data['Category']) + list_to_string(data['SMD']) + ' ' + list_to_string(data['Color']) + list_to_string(data['Videorecording format']) + list_to_string(data['Sound on medium or separate']) + list_to_string(data['Medium for sound']) + list_to_string(data['Dimensions']) + list_to_string(data['Configuration of playback channels'])
    formatted_007 = replace_slash(formatted_007)
    return formatted_007

#Converts Folio 008 into Marc field.
def format_008(data):
    formatted_008 = ''
    field_008_start = list_to_string(data['Entered']) + list_to_string(data['DtSt']) + list_to_string(data['Date1']) + list_to_string(data['Date2']) + list_to_string(data['Ctry'])
    field_008_end = list_to_string(data['Lang']) + list_to_string(data['MRec']) + list_to_string(data['Srce'])
    #Books
    if data['Type'] == 't' or (data['Type'] == 'a' and data['BLvl'] in ['a', 'c', 'd', 'm']):
        field_008_middle = list_to_string(data['Ills']) + list_to_string(data['Audn']) + list_to_string(data['Form']) + list_to_string(data['Cont']) + list_to_string(data['GPub']) + list_to_string(data['Conf']) + list_to_string(data['Fest']) + list_to_string(data['Indx']) + ' ' + list_to_string(data['LitF']) + list_to_string(data['Biog'])
    #Continuing Resources
    elif data['Type'] == 'a' and data['BLvl'] in ['b', 'i', 's']:
        field_008_middle = list_to_string(data['Freq']) + list_to_string(data['Regl']) + ' ' + list_to_string(data['SrTp']) + list_to_string(data['Orig']) + list_to_string(data['Form']) + list_to_string(data['EntW']) + list_to_string(data['Cont']) + list_to_string(data['GPub']) + list_to_string(data['Conf']) + ' ' + ' ' + ' ' + list_to_string(data['Alph']) + list_to_string(data['S/L'])
    #Computer Files
    elif data['Type'] == 'm':
        field_008_middle = ' ' + ' ' + ' ' + ' ' + list_to_string(data['Audn']) + list_to_string(data['Form']) + ' ' + ' ' + list_to_string(data['File']) + ' ' + list_to_string(data['GPub']) + ' ' + ' ' + ' ' + ' ' + ' ' + ' '
    #Maps
    elif data['Type'] == 'e' or data['Type'] == 'f':
        field_008_middle = list_to_string(data['Relf']) + list_to_string(data['Proj']) + ' ' + list_to_string(data['CrTp']) + ' ' + ' ' + list_to_string(data['GPub']) + list_to_string(data['Form']) + ' ' + list_to_string(data['Indx']) + ' ' + list_to_string(data['SpFm'])
    #Mixed Materials
    elif data['Type'] == 'p':
        field_008_middle = ' ' + ' ' + ' ' + ' ' + ' ' + list_to_string(data['Form']) + ' ' + ' ' + ' ' + ' ' + ' ' + ' ' + ' ' + ' ' + ' ' + ' ' + ' '
    #Sound Recordings
    elif data['Type'] == 'i' or data['Type'] == 'j':
        field_008_middle = list_to_string(data['Comp']) + list_to_string(data['FMus']) + list_to_string(data['Part']) + list_to_string(data['Audn']) + list_to_string(data['Form']) + list_to_string(data['AccM']) + list_to_string(data['LTxt']) + ' ' + list_to_string(data['TrAr']) + ' '
    #Scores
    elif data['Type'] == 'c' or data['Type'] == 'd':
        field_008_middle = list_to_string(data['Comp']) + list_to_string(data['FMus']) + list_to_string(data['Part']) + list_to_string(data['Audn']) + list_to_string(data['Form']) + list_to_string(data['AccM']) + list_to_string(data['LTxt']) + ' ' + list_to_string(data['TrAr']) + ' '
    #Visual Materials
    elif data['Type'] == 'g' or data['Type'] == 'k' or data['Type'] == 'o' or data['Type'] == 'r':
        field_008_middle = list_to_string(data['Time']) + ' ' + list_to_string(data['Audn']) + ' ' + ' ' + ' ' + ' ' + ' ' + list_to_string(data['GPub']) + list_to_string(data['Form']) + ' ' + ' ' + ' ' + list_to_string(data['TMat']) + list_to_string(data['Tech'])
    formatted_008 = replace_slash(field_008_start + field_008_middle + field_008_end)
    return formatted_008

#Gets marc record from API.
#Converts record from json to marc then returns record.
def get_marc(uuid = None):
    if uuid is None:
        uuid = get_uuid_from_user()
    headers = {'Accept' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    url = config.config['data']['okapi_url'] + '/records-editor/records?externalId=' + uuid
    req = requests.get(url, headers=headers)
    json_record = req.json()
    record = pymarc.Record()
    record.leader = replace_slash(json_record['leader'])
    row_006 = None
    for row in json_record['fields']:
        tag = row['tag']
        if int(row['tag']) < 10:
            if int(row['tag']) in [6, 7, 8]:
                if int(row['tag']) == 6:
                    row['content'] = format_006(row['content'])
                if int(row['tag']) == 7:
                    row['content'] = format_007(row['content'], uuid)
                if int(row['tag']) == 8:
                    row['content'] = format_008(row['content'])
            record.add_field(pymarc.Field(row['tag'], data=row['content']))
        else:
            subfields = []
            for match in re.finditer('(?:\$)([^\$])([^\$]+)', row['content']):
                subfields.append(match.group(1))
                subfields.append(match.group(2))
            record.add_field(pymarc.Field(row['tag'], row['indicators'], subfields))
    return record

#Returns all the instance records.
def get_instance_records_all(start = 0, end = None, limit = 100000, return_type = 'json'):
    headers = {'Accept' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    url = config.config['data']['okapi_url'] + '/instance-storage/instances'
    a = start
    if end is None:
        params = {'limit' : '1'}
        req = requests.get(url, params=params, headers=headers)
        json_record = req.json()
        end = json_record['totalRecords']
    while start < end:
        params = {'offset' : str(start), 'limit' : str(limit)}
        req = requests.get(url, params=params, headers=headers)
        json_record = req.json()
        for instance_record in json_record['instances']:
            if return_type == 'text':
                yield instance_record.text
            else:
                yield instance_record
            a += 1
            print_status(a, end)
        start += limit


#Gets marc record from API via oclc number.
def get_marc_record_from_oclc(oclc = None, return_type = 'json'):
    headers = {'Accept' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    if oclc != '' and oclc is not None:
        query = 'oclc="' + str(oclc) + '"'
        url = config.config['data']['okapi_url'] + '/search/instances?query=' + query + '&limit=1'
        req = requests.get(url, headers=headers)
        returned_instances = req.json()
        if returned_instances['instances'] != []:
            return get_marc(returned_instances['instances'][0]['id'])
        return None

#Gets instance_record.
def get_instance_record(uuid = None, return_type = 'json'):
    if uuid is None:
        uuid = get_uuid_from_user()
    headers = {'Accept' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    if uuid != '' and uuid is not None:
        url = config.config['data']['okapi_url'] + '/instance-storage/instances/' + uuid
        req = requests.get(url, headers=headers)
    else:
        url = config.config['data']['okapi_url'] + '/instance-storage/instances?limit=0'
        req = requests.get(url, headers=headers)
    if return_type == 'text':
        return req.text
    elif return_type == 'json':
        return req.json()
    else:
        return None

#Gets instance_record.
def get_location_from_location_id(uuid = None, return_type = 'json'):
    if uuid is None:
        uuid = get_uuid_from_user()
    headers = {'Accept' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    if uuid != '' and uuid is not None:
        url = config.config['data']['okapi_url'] + '/locations/' + uuid
        req = requests.get(url, headers=headers)
    else:
        url = config.config['data']['okapi_url'] + '/locations?limit=0'
        req = requests.get(url, headers=headers)
    if return_type == 'json':
        return req.json()['name']
    else:
        return None

#Gets item records from holdings_id.
def get_item_records_from_holdings_id(uuid = None, return_type = 'json'):
    if uuid is None:
        uuid = get_uuid_from_user()
    headers = {'Accept' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    if uuid != '' and uuid is not None:
        url = config.config['data']['okapi_url'] + '/item-storage/items?limit=10000&query=holdingsRecordId==' + uuid
        req = requests.get(url, headers=headers)
    else:
        url = config.config['data']['okapi_url'] + '/item-storage/items?limit=0'
        req = requests.get(url, headers=headers)
    if return_type == 'text':
        return req.text
    elif return_type == 'json':
        return req.json()
    else:
        return None

#Gets holding records from instance_id.
def get_holdings_records_from_instance_id(uuid = None, return_type = 'json'):
    if uuid is None:
        uuid = get_uuid_from_user()
    headers = {'Accept' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    if uuid != '' and uuid is not None:
        url = config.config['data']['okapi_url'] + '/holdings-storage/holdings?limit=10000&query=instanceId==' + uuid
        req = requests.get(url, headers=headers)
    else:
        url = config.config['data']['okapi_url'] + '/holdings-storage/holdings?limit=0'
        req = requests.get(url, headers=headers)
    if return_type == 'text':
        return req.text
    elif return_type == 'json':
        return req.json()
    else:
        return None

#Gets item records from instance_id.
def get_item_records_from_instance_id(uuid = None, return_type = 'json'):
    if uuid is None:
        uuid = get_uuid_from_user()
    headers = {'Accept' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    if uuid != '' and uuid is not None:
        url = config.config['data']['okapi_url'] + '/item-storage/items?limit=10000&query=instanceId==' + uuid
        req = requests.get(url, headers=headers)
    else:
        url = config.config['data']['okapi_url'] + '/item-storage/items?limit=0'
        req = requests.get(url, headers=headers)
    if return_type == 'text':
        return req.text
    else:
        return req.json()

#Gets holding record.
def get_holdings_record(uuid = None, return_type = 'json'):
    if uuid is None:
        uuid = get_uuid_from_user()
    while re.match('(.*)(?:[^a-f0-9\-]+)(.*$)', uuid):
        uuid = re.match('(.*)(?:[^a-f0-9\-]+)(.*$)', uuid).group(1) + re.match('(.*)(?:[^a-f0-9\-]+)(.*$)', uuid).group(2)
    headers = {'Accept' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    if uuid != '' and uuid is not None:
        url = config.config['data']['okapi_url'] + '/holdings-storage/holdings/' + uuid
        req = requests.get(url, headers=headers)
    else:
        url = config.config['data']['okapi_url'] + '/holdings-storage/holdings?limit=0'
        params = {'limit': '0'}
        req = requests.get(url, params=params, headers=headers)
    if return_type == 'text':
        return req.text
    else:
        return req.json()

#Get suppress status.
def get_suppress_status(uuid):
    headers = {'Content-type' : 'application/json', 'origin' : 'https://crl.folio.ebsco.com', 'referer' : 'https://crl.folio.ebsco.com/', 'user-agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36 Edg/105.0.1343.33', 'Accept' : 'text/plain', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    url = config.config['data']['okapi_url'] + '/instance-storage/instances/' + uuid
    req = requests.get(url, headers=headers)
    json_record = req.json()
    discoverySuppress = False
    staffSuppress = False
    if 'discoverySuppress' in json_record and json_record['discoverySuppress'] is not None:
        discoverySuppress = json_record['discoverySuppress']
    if 'staffSuppress' in json_record and json_record['staffSuppress'] is not None:
        staffSuppress = json_record['staffSuppress']
    return {'discoverySuppress' : discoverySuppress, 'staffSuppress' : staffSuppress}

#Suppresses record.
def suppress_record(uuid = None):
    if uuid is None:
        uuid = get_uuid_from_user()
    while re.match('(.*)(?:[^a-f0-9\-]+)(.*$)', uuid):
        uuid = re.match('(.*)(?:[^a-f0-9\-]+)(.*$)', uuid).group(1) + re.match('(.*)(?:[^a-f0-9\-]+)(.*$)', uuid).group(2)
    headers = {'Content-type' : 'application/json', 'origin' : 'https://crl.folio.ebsco.com', 'referer' : 'https://crl.folio.ebsco.com/', 'user-agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36 Edg/105.0.1343.33', 'Accept' : 'text/plain', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    if uuid != '' and uuid is not None:
        url = config.config['data']['okapi_url'] + '/instance-storage/instances/' + uuid
        req = requests.get(url, headers=headers)
        json_record = req.json()
        json_record['discoverySuppress'] = True
        json_record['staffSuppress'] = True
        data = json.dumps(json_record)
        response = requests.put(url, headers=headers, data=data)
        print(uuid + ' suppressed')

#Gets all the holding record ids.
def get_holdings_record_ids(start = 0, limit = 100000, return_type = 'json'):
    headers = {'Accept' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    url = config.config['data']['okapi_url'] + '/holdings-storage/holdings'
    a = start
    params = {'limit' : '1'}
    req = requests.get(url, params=params, headers=headers)
    json_record = req.json()
    end = json_record['totalRecords']
    while start <= end:
        params = {'offset' : str(start), 'limit' : str(limit)}
        req = requests.get(url, params=params, headers=headers)
        json_record = req.json()
        for holdings_record in json_record['holdingsRecords']:
            yield holdings_record['id']
            a += 1
            print_status(a, end)
        start += limit

#Gets the item record.
def get_item_record(uuid = None, return_type = 'json'):
    if uuid is None:
        uuid = get_uuid_from_user()
    headers = {'Accept' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    if uuid != '' and uuid is not None:
        url = config.config['data']['okapi_url'] + '/item-storage/items/' + uuid
        req = requests.get(url, headers=headers)
    else:
        url = config.config['data']['okapi_url'] + '/item-storage/items?limit=0'
        req = requests.get(url, headers=headers)
    if return_type == 'text':
        return req.text
    elif return_type == 'json':
        return req.json()
    else:
        return None

#Updates item record.
def update_item_record(uuid, json_record):
    while re.match('(.*)(?:[^a-f0-9\-]+)(.*$)', uuid):
        uuid = re.match('(.*)(?:[^a-f0-9\-]+)(.*$)', uuid).group(1) + re.match('(.*)(?:[^a-f0-9\-]+)(.*$)', uuid).group(2)
    headers = {'Content-type' : 'application/json', 'origin' : 'https://crl.folio.ebsco.com', 'referer' : 'https://crl.folio.ebsco.com/', 'user-agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36 Edg/105.0.1343.33', 'Accept' : 'text/plain', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    if uuid != '' and uuid is not None:
        url = config.config['data']['okapi_url'] + '/item-storage/items/' + uuid
        data = json.dumps(json_record)
        response = requests.put(url, headers=headers, data=data)

#Updates item record volume.
def update_item_record_volume(uuid, user_json_record):
    while re.match('(.*)(?:[^a-f0-9\-]+)(.*$)', uuid):
        uuid = re.match('(.*)(?:[^a-f0-9\-]+)(.*$)', uuid).group(1) + re.match('(.*)(?:[^a-f0-9\-]+)(.*$)', uuid).group(2)
    headers = {'Content-type' : 'application/json', 'origin' : 'https://crl.folio.ebsco.com', 'referer' : 'https://crl.folio.ebsco.com/', 'user-agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36 Edg/105.0.1343.33', 'Accept' : 'text/plain', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    if uuid != '' and uuid is not None:
        url = config.config['data']['okapi_url'] + '/item-storage/items/' + uuid
        req = requests.get(url, headers=headers)
        json_record = req.json()
        json_record['volume'] = user_json_record['volume']
        data = json.dumps(json_record)
        response = requests.put(url, headers=headers, data=data)

#Returns all the item records.
def get_item_records_all(start = 0, end = None, limit = 100000, return_type = 'json'):
    headers = {'Accept' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    url = config.config['data']['okapi_url'] + '/item-storage/items'
    a = start
    if end is None:
        params = {'limit' : '1'}
        req = requests.get(url, params=params, headers=headers)
        json_record = req.json()
        end = json_record['totalRecords']
    while start < end:
        params = {'offset' : str(start), 'limit' : str(limit)}
        req = requests.get(url, params=params, headers=headers)
        json_record = req.json()
        for item_record in json_record['items']:
            if return_type == 'text':
                yield item_record.text
            else:
                yield item_record
            a += 1
            print_status(a, end)
        start += limit

#Returns all the holdings records.
def get_holdings_records_all(start = 0, end = None, limit = 100000, return_type = 'json'):
    headers = {'Accept' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    url = config.config['data']['okapi_url'] + '/holdings-storage/holdings'
    a = start
    if end is None:
        params = {'limit' : '1'}
        req = requests.get(url, params=params, headers=headers)
        json_record = req.json()
        end = json_record['totalRecords']
    while start < end:
        params = {'offset' : str(start), 'limit' : str(limit)}
        req = requests.get(url, params=params, headers=headers)
        json_record = req.json()
        for item_record in json_record['holdingsRecords']:
            if return_type == 'text':
                yield item_record.text
            else:
                yield item_record
            a += 1
            print_status(a, end)
        start += limit

#Returns all the authority records.
def search_authority_records_all(start = 0, end = None, limit = 500, return_type = 'json'):
    headers = {'Accept' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    url = config.config['data']['okapi_url'] + '/search/authorities'
    a = start
    if limit > 500:
        limit = 500
    if end is None:
        params = {'limit' : '1', 'query' : 'keyword=*'}
        req = requests.get(url, params=params, headers=headers)
        json_record = req.json()
        end = json_record['totalRecords']
    while start < end:
        params = {'offset' : str(start), 'limit' : str(limit), 'query' : 'keyword=*'}
        req = requests.get(url, params=params, headers=headers)
        json_record = req.json()
        for item_record in json_record['authorities']:
            if return_type == 'text':
                yield item_record.text
            else:
                yield item_record
            a += 1
            print_status(a, end)
        start += limit

#Returns all the authority records.
def get_authority_records_all(start = 0, end = None, limit = 500, return_type = 'json'):
    headers = {'Accept' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
#    url = config.config['data']['okapi_url'] + '/search/authorities'
    url = config.config['data']['okapi_url'] + '/authority-storage/authorities'
    a = start
    if limit > 500:
        limit = 500
    if end is None:
        params = {'limit' : '1', 'query' : 'keyword=*'}
        req = requests.get(url, params=params, headers=headers)
        json_record = req.json()
        end = json_record['totalRecords']
    while start < end:
        params = {'offset' : str(start), 'limit' : str(limit), 'query' : 'keyword=*'}
        req = requests.get(url, params=params, headers=headers)
        json_record = req.json()
        for item_record in json_record['authorities']:
            if return_type == 'text':
                yield item_record.text
            else:
                yield item_record
            a += 1
            print_status(a, end)
        start += limit

#Returns all the instance records.
def get_instance_relationships_all(start = 0, end = None, limit = 100000, return_type = 'json'):
    headers = {'Accept' : 'application/json', 'X-Okapi-Tenant' : config.config['data']['tenant'], 'x-okapi-token' : config.config['data']['okapi_token']}
    url = config.config['data']['okapi_url'] + '/instance-storage/instance-relationships'
    a = start
    if end is None:
        params = {'limit' : '1'}
        req = requests.get(url, params=params, headers=headers)
        json_record = req.json()
        end = json_record['totalRecords']
    while start < end:
        params = {'offset' : str(start), 'limit' : str(limit)}
        req = requests.get(url, params=params, headers=headers)
        json_record = req.json()
        for instance_record in json_record['instanceRelationships']:
            if return_type == 'text':
                yield instance_record.text
            else:
                yield instance_record
            a += 1
            print_status(a, end)
        start += limit