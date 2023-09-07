#*- coding: utf-8 -*-
import re
import io
import pymarc
from bs4 import BeautifulSoup
import bs4
import lxml

#Converts text marc record from Worldcat to pymarc record.
def get_marc_worldcat(marc_record):
    record = None
    #If marc_record is a sting, creates generator with re.finditer otherwise uses given generator.
    if type(marc_record) is str:
        marc_text_generator = re.finditer('(^.*)(?:\n|$)', marc_record, re.MULTILINE)
    else:
        marc_text_generator = marc_record
    a = 1
    #Create pymarc record from text marc record
    for line0 in marc_text_generator:
        if type(marc_record) is str:
            line = line0.group(1)
        else:
            line = line0
        tag = ''
        indicators = []
        subfields = []
        entry_data = ''
        if re.match('(?:\=LDR\s*|\=LEADER\s*)(.+)', line):
            if record is not None:
                yield record
                a += 1
            record = pymarc.Record()
            record.leader = re.match('(?:\=LDR\s*|\=LEADER\s*)(.+)', line).group(1)
        elif re.match('(?:\=)(\d{3})', line):
            tag = re.search('(?:\=)(\d{3})', line).group(1)
            if int(re.search('(?:\=)(\d{3})', line).group(1)) < 10:
                if re.search('(?:\=\d{3}\s+)(.+$)', line):
                    entry_data = re.search('(?:\=\d{3}\s+)(.+$)', line).group(1)
                record.add_field(pymarc.Field(tag, data=entry_data))
            else:
                if re.search('(?:\=\d{3}\s+)(.)(.)(.+$)', line):
                    if re.match('(?:\s$)', re.search('(?:\=\d{3}\s+)(.)(.)(.+$)', line).group(1)):
                        indicators.append('\\')
                    else:
                        indicators.append(re.search('(?:\=\d{3}\s+)(.)(.)(.+$)', line).group(1))
                    if re.match('(?:\s$)', re.search('(?:\=\d{3}\s+)(.)(.)(.+$)', line).group(2)):
                        indicators.append('\\')
                    else:
                        indicators.append(re.search('(?:\=\d{3}\s+)(.)(.)(.+$)', line).group(2))
                    base_subfields = re.search('(?:\=\d{3}\s+)(.)(.)(.+$)', line).group(3)
                    for m in re.finditer('(?:\$)([^\$])([^\$]+)', base_subfields):
                        subfields.append(m.group(1))
                        subfields.append(m.group(2))
                record.add_field(pymarc.Field(tag, indicators, subfields))
    yield record

#Converts text marc record from Millennium to pymarc record.
def get_marc_millennium(marc_record):
    record = None
    #If marc_record is a sting, creates generator with re.finditer otherwise uses given generator.
    if type(marc_record) is str:
        marc_record_temp = ''
        #Removes excess newline characters.
        while re.match('(.*)(?:\n{2})(.*$)', marc_record, re.MULTILINE):
            marc_record = re.match('(.*)(?:\n{2})(.*$)', marc_record, re.MULTILINE).group(1) + '\n' + re.match('(.*)(?:\n{2})(.*$)', marc_record, re.MULTILINE).group(2)
        #Removes newlines caused by web display.
        while re.match('(.*)(?:\n\s{2}\s+)(.*$)', marc_record, re.MULTILINE):
            marc_record = re.match('(.*)(?:\n\s{2}\s+)(.*$)', marc_record, re.MULTILINE).group(1) + re.match('(.*)(?:\n\s{2}\s+)(.*$)', marc_record, re.MULTILINE).group(2)
        for line0 in re.finditer('(^.*)(?:\n|$)', marc_record, re.MULTILINE):
            if marc_record_temp != '' and re.search('(\d{3}\s[\s\d]{2}\s|LDR|LEADER)', line0.group(1)):
                marc_record_temp = marc_record_temp + '\n'
            if re.match('(?:\s+)(\S.*$)', line0.group(1)):
                marc_record_temp = marc_record_temp + re.match('(?:\s+)(\S.*$)', line0.group(1)).group(1)
            else:
                marc_record_temp = marc_record_temp + line0.group(1)
        marc_text_generator = re.finditer('(^.*)(?:\n|$)', marc_record, re.MULTILINE)
    else:
        marc_text_generator = marc_record
    a = 1
    for line0 in marc_text_generator:
        if type(marc_record) is str:
            line = line0.group(1)
        else:
            line = line0
        tag = ''
        indicators = []
        subfields = []
        entry_data = ''
        if re.match('(?:LDR\s*|LEADER\s*)(.+)', line):
            if record is not None:
                yield record
            record = pymarc.Record()
            record.leader = re.match('(?:LDR\s*|LEADER\s*)(.+)', line).group(1)
        elif re.match('(\d{3})', line):
            tag = re.search('(\d{3})', line).group(1)
            if int(re.search('(\d{3})', line).group(1)) < 10:
                if re.search('(?:\d{3}\s+)(.+$)', line):
                    entry_data = re.search('(?:\d{3}\s+)(.+$)', line).group(1)
                record.add_field(pymarc.Field(tag, data=entry_data))
            else:
                if re.search('(?:\d{3}\s)(.)(.)(?:\s)(.+$)', line):
                    if re.match('(?:\s$)', re.search('(?:\d{3}\s)(.)(.)(?:\s)(.+$)', line).group(1)):
                        indicators.append('\\')
                    else:
                        indicators.append(re.search('(?:\d{3}\s)(.)(.)(?:\s)(.+$)', line).group(1))
                    if re.match('(?:\s$)', re.search('(?:\d{3}\s)(.)(.)(?:\s)(.+$)', line).group(2)):
                        indicators.append('\\')
                    else:
                        indicators.append(re.search('(?:\d{3}\s)(.)(.)(?:\s)(.+$)', line).group(2))
                    base_subfields = re.search('(?:\d{3}\s)(.)(.)(?:\s)(.+$)', line).group(3)
                    if not re.match('(?:\|)', base_subfields):
                        base_subfields = '|a' + base_subfields
                    for m in re.finditer('(?:\|)([^\|])([^\|]+)', base_subfields):
                        subfields.append(m.group(1))
                        subfields.append(m.group(2))
                record.add_field(pymarc.Field(tag, indicators, subfields))
    yield record

#Converts text marc record from ISSN database to pymarc record.
def get_marc_issn(marc_record):
    record = None
    #If marc_record is a sting, creates generator with re.finditer otherwise uses given generator.
    if type(marc_record) is str:
        marc_record_temp = ''
        #Removes excess newline characters.
        while re.match('(.*)(?:\n{2})(.*$)', marc_record, re.MULTILINE):
            marc_record = re.match('(.*)(?:\n{2})(.*$)', marc_record, re.MULTILINE).group(1) + '\n' + re.match('(.*)(?:\n{2})(.*$)', marc_record, re.MULTILINE).group(2)
        while re.match('(.*)(?:\n\s{2}\s+)(.*$)', marc_record, re.MULTILINE):
            marc_record = re.match('(.*)(?:\n\s{2}\s+)(.*$)', marc_record, re.MULTILINE).group(1) + re.match('(.*)(?:\n\s{2}\s+)(.*$)', marc_record, re.MULTILINE).group(2)
        for line0 in re.finditer('(^.*)(?:\n|$)', marc_record, re.MULTILINE):
            if marc_record_temp != '' and re.search('(\=\d{3}\s{2}|\=LDR|\=LEADER)', line0.group(1)):
                marc_record_temp = marc_record_temp + '\n'
            if re.match('(?:\s+)(\S.*$)', line0.group(1)):
                marc_record_temp = marc_record_temp + re.match('(?:\s+)(\S.*$)', line0.group(1)).group(1)
            else:
                marc_record_temp = marc_record_temp + line0.group(1)
        marc_text_generator = re.finditer('(^.*)(?:\n|$)', marc_record, re.MULTILINE)
    else:
        marc_text_generator = marc_record
    for line0 in marc_text_generator:
        if type(marc_record) is str:
            line = line0.group(1)
        else:
            line = line0
        tag = ''
        indicators = []
        subfields = []
        entry_data = ''
        if re.match('(?:\=LDR\s*|\=LEADER\s*)(.+)', line):
            if record is not None:
                yield record
            record = pymarc.Record()
            record.leader = re.match('(?:\=LDR\s*|\=LEADER\s*)(.+)', line).group(1)
        elif re.match('(?:\=)(\d{3})', line):
            tag = re.search('(?:\=)(\d{3})', line).group(1)
            if int(re.search('(?:\=)(\d{3})', line).group(1)) < 10:
                if re.search('(?:\=\d{3}\s+)(.+$)', line):
                    entry_data = re.search('(?:\=\d{3}\s+)(.+$)', line).group(1)
                record.add_field(pymarc.Field(tag, data=entry_data))
            else:
                if re.search('(?:\=\d{3}\s{2})(.)(.)(.+$)', line):
                    if re.match('(?:\s$)', re.search('(?:\=\d{3}\s{2})(.)(.)(.+$)', line).group(1)):
                        indicators.append('\\')
                    else:
                        indicators.append(re.search('(?:\=\d{3}\s{2})(.)(.)(.+$)', line).group(1))
                    if re.match('(?:\s$)', re.search('(?:\=\d{3}\s{2})(.)(.)(.+$)', line).group(2)):
                        indicators.append('\\')
                    else:
                        indicators.append(re.search('(?:\=\d{3}\s{2})(.)(.)(.+$)', line).group(2))
                    base_subfields = re.search('(?:\=\d{3}\s{2})(.)(.)(.+$)', line).group(3)
                    if not re.match('(?:\$)', base_subfields):
                        base_subfields = '$a' + base_subfields
                    for m in re.finditer('(?:\$)([^\$])([^\$]+)', base_subfields):
                        subfields.append(m.group(1))
                        subfields.append(m.group(2))
                record.add_field(pymarc.Field(tag, indicators, subfields))
    yield record

#Converts text marc record from Folio to pymarc record.
def get_marc_folio(marc_record):
    record = None
    if type(marc_record) is bs4.element.Tag:
        marc_text_generator = marc_record.find_all('tr')
    else:
        marc_text_generator = marc_record
    a = 1
    for line in marc_text_generator:
        tag = ''
        indicators = [None, None]
        subfields = []
        entry_data = ''
        if 'class' in list(line.attrs):
            if line['class'][0] == 'marc-row-LEADER':
                if record is not None:
                    yield record
                record = pymarc.Record()
                record.leader = line.find('td').get_text()
            elif re.match('(marc-row-)', line['class'][0]):
                tag = line.find('th').get_text()
                if int(tag) < 10:
                    entry_data = line.find('td').get_text()
                    record.add_field(pymarc.Field(tag, data=entry_data))
                else:
                    indicators[0], indicators[1], record_data = line.find_all('td')
                    indicators[0] = indicators[0].get_text()
                    indicators[1] = indicators[1].get_text()
                    if re.match('(?:\s$)', indicators[0]):
                        indicators[0] = '\\'
                    if re.match('(?:\s$)', indicators[1]):
                        indicators[1] = '\\'
                    for subfield_data in re.finditer('(?:<strong>\|)(.)(?:<\/strong>\s+)(.+)(?:\s+\n|\s+$)', str(record_data)):
                        subfields.append(subfield_data.group(1))
                        subfields.append(subfield_data.group(2))
                    record.add_field(pymarc.Field(tag, indicators, subfields))
    yield record

#Converts xml marc record from Worldcat to pymarc record.
def get_marc_worldcat_xml(marc_record):
    record = None
    if type(marc_record) is bs4.element.Tag:
        marc_text_generator = marc_record.find_all('record')
    else:
        marc_record_soup = BeautifulSoup(marc_record, 'lxml-xml', from_encoding='utf-8')
        marc_text_generator = marc_record_soup.find_all('record')
    a = 1
    for xml_record in marc_text_generator:
        if ('xmlns:' in list(xml_record.attrs) and (xml_record['xmlns:'] == 'http://www.loc.gov/MARC21/slim' or 'http://www.loc.gov/MARC21/slim' in xml_record['xmlns:'])) or ('xmlns' in list(xml_record.attrs) and (xml_record['xmlns'] == 'http://www.loc.gov/MARC21/slim' or 'http://www.loc.gov/MARC21/slim' in xml_record['xmlns'])):
            if record is not None:
                yield record
            record = pymarc.Record()
            record.leader = xml_record.find('leader').get_text()
            for xml_field in xml_record.find_all('controlfield'):
                tag = xml_field['tag']
                entry_data = xml_field.get_text()
                record.add_field(pymarc.Field(tag, data=entry_data))
            for xml_field in xml_record.find_all('datafield'):
                tag = xml_field['tag']
                indicators = [xml_field['ind1'], xml_field['ind2']]
                if re.match('(?:\s$)', indicators[0]):
                    indicators[0] = '\\'
                if re.match('(?:\s$)', indicators[1]):
                    indicators[1] = '\\'
                subfields = []
                for xml_subfield in xml_field.find_all('subfield'):
                    subfields.append(xml_subfield['code'])
                    subfields.append(xml_subfield.get_text())
                record.add_field(pymarc.Field(tag, indicators, subfields))
    yield record

#Converts text marc record from BTAA to pymarc record.
def get_marc_btaa_xml(marc_record):
    record = None
    if type(marc_record) is bs4.element.Tag:
        marc_text_generator = marc_record
    elif type(marc_record) is str:
        marc_record_soup = BeautifulSoup(marc_record, 'lxml-xml', from_encoding='utf-8')
        marc_text_generator = marc_record_soup.find_all('record')
    else:
        marc_text_generator = marc_record
    a = 1
    for xml_record in marc_text_generator:
        if record is not None:
            yield record
        record = pymarc.Record()
        record.leader = xml_record.find('leader').get_text()
        for xml_field in xml_record.find_all('controlfield'):
            tag = xml_field['tag']
            entry_data = xml_field.get_text()
            record.add_field(pymarc.Field(tag, data=entry_data))
        for xml_field in xml_record.find_all('datafield'):
            tag = xml_field['tag']
            indicators = [xml_field['ind1'], xml_field['ind2']]
            if re.match('(?:\s$)', indicators[0]):
                indicators[0] = '\\'
            if re.match('(?:\s$)', indicators[1]):
                indicators[1] = '\\'
            subfields = []
            for xml_subfield in xml_field.find_all('subfield'):
                subfields.append(xml_subfield['code'])
                subfields.append(xml_subfield.get_text())
            record.add_field(pymarc.Field(tag, indicators, subfields))
    yield record



def line_generators(lines):
    for line in lines:
        if not re.match('([\s\n\r]+$)', line):
            yield line

def reader(data, marc_type = 'worldcat'):
    marc_records = None
    #Input as file
    if isinstance(data, io.TextIOWrapper):
        lines = data.readlines()
        marc_records = line_generators(lines)
    #Input as text
    else:
        marc_records = data
    if marc_type == 'worldcat':
        return get_marc_worldcat(marc_records)
    elif marc_type == 'millennium':
        return get_marc_millennium(marc_records)
    elif marc_type == 'issn':
        return get_marc_issn(marc_records)
    elif marc_type == 'worldcat_xml':
        return get_marc_worldcat_xml(marc_records)
    elif marc_type == 'btaa_xml':
        return get_marc_btaa_xml(marc_records)
