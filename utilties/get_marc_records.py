#*- coding: utf-8 -*-
import sys
import re
import os
import io
import pymarc
#sys.path.append(os.path.abspath('C:\\Users\\tmoss\\Desktop\\Python summary holding modules\\format_date'))
sys.path.append(os.path.abspath('C:\\Users\\tmoss\\Desktop\\Python files\\crl_machine\\app'))
#from format_date import *
from crl.fetch_from_api import marc_from_oclc, marc_from_issn, marc_from_lccn


#Converts text marc record from Millennium to pymarc record.
def get_marc_millennium(marc_record):
    record = None
    marc_record_temp = ''
    while re.match('(.*)(?:\n{2})(.*$)', marc_record, re.MULTILINE):
        marc_record = re.match('(.*)(?:\n{2})(.*$)', marc_record, re.MULTILINE).group(1) + '\n' + re.match('(.*)(?:\n{2})(.*$)', marc_record, re.MULTILINE).group(2)
    while re.match('(.*)(?:\n\s{2}\s+)(.*$)', marc_record, re.MULTILINE):
        marc_record = re.match('(.*)(?:\n\s{2}\s+)(.*$)', marc_record, re.MULTILINE).group(1) + re.match('(.*)(?:\n\s{2}\s+)(.*$)', marc_record, re.MULTILINE).group(2)
    for line0 in re.finditer('(^.*)(?:\n|$)', marc_record, re.MULTILINE):
        if marc_record_temp != '' and re.search('(\d{3}\s[\s\d]{2}\s|LDR|LEADER)', line0.group(1)):
            marc_record_temp = marc_record_temp + '\n'
        if re.match('(?:\s+)(\S.*$)', line0.group(1)):
            marc_record_temp = marc_record_temp + re.match('(?:\s+)(\S.*$)', line0.group(1)).group(1)
        else:
            marc_record_temp = marc_record_temp + line0.group(1)
    for line0 in re.finditer('(^.*)(?:\n|$)', marc_record_temp, re.MULTILINE):
        line = line0.group(1)
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
#    if re.search('(?:\n)([^\n]*$)', marc_record_temp):
#        line = re.search('(?:\n)([^\n]*$)', marc_record_temp).group(1)
#        tag = ''
#        indicators = []
#        subfields = []
#        entry_data = ''
#        if re.match('(?:LDR\s*|LEADER\s*)(.+)', line):
#            record.leader = re.match('(?:LDR\s*|LEADER\s*)(.+)', line).group(1)
#        elif re.match('(\d{3})', line):
#            tag = re.search('(\d{3})', line).group(1)
#            if int(re.search('(\d{3})', line).group(1)) < 10:
#                if re.search('(?:\d{3}\s+)(.+$)', line):
#                    entry_data = re.search('(?:\d{3}\s+)(.+$)', line).group(1)
#                record.add_field(pymarc.Field(tag, data=entry_data))
#            else:
#                if re.search('(?:\d{3}\s)(.)(.)(?:\s)(.+$)', line):
#                    if re.match('(?:\s$)', re.search('(?:\d{3}\s)(.)(.)(?:\s)(.+$)', line).group(1)):
#                        indicators.append('\\')
#                    else:
#                        indicators.append(re.search('(?:\d{3}\s)(.)(.)(?:\s)(.+$)', line).group(1))
#                    if re.match('(?:\s$)', re.search('(?:\d{3}\s)(.)(.)(?:\s)(.+$)', line).group(2)):
#                        indicators.append('\\')
#                    else:
#                        indicators.append(re.search('(?:\d{3}\s)(.)(.)(?:\s)(.+$)', line).group(2))
#                    base_subfields = re.search('(?:\d{3}\s)(.)(.)(?:\s)(.+$)', line).group(3)
#                    if not re.match('(?:\|)', base_subfields):
#                        base_subfields = '|a' + base_subfields
#                    for m in re.finditer('(?:\|)([^\|])([^\|]+)', base_subfields):
#                        subfields.append(m.group(1))
#                        subfields.append(m.group(2))
#                record.add_field(pymarc.Field(tag, indicators, subfields))
    yield record

#Converts text marc record from Worldcat to pymarc record.
def get_marc_worldcat(marc_record):
    record = None
    marc_record_temp = ''
    while re.match('(.*)(?:\n{2})(.*$)', marc_record, re.MULTILINE):
        marc_record = re.match('(.*)(?:\n{2})(.*$)', marc_record, re.MULTILINE).group(1) + '\n' + re.match('(.*)(?:\n{2})(.*$)', marc_record, re.MULTILINE).group(2)
    while re.match('(.*)(?:\n\s{2}\s+)(.*$)', marc_record, re.MULTILINE):
        marc_record = re.match('(.*)(?:\n\s{2}\s+)(.*$)', marc_record, re.MULTILINE).group(1) + re.match('(.*)(?:\n\s{2}\s+)(.*$)', marc_record, re.MULTILINE).group(2)
    for line0 in re.finditer('(^.*)(?:\n)', marc_record, re.MULTILINE):
        if marc_record_temp != '' and re.search('(\d{3}\s[\s\d]{2}\s|LDR|LEADER)', line0.group(1)):
            marc_record_temp = marc_record_temp + '\n'
        if re.match('(?:\s+)(\S.*$)', line0.group(1)):
            marc_record_temp = marc_record_temp + re.match('(?:\s+)(\S.*$)', line0.group(1)).group(1)
        else:
            marc_record_temp = marc_record_temp + line0.group(1)
    #Create pymarc record from text marc record
    for line0 in re.finditer('(^.*)(?:\n|$)', marc_record, re.MULTILINE):
        line = line0.group(1)
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
    #Create pymarc record from text marc record
#    if re.search('(?:\n)([^\n]*$)', marc_record_temp):
#        line = re.search('(?:\n)([^\n]*$)', marc_record_temp).group(1)
#        tag = ''
#        indicators = []
#        subfields = []
#        entry_data = ''
#        if re.match('(?:LDR\s*|LEADER\s*)(.+)', line):
#            record.leader = re.match('(?:LDR\s*|LEADER\s*)(.+)', line).group(1)
#        elif re.match('(\d{3})', line):
#            tag = re.search('(\d{3})', line).group(1)
#            if int(re.search('(\d{3})', line).group(1)) < 10:
#                if re.search('(?:\d{3}\s+)(.+$)', line):
#                    entry_data = re.search('(?:\d{3}\s+)(.+$)', line).group(1)
#                record.add_field(pymarc.Field(tag, data=entry_data))
#            else:
#                if re.search('(?:\=\d{3}\s+)(.)(.)(.+$)', line):
#                    if re.match('(?:\s$)', re.search('(?:\=\d{3}\s+)(.)(.)(.+$)', line).group(1)):
#                        indicators.append('\\')
#                    else:
#                        indicators.append(re.search('(?:\=\d{3}\s+)(.)(.)(.+$)', line).group(1))
#                    if re.match('(?:\s$)', re.search('(?:\=\d{3}\s+)(.)(.)(.+$)', line).group(2)):
#                        indicators.append('\\')
#                    else:
#                        indicators.append(re.search('(?:\=\d{3}\s+)(.)(.)(.+$)', line).group(2))
#                    base_subfields = re.search('(?:\=\d{3}\s+)(.)(.)(.+$)', line).group(3)
#                    for m in re.finditer('(?:\$)([^\$])([^\$]+)', base_subfields):
#                        subfields.append(m.group(1))
#                        subfields.append(m.group(2))
#                record.add_field(pymarc.Field(tag, indicators, subfields))
    yield record

#Converts text marc record from ISSN database to pymarc record.
def get_marc_issn(marc_record):
    record = None
    marc_record_temp = ''
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
    for line0 in re.finditer('(^.*)(?:\n|$)', marc_record_temp, re.MULTILINE):
        line = line0.group(1)
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
#    if re.search('(?:\n)([^\n]*$)', marc_record_temp):
#        line = re.search('(?:\n)([^\n]*$)', marc_record_temp).group(1)
#        tag = ''
#        indicators = []
#        subfields = []
#        entry_data = ''
#        if re.match('(?:\=LDR\s*|\=LEADER\s*)(.+)', line):
#            record.leader = re.match('(?:\=LDR\s*|\=LEADER\s*)(.+)', line).group(1)
#        elif re.match('(?:\=)(\d{3})', line):
#            tag = re.search('(?:\=)(\d{3})', line).group(1)
#            if int(re.search('(?:\=)(\d{3})', line).group(1)) < 10:
#                if re.search('(?:\=\d{3}\s+)(.+$)', line):
#                    entry_data = re.search('(?:\=\d{3}\s+)(.+$)', line).group(1)
#                record.add_field(pymarc.Field(tag, data=entry_data))
#            else:
#                if re.search('(?:\=\d{3}\s{2})(.)(.)(.+$)', line):
#                    if re.match('(?:\s$)', re.search('(?:\=\d{3}\s{2})(.)(.)(.+$)', line).group(1)):
#                        indicators.append('\\')
#                    else:
#                        indicators.append(re.search('(?:\=\d{3}\s{2})(.)(.)(.+$)', line).group(1))
#                    if re.match('(?:\s$)', re.search('(?:\=\d{3}\s{2})(.)(.)(.+$)', line).group(2)):
#                        indicators.append('\\')
#                    else:
#                        indicators.append(re.search('(?:\=\d{3}\s{2})(.)(.)(.+$)', line).group(2))
#                    base_subfields = re.search('(?:\=\d{3}\s{2})(.)(.)(.+$)', line).group(3)
#                    if not re.match('(?:\$)', base_subfields):
#                        base_subfields = '$a' + base_subfields
#                    for m in re.finditer('(?:\$)([^\$])([^\$]+)', base_subfields):
#                        subfields.append(m.group(1))
#                        subfields.append(m.group(2))
#                record.add_field(pymarc.Field(tag, indicators, subfields))
    yield record

def reader(data, marc_type = 'worldcat'):
    marc_records = None
    #Input as file
    if isinstance(data, io.TextIOWrapper):
        lines = data.readlines()
        marc_records = ''
        for line in lines:
            marc_records = marc_records + line
    #Input as text
    else:
        marc_records = data
    if marc_type == 'worldcat':
        return get_marc_worldcat(marc_records)
    elif marc_type == 'millennium':
        return get_marc_millennium(marc_records)
    elif marc_type == 'issn':
        return get_marc_issn(marc_records)