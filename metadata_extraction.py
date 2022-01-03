#*- coding: utf-8 -*-
import sys
import csv
import re
import unicodedata
import subprocess
import shlex
import os
import time
import shutil
import copy
import pymarc
import sqlite3
from bs4 import BeautifulSoup
import urllib.request
import tkinter
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename
from tkinter.filedialog import askdirectory
from tkinter import ttk
from tkinter import font
import threading
#sys.path.append(os.path.abspath('C:\\Users\\tmoss\\Desktop\\Python summary holding modules\\format_date'))
sys.path.append(os.path.abspath('C:\\Users\\tmoss\\Desktop\\Python files\\crl_machine\\app'))
#from format_date import *
from crl.fetch_from_api import marc_from_oclc, marc_from_issn, marc_from_lccn
#import get_marc_records
#sys.path.append(os.path.abspath('C:\\Users\\tmoss\\Desktop\\Python files\\get_marc_records'))
#import utilties.get_marc_records as get_marc_records
#from . import get_marc_records
#from utilties import get_marc_records
#sys.path.append(os.path.abspath('C:\\Users\\tmoss\\Desktop\\Python files\\metadata_extraction\\utilties'))
sys.path.append(os.path.abspath('C:\\Users\\tmoss\\Documents\\GitHub\\metadata_extraction\\utilties'))
import get_marc_records
from functools import partial
import configparser
from collections import OrderedDict
import openpyxl
import win32clipboard

#Creates the folder if it does not already exist
def check_or_create_dir(path):
    if not os.path.exists(path):
        if not os.path.exists(os.path.dirname(path)):
            check_or_create_dir(os.path.dirname(path))
        os.mkdir(path)

file_direct = os.path.dirname(os.path.realpath('__file__'))
output_folder = os.path.join(file_direct, 'Output')
check_or_create_dir(output_folder)

country_dict = {}
country_dict['sdu'] = 'United States'
country_dict['xs'] = 'South Georgia and the South Sandwich Islands'
country_dict['sd'] = 'South Sudan'
country_dict['sp'] = 'Spain'
country_dict['sh'] = 'Spanish North Africa'
country_dict['xp'] = 'Spratly Island'
country_dict['ce'] = 'Sri Lanka'
country_dict['sj'] = 'Sudan'
country_dict['sr'] = 'Surinam'
country_dict['sq'] = 'Swaziland'
country_dict['sw'] = 'Sweden'
country_dict['sz'] = 'Switzerland'
country_dict['sy'] = 'Syria'
country_dict['ta'] = 'Tajikistan'
country_dict['tz'] = 'Tanzania'
country_dict['tma'] = 'Australia'
country_dict['tnu'] = 'United States'
country_dict['fs'] = 'Terres australes et antarctiques françaises'
country_dict['txu'] = 'United States'
country_dict['th'] = 'Thailand'
country_dict['em'] = 'Timor-Leste'
country_dict['tg'] = 'Togo'
country_dict['tl'] = 'Tokelau'
country_dict['to'] = 'Tonga'
country_dict['tr'] = 'Trinidad and Tobago'
country_dict['ti'] = 'Tunisia'
country_dict['tu'] = 'Turkey'
country_dict['tk'] = 'Turkmenistan'
country_dict['tc'] = 'Turks and Caicos Islands'
country_dict['tv' ] = 'Tuvalu'
country_dict['ug'] = 'Uganda'
country_dict['un'] = 'Ukraine'
country_dict['ts'] = 'United Arab Emirates'
country_dict['xxk'] = 'United Kingdom'
country_dict['xxu' ] = 'United States'
country_dict['uc'] = 'United States Misc. Caribbean Islands'
country_dict['up'] = 'United States Misc. Pacific Islands'
country_dict['uy'] = 'Uruguay'
country_dict['utu'] = 'United States'
country_dict['uz'] = 'Uzbekistan'
country_dict['nn'] = 'Vanuatu'
country_dict['vp'] = 'Various places'
country_dict['vc'] = 'Vatican City'
country_dict['ve'] = 'Venezuela'
country_dict['vtu'] = 'United States'
country_dict['vra'] = 'Australia'
country_dict['vm'] = 'Vietnam'
country_dict['vi' ] = 'Virgin Islands of the United States'
country_dict['vau'] = 'United States'
country_dict['wk'] = 'Wake Island'
country_dict['wlk'] = 'United Kingdom'
country_dict['wf'] = 'Wallis and Futuna'
country_dict['wau'] = 'United States'
country_dict['wj'] = 'West Bank of the Jordan River'
country_dict['wvu'] = 'United States'
country_dict['wea'] = 'Australia'
country_dict['ss'] = 'Western Sahara'
country_dict['wiu'] = 'United States'
country_dict['wyu'] = 'United States'
country_dict['ye'] = 'Yemen'
country_dict['ykc'] = 'Canada'
country_dict['za'] = 'Zambia'
country_dict['rh'] = 'Zimbabwe'
country_dict['ac'] = 'Australia'
country_dict['ai'] = 'Armenia (Republic)'
country_dict['air'] = 'Armenia (Republic)'
country_dict['ajr'] = 'Azerbaijan'
country_dict['bwr'] = 'Belarus'
country_dict['cn'] = 'Canada'
country_dict['cp'] = 'Kiribati'
country_dict['cs'] = 'Czechoslovakia'
country_dict['cz'] = 'Panama'
country_dict['err'] = 'Estonia'
country_dict['ge'] = 'Germany'
country_dict['gn'] = 'Kiribati'
country_dict['gsr'] = 'Georgia (Republic)'
country_dict['hk'] = 'China'
country_dict['iu'] = 'Israel'
country_dict['iw'] = 'Israel'
country_dict['jn'] = 'Norway'
country_dict['kgr'] = 'Kyrgyzstan'
country_dict['kzr'] = 'Kazakhstan'
country_dict['lir'] = 'Lithuania'
country_dict['ln'] = 'Kiribati'
country_dict['lvr'] = 'Latvia'
country_dict['mh'] = 'China'
country_dict['mvr'] = 'Moldova'
country_dict['na'] = 'Netherlands Antilles'
country_dict['nm'] = 'Northern Mariana Islands'
country_dict['pt'] = 'Timor-Leste'
country_dict['rur'] = 'Russia (Federation)'
country_dict['ry'] = 'Japan'
country_dict['sb'] = 'Norway'
country_dict['sk'] = 'India'
country_dict['sv'] = 'Honduras'
country_dict['tar'] = 'Tajikistan'
country_dict['tkr'] = 'Turkmenistan'
country_dict['tt'] = 'Trust Territory of the Pacific Islands'
country_dict['ui'] = 'United Kingdom Misc. Islands'
country_dict['uik'] = 'United Kingdom Misc. Islands'
country_dict['uk'] = 'United Kingdom'
country_dict['unr'] = 'Ukraine'
country_dict['ur'] = 'Soviet Union'
country_dict['us'] = 'United States'
country_dict['uzr'] = 'Uzbekistan'
country_dict['vn'] = 'Vietnam'
country_dict['vs'] = 'Vietnam'
country_dict['wb'] = 'Germany'
country_dict['xi'] = 'Saint Kitts-Nevis-Anguilla'
country_dict['xxr'] = 'Soviet Union'
country_dict['ys'] = 'Yemen'
country_dict['yu'] = 'Serbia and Montenegro'

country_dict['aa'] = 'Albania'
country_dict['abc'] = 'Canada'
country_dict['aca'] = 'Australia'
country_dict['ae'] = 'Algeria'
country_dict['af'] = 'Afghanistan'
country_dict['ag'] = 'Argentina'
country_dict['aj'] = 'Azerbaijan'
country_dict['aku'] = 'United States'
country_dict['alu'] = 'United States'
country_dict['am'] = 'Anguilla'
country_dict['an'] = 'Andorra'
country_dict['ao'] = 'Angola'
country_dict['aq'] = 'Antigua and Barbuda'
country_dict['aru'] = 'United States'
country_dict['as'] = 'American Samoa'
country_dict['at'] = 'Australia'
country_dict['au'] = 'Austria'
country_dict['aw'] = 'Aruba'
country_dict['ay'] = 'Antarctica'
country_dict['azu'] = 'United States'
country_dict['ba'] = 'Bahrain'
country_dict['bb'] = 'Barbados'
country_dict['bcc'] = 'Canada'
country_dict['bd'] = 'Burundi'
country_dict['be'] = 'Belgium'
country_dict['bf'] = 'Bahamas'
country_dict['bg'] = 'Bangladesh'
country_dict['bh'] = 'Belize'
country_dict['bi'] = 'British Indian Ocean Territory'
country_dict['bl'] = 'Brazil'
country_dict['bm'] = 'Bermuda Islands'
country_dict['bn'] = 'Bosnia and Herzegovina'
country_dict['bo'] = 'Bolivia'
country_dict['bp'] = 'Solomon Islands'
country_dict['br'] = 'Burma'
country_dict['bs'] = 'Botswana'
country_dict['bt'] = 'Bhutan'
country_dict['bu'] = 'Bulgaria'
country_dict['bv'] = 'Bouvet Island'
country_dict['bw'] = 'Belarus'
country_dict['bx'] = 'Brunei'
country_dict['ca'] = 'Caribbean Netherlands'
country_dict['cau'] = 'United States'
country_dict['cb'] = 'Cambodia'
country_dict['cc'] = 'China'
country_dict['cd'] = 'Chad'
country_dict['cf'] = 'Congo (Brazzaville)'
country_dict['cg'] = 'Congo (Democratic Republic)'
country_dict['ch'] = 'China (Republic : 1949- )'
country_dict['ci'] = 'Croatia'
country_dict['cj'] = 'Cayman Islands'
country_dict['ck'] = 'Colombia'
country_dict['cl'] = 'Chile'
country_dict['cm'] = 'Cameroon'
country_dict['co'] = 'Curaçao'
country_dict['cou'] = 'United States'
country_dict['cq'] = 'Comoros'
country_dict['cr'] = 'Costa Rica'
country_dict['ctu'] = 'United States'
country_dict['cu'] = 'Cuba'
country_dict['cv'] = 'Cabo Verde'
country_dict['cw'] = 'Cook Islands'
country_dict['cx'] = 'Central African Republic'
country_dict['cy'] = 'Cyprus'
country_dict['dcu'] = 'United States'
country_dict['deu'] = 'United States'
country_dict['dk'] = 'Denmark'
country_dict['dm'] = 'Benin'
country_dict['dq'] = 'Dominica'
country_dict['dr'] = 'Dominican Republic'
country_dict['ea'] = 'Eritrea'
country_dict['ec'] = 'Ecuador'
country_dict['eg'] = 'Equatorial Guinea'
country_dict['enk'] = 'United Kingdom'
country_dict['er'] = 'Estonia'
country_dict['es'] = 'El Salvador'
country_dict['et'] = 'Ethiopia'
country_dict['fa'] = 'Faroe Islands'
country_dict['fg'] = 'French Guiana'
country_dict['fi'] = 'Finland'
country_dict['fj'] = 'Fiji'
country_dict['fk'] = 'Falkland Islands'
country_dict['flu'] = 'United States'
country_dict['fm'] = 'Micronesia (Federated States)'
country_dict['fp'] = 'French Polynesia'
country_dict['fr'] = 'France'
country_dict['ft'] = 'Djibouti'
country_dict['gau'] = 'United States'
country_dict['gb'] = 'Kiribati'
country_dict['gd'] = 'Grenada'
country_dict['gg'] = 'Guernsey'
country_dict['gh'] = 'Ghana'
country_dict['gi'] = 'Gibraltar'
country_dict['gl'] = 'Greenland'
country_dict['gm'] = 'Gambia'
country_dict['go'] = 'Gabon'
country_dict['gp'] = 'Guadeloupe'
country_dict['gr'] = 'Greece'
country_dict['gs'] = 'Georgia (Republic)'
country_dict['gt'] = 'Guatemala'
country_dict['gu'] = 'Guam'
country_dict['gv'] = 'Guinea'
country_dict['gw'] = 'Germany'
country_dict['gy'] = 'Guyana'
country_dict['gz'] = 'Gaza Strip'
country_dict['hiu'] = 'United States'
country_dict['hm'] = 'Heard and McDonald Islands'
country_dict['ho'] = 'Honduras'
country_dict['ht'] = 'Haiti'
country_dict['hu'] = 'Hungary'
country_dict['iau'] = 'United States'
country_dict['ic'] = 'Iceland'
country_dict['idu'] = 'United States'
country_dict['ie'] = 'Ireland'
country_dict['ii'] = 'India'
country_dict['ilu'] = 'United States'
country_dict['im'] = 'Isle of Man'
country_dict['inu'] = 'United States'
country_dict['io'] = 'Indonesia'
country_dict['iq'] = 'Iraq'
country_dict['ir'] = 'Iran'
country_dict['is'] = 'Israel'
country_dict['it'] = 'Italy'
country_dict['iv'] = 'Côte d\'Ivoire'
country_dict['iy'] = 'Iraq-Saudi Arabia Neutral Zone'
country_dict['ja'] = 'Japan'
country_dict['je'] = 'Jersey'
country_dict['ji'] = 'Johnston Atoll'
country_dict['jm'] = 'Jamaica'
country_dict['jo'] = 'Jordan'
country_dict['ke'] = 'Kenya'
country_dict['kg'] = 'Kyrgyzstan'
country_dict['kn'] = 'Korea (North)'
country_dict['ko'] = 'Korea (South)'
country_dict['ksu'] = 'United States'
country_dict['ku'] = 'Kuwait'
country_dict['kv'] = 'Kosovo'
country_dict['kyu'] = 'United States'
country_dict['kz'] = 'Kazakhstan'
country_dict['lau'] = 'United States'
country_dict['lb'] = 'Liberia'
country_dict['le'] = 'Lebanon'
country_dict['lh'] = 'Liechtenstein'
country_dict['li'] = 'Lithuania'
country_dict['lo'] = 'Lesotho'
country_dict['ls'] = 'Laos'
country_dict['lu'] = 'Luxembourg'
country_dict['lv'] = 'Latvia'
country_dict['ly'] = 'Libya'
country_dict['mau'] = 'United States'
country_dict['mbc'] = 'Canada'
country_dict['mc'] = 'Monaco'
country_dict['mdu'] = 'United States'
country_dict['meu'] = 'United States'
country_dict['mf'] = 'Mauritius'
country_dict['mg'] = 'Madagascar'
country_dict['miu'] = 'United States'
country_dict['mj'] = 'Montserrat'
country_dict['mk'] = 'Oman'
country_dict['ml'] = 'Mali'
country_dict['mm'] = 'Malta'
country_dict['mnu'] = 'United States'
country_dict['mo'] = 'Montenegro'
country_dict['mou'] = 'United States'
country_dict['mp'] = 'Mongolia'
country_dict['mq'] = 'Martinique'
country_dict['mr'] = 'Morocco'
country_dict['msu'] = 'United States'
country_dict['mtu'] = 'United States'
country_dict['mu'] = 'Mauritania'
country_dict['mv'] = 'Moldova'
country_dict['mw'] = 'Malawi'
country_dict['mx'] = 'Mexico'
country_dict['my'] = 'Malaysia'
country_dict['mz'] = 'Mozambique'
country_dict['nbu'] = 'United States'
country_dict['ncu'] = 'United States'
country_dict['ndu'] = 'United States'
country_dict['ne'] = 'Netherlands'
country_dict['nfc'] = 'Canada'
country_dict['ng'] = 'Niger'
country_dict['nhu'] = 'United States'
country_dict['nik'] = 'United Kingdom'
country_dict['nju'] = 'United States'
country_dict['nkc'] = 'Canada'
country_dict['nl'] = 'New Caledonia'
country_dict['nmu'] = 'United States'
country_dict['no'] = 'Norway'
country_dict['np'] = 'Nepal'
country_dict['nq'] = 'Nicaragua'
country_dict['nr'] = 'Nigeria'
country_dict['nsc'] = 'Canada'
country_dict['ntc'] = 'Canada'
country_dict['nu'] = 'Nauru'
country_dict['nuc'] = 'Canada'
country_dict['nvu'] = 'United States'
country_dict['nw'] = 'Northern Mariana Islands'
country_dict['nx'] = 'Norfolk Island'
country_dict['nyu'] = 'United States'
country_dict['nz'] = 'New Zealand'
country_dict['ohu'] = 'United States'
country_dict['oku'] = 'United States'
country_dict['onc'] = 'Canada'
country_dict['oru'] = 'United States'
country_dict['ot'] = 'Mayotte'
country_dict['pau'] = 'United States'
country_dict['pc'] = 'Pitcairn Island'
country_dict['pe'] = 'Peru'
country_dict['pf'] = 'Paracel Islands'
country_dict['pg'] = 'Guinea-Bissau'
country_dict['ph'] = 'Philippines'
country_dict['pic'] = 'Canada'
country_dict['pk'] = 'Pakistan'
country_dict['pl'] = 'Poland'
country_dict['pn'] = 'Panama'
country_dict['po'] = 'Portugal'
country_dict['pp'] = 'Papua New Guinea'
country_dict['pr'] = 'Puerto Rico'
country_dict['pw'] = 'Palau'
country_dict['py'] = 'Paraguay'
country_dict['qa'] = 'Qatar'
country_dict['qea'] = 'Queensland'
country_dict['quc'] = 'Canada'
country_dict['rb'] = 'Serbia'
country_dict['re'] = 'Réunion'
country_dict['riu'] = 'United States'
country_dict['rm'] = 'Romania'
country_dict['ru'] = 'Russia (Federation)'
country_dict['rw'] = 'Rwanda'
country_dict['sa'] = 'South Africa'
country_dict['sc'] = 'Saint-Barthélemy'
country_dict['scu'] = 'United States'
country_dict['se'] = 'Seychelles'
country_dict['sf'] = 'Sao Tome and Principe'
country_dict['sg'] = 'Senegal'
country_dict['si'] = 'Singapore'
country_dict['sl'] = 'Sierra Leone'
country_dict['sm'] = 'San Marino'
country_dict['sn'] = 'Sint Maarten'
country_dict['snc'] = 'Canada'
country_dict['so'] = 'Somalia'
country_dict['st'] = 'Saint-Martin'
country_dict['stk'] = 'United Kingdom'
country_dict['su'] = 'Saudi Arabia'
country_dict['sx'] = 'Namibia'
country_dict['tv'] = 'Tuvalu'
country_dict['ua'] = 'Egypt'
country_dict['uv'] = 'Burkina Faso'
country_dict['vb'] = 'British Virgin Islands'
country_dict['vi'] = 'Virgin Islands of the United States'
country_dict['ws'] = 'Samoa'
country_dict['xa'] = 'Christmas Island (Indian Ocean)'
country_dict['xb'] = 'Cocos (Keeling) Islands'
country_dict['xc'] = 'Maldives'
country_dict['xd'] = 'Saint Kitts-Nevis'
country_dict['xe'] = 'Marshall Islands'
country_dict['xf'] = 'Midway Islands'
country_dict['xga'] = 'Coral Sea Islands Territory'
country_dict['xh'] = 'Niue'
country_dict['xj'] = 'Saint Helena'
country_dict['xk'] = 'Saint Lucia'
country_dict['xl'] = 'Saint Pierre and Miquelon'
country_dict['xm'] = 'Saint Vincent and the Grenadines'
country_dict['xn'] = 'Macedonia'
country_dict['xna'] = 'New South Wales'
country_dict['xo'] = 'Slovakia'
country_dict['xoa'] = 'Northern Territory'
country_dict['xr'] = 'Czech Republic'
country_dict['xra'] = 'South Australia'
country_dict['xv'] = 'Slovenia'
country_dict['xx'] = 'No place, unknown, or undetermined'
country_dict['xxc'] = 'Canada'
country_dict['xxu'] = 'United States'


language_dict = {}
language_dict['aar'] = 'Afar'
language_dict['abk'] = 'Abkhaz'
language_dict['ace'] = 'Achinese'
language_dict['ach'] = 'Acoli'
language_dict['ada'] = 'Adangme'
language_dict['ady'] = 'Adygei'
language_dict['afa'] = 'Afroasiatic (Other)'
language_dict['afh'] = 'Afrihili (Artificial language)'
language_dict['afr'] = 'Afrikaans'
language_dict['ain'] = 'Ainu'
language_dict['ajm'] = 'Aljamía'
language_dict['aka'] = 'Akan'
language_dict['akk'] = 'Akkadian'
language_dict['alb'] = 'Albanian'
language_dict['ale'] = 'Aleut'
language_dict['alg'] = 'Algonquian (Other)'
language_dict['alt'] = 'Altai'
language_dict['amh'] = 'Amharic'
language_dict['ang'] = 'English, Old (ca. 450-1100)'
language_dict['anp'] = 'Angika'
language_dict['apa'] = 'Apache languages'
language_dict['ara'] = 'Arabic'
language_dict['arc'] = 'Aramaic'
language_dict['arg'] = 'Aragonese'
language_dict['arm'] = 'Armenian'
language_dict['arn'] = 'Mapuche'
language_dict['arp'] = 'Arapaho'
language_dict['art'] = 'Artificial (Other)'
language_dict['arw'] = 'Arawak'
language_dict['asm'] = 'Assamese'
language_dict['ast'] = 'Bable'
language_dict['ath'] = 'Athapascan (Other)'
language_dict['aus'] = 'Australian languages'
language_dict['ava'] = 'Avaric'
language_dict['ave'] = 'Avestan'
language_dict['awa'] = 'Awadhi'
language_dict['aym'] = 'Aymara'
language_dict['aze'] = 'Azerbaijani'
language_dict['bad'] = 'Banda languages'
language_dict['bai'] = 'Bamileke languages'
language_dict['bak'] = 'Bashkir'
language_dict['bal'] = 'Baluchi'
language_dict['bam'] = 'Bambara'
language_dict['ban'] = 'Balinese'
language_dict['baq'] = 'Basque'
language_dict['bas'] = 'Basa'
language_dict['bat'] = 'Baltic (Other)'
language_dict['bej'] = 'Beja'
language_dict['bel'] = 'Belarusian'
language_dict['bem'] = 'Bemba'
language_dict['ben'] = 'Bengali'
language_dict['ber'] = 'Berber (Other)'
language_dict['bho'] = 'Bhojpuri'
language_dict['bih'] = 'Bihari (Other) '
language_dict['bik'] = 'Bikol'
language_dict['bin'] = 'Edo'
language_dict['bis'] = 'Bislama'
language_dict['bla'] = 'Siksika'
language_dict['bnt'] = 'Bantu (Other)'
language_dict['bos'] = 'Bosnian'
language_dict['bra'] = 'Braj'
language_dict['bre'] = 'Breton'
language_dict['btk'] = 'Batak'
language_dict['bua'] = 'Buriat'
language_dict['bug'] = 'Bugis'
language_dict['bul'] = 'Bulgarian'
language_dict['bur'] = 'Burmese'
language_dict['byn'] = 'Bilin'
language_dict['cad'] = 'Caddo'
language_dict['cai'] = 'Central American Indian (Other)'
language_dict['cam'] = 'Khmer'
language_dict['car'] = 'Carib'
language_dict['cat'] = 'Catalan'
language_dict['cau'] = 'Caucasian (Other)'
language_dict['ceb'] = 'Cebuano'
language_dict['cel'] = 'Celtic (Other)'
language_dict['cha'] = 'Chamorro'
language_dict['chb'] = 'Chibcha'
language_dict['che'] = 'Chechen'
language_dict['chg'] = 'Chagatai'
language_dict['chi'] = 'Chinese'
language_dict['chk'] = 'Chuukese'
language_dict['chm'] = 'Mari'
language_dict['chn'] = 'Chinook jargon'
language_dict['cho'] = 'Choctaw'
language_dict['chp'] = 'Chipewyan'
language_dict['chr'] = 'Cherokee'
language_dict['chu'] = 'Church Slavic'
language_dict['chv'] = 'Chuvash'
language_dict['chy'] = 'Cheyenne'
language_dict['cmc'] = 'Chamic languages'
language_dict['cnr'] = 'Montenegrin'
language_dict['cop'] = 'Coptic'
language_dict['cor'] = 'Cornish'
language_dict['cos'] = 'Corsican'
language_dict['cpe'] = 'Creoles and Pidgins, English-based (Other)'
language_dict['cpf'] = 'Creoles and Pidgins, French-based (Other)'
language_dict['cpp'] = 'Creoles and Pidgins, Portuguese-based (Other)'
language_dict['cre'] = 'Cree'
language_dict['crh'] = 'Crimean Tatar'
language_dict['crp'] = 'Creoles and Pidgins (Other)'
language_dict['csb'] = 'Kashubian'
language_dict['cus'] = 'Cushitic (Other)'
language_dict['cze'] = 'Czech'
language_dict['dak'] = 'Dakota'
language_dict['dan'] = 'Danish'
language_dict['dar'] = 'Dargwa'
language_dict['day'] = 'Dayak'
language_dict['del'] = 'Delaware'
language_dict['den'] = 'Slavey'
language_dict['dgr'] = 'Dogrib'
language_dict['din'] = 'Dinka'
language_dict['div'] = 'Divehi'
language_dict['doi'] = 'Dogri'
language_dict['dra'] = 'Dravidian (Other)'
language_dict['dsb'] = 'Lower Sorbian'
language_dict['dua'] = 'Duala'
language_dict['dum'] = 'Dutch, Middle (ca. 1050-1350)'
language_dict['dut'] = 'Dutch'
language_dict['dyu'] = 'Dyula'
language_dict['dzo'] = 'Dzongkha'
language_dict['efi'] = 'Efik'
language_dict['egy'] = 'Egyptian'
language_dict['eka'] = 'Ekajuk'
language_dict['elx'] = 'Elamite'
language_dict['eng'] = 'English'
language_dict['enm'] = 'English, Middle (1100-1500)'
language_dict['epo'] = 'Esperanto'
language_dict['esk'] = 'Eskimo languages'
language_dict['esp'] = 'Esperanto'
language_dict['est'] = 'Estonian'
language_dict['eth'] = 'Ethiopic'
language_dict['ewe'] = 'Ewe'
language_dict['ewo'] = 'Ewondo'
language_dict['fan'] = 'Fang'
language_dict['fao'] = 'Faroese'
language_dict['far'] = 'Faroese'
language_dict['fat'] = 'Fanti'
language_dict['fij'] = 'Fijian'
language_dict['fil'] = 'Filipino'
language_dict['fin'] = 'Finnish'
language_dict['fiu'] = 'Finno-Ugrian (Other)'
language_dict['fon'] = 'Fon'
language_dict['fre'] = 'French'
language_dict['fri'] = 'Frisian'
language_dict['frm'] = 'French, Middle (ca. 1300-1600)'
language_dict['fro'] = 'French, Old (ca. 842-1300)'
language_dict['frr'] = 'North Frisian'
language_dict['frs'] = 'East Frisian'
language_dict['fry'] = 'Frisian'
language_dict['ful'] = 'Fula'
language_dict['fur'] = 'Friulian'
language_dict['gaa'] = 'Gã'
language_dict['gae'] = 'Scottish Gaelix'
language_dict['gag'] = 'Galician'
language_dict['gal'] = 'Oromo'
language_dict['gay'] = 'Gayo'
language_dict['gba'] = 'Gbaya'
language_dict['gem'] = 'Germanic (Other)'
language_dict['geo'] = 'Georgian'
language_dict['ger'] = 'German'
language_dict['gez'] = 'Ethiopic'
language_dict['gil'] = 'Gilbertese'
language_dict['gla'] = 'Scottish Gaelic'
language_dict['gle'] = 'Irish'
language_dict['glg'] = 'Galician'
language_dict['glv'] = 'Manx'
language_dict['gmh'] = 'German, Middle High (ca. 1050-1500)'
language_dict['goh'] = 'German, Old High (ca. 750-1050)'
language_dict['gon'] = 'Gondi'
language_dict['gor'] = 'Gorontalo'
language_dict['got'] = 'Gothic'
language_dict['grb'] = 'Grebo'
language_dict['grc'] = 'Greek, Ancient (to 1453)'
language_dict['gre'] = 'Greek, Modern (1453-)'
language_dict['grn'] = 'Guarani'
language_dict['gsw'] = 'Swiss German'
language_dict['gua'] = 'Guarani'
language_dict['guj'] = 'Gujarati'
language_dict['gwi'] = 'Gwich\'in'
language_dict['hai'] = 'Haida'
language_dict['hat'] = 'Haitian French Creole'
language_dict['hau'] = 'Hausa'
language_dict['haw'] = 'Hawaiian'
language_dict['heb'] = 'Hebrew'
language_dict['her'] = 'Herero'
language_dict['hil'] = 'Hiligaynon'
language_dict['him'] = 'Western Pahari languages'
language_dict['hin'] = 'Hindi'
language_dict['hit'] = 'Hittite'
language_dict['hmn'] = 'Hmong'
language_dict['hmo'] = 'Hiri Motu'
language_dict['hrv'] = 'Croatian'
language_dict['hsb'] = 'Upper Sorbian'
language_dict['hun'] = 'Hungarian'
language_dict['hup'] = 'Hupa'
language_dict['iba'] = 'Iban'
language_dict['ibo'] = 'Igbo'
language_dict['ice'] = 'Icelandic'
language_dict['ido'] = 'Ido'
language_dict['iii'] = 'Sichuan Yi'
language_dict['ijo'] = 'Ijo'
language_dict['iku'] = 'Inuktitut'
language_dict['ile'] = 'Interlingue'
language_dict['ilo'] = 'Iloko'
language_dict['ina'] = 'Interlingua (International Auxiliary Language Association)'
language_dict['inc'] = 'Indic (Other)'
language_dict['ind'] = 'Indonesian'
language_dict['ine'] = 'Indo-European (Other)'
language_dict['inh'] = 'Ingush'
language_dict['int'] = 'Interlingua (International Auxiliary Language Association)'
language_dict['ipk'] = 'Inupiaq'
language_dict['ira'] = 'Iranian (Other)'
language_dict['iri'] = 'Irish'
language_dict['iro'] = 'Iroquoian (Other)'
language_dict['ita'] = 'Italian'
language_dict['jav'] = 'Javanese'
language_dict['jbo'] = 'Lojban (Artificial language)'
language_dict['jpn'] = 'Japanese'
language_dict['jpr'] = 'Judeo-Persian'
language_dict['jrb'] = 'Judeo-Arabic'
language_dict['kaa'] = 'Kara-Kalpak'
language_dict['kab'] = 'Kabyle'
language_dict['kac'] = 'Kachin'
language_dict['kal'] = 'Kalâtdlisut'
language_dict['kam'] = 'Kamba'
language_dict['kan'] = 'Kannada'
language_dict['kar'] = 'Karen languages'
language_dict['kas'] = 'Kashmiri'
language_dict['kau'] = 'Kanuri'
language_dict['kaw'] = 'Kawi'
language_dict['kaz'] = 'Kazakh'
language_dict['kbd'] = 'Kabardian'
language_dict['kha'] = 'Khasi'
language_dict['khi'] = 'Khoisan (Other)'
language_dict['khm'] = 'Khmer'
language_dict['kho'] = 'Khotanese'
language_dict['kik'] = 'Kikuyu'
language_dict['kin'] = 'Kinyarwanda'
language_dict['kir'] = 'Kyrgyz'
language_dict['kmb'] = 'Kimbundu'
language_dict['kok'] = 'Konkani'
language_dict['kom'] = 'Komi'
language_dict['kon'] = 'Kongo'
language_dict['kor'] = 'Korean'
language_dict['kos'] = 'Kosraean'
language_dict['kpe'] = 'Kpelle'
language_dict['krc'] = 'Karachay-Balkar'
language_dict['krl'] = 'Karelian'
language_dict['kro'] = 'Kru (Other)'
language_dict['kru'] = 'Kurukh'
language_dict['kua'] = 'Kuanyama'
language_dict['kum'] = 'Kumyk'
language_dict['kur'] = 'Kurdish'
language_dict['kus'] = 'Kusaie'
language_dict['kut'] = 'Kootenai'
language_dict['lad'] = 'Ladino'
language_dict['lah'] = 'Lahndā'
language_dict['lam'] = 'Lamba (Zambia and Congo)'
language_dict['lan'] = 'Occitan (post 1500)'
language_dict['lao'] = 'Lao'
language_dict['lap'] = 'Sami'
language_dict['lat'] = 'Latin'
language_dict['lav'] = 'Latvian'
language_dict['lez'] = 'Lezgian'
language_dict['lim'] = 'Limburgish'
language_dict['lin'] = 'Lingala'
language_dict['lit'] = 'Lithuanian'
language_dict['lol'] = 'Mongo-Nkundu'
language_dict['loz'] = 'Lozi'
language_dict['ltz'] = 'Luxembourgish'
language_dict['lua'] = 'Luba-Lulua'
language_dict['lub'] = 'Luba-Katanga'
language_dict['lug'] = 'Ganda'
language_dict['lui'] = 'Luiseño'
language_dict['lun'] = 'Lunda'
language_dict['luo'] = 'Luo (Kenya and Tanzania)'
language_dict['lus'] = 'Lushai'
language_dict['mac'] = 'Macedonian'
language_dict['mad'] = 'Madurese'
language_dict['mag'] = 'Magahi'
language_dict['mah'] = 'Marshallese'
language_dict['mai'] = 'Maithili'
language_dict['mak'] = 'Makasar'
language_dict['mal'] = 'Malayalam'
language_dict['man'] = 'Mandingo'
language_dict['mao'] = 'Maori'
language_dict['map'] = 'Austronesian (Other)'
language_dict['mar'] = 'Marathi'
language_dict['mas'] = 'Maasai'
language_dict['max'] = 'Manx'
language_dict['may'] = 'Malay'
language_dict['mdf'] = 'Moksha'
language_dict['mdr'] = 'Mandar'
language_dict['men'] = 'Mende'
language_dict['mga'] = 'Irish, Middle (ca. 1100-1550)'
language_dict['mic'] = 'Micmac'
language_dict['min'] = 'Minangkabau'
language_dict['mis'] = 'Miscellaneous languages'
language_dict['mkh'] = 'Mon-Khmer (Other)'
language_dict['mla'] = 'Malagasy'
language_dict['mlg'] = 'Malagasy'
language_dict['mlt'] = 'Maltese'
language_dict['mnc'] = 'Manchu'
language_dict['mni'] = 'Manipuri'
language_dict['mno'] = 'Manobo languages'
language_dict['moh'] = 'Mohawk'
language_dict['mol'] = 'Moldavian'
language_dict['mon'] = 'Mongolian'
language_dict['mos'] = 'Mooré'
language_dict['mul'] = 'Multiple languages'
language_dict['mun'] = 'Munda (Other)'
language_dict['mus'] = 'Creek'
language_dict['mwl'] = 'Mirandese'
language_dict['mwr'] = 'Marwari'
language_dict['myn'] = 'Mayan languages'
language_dict['myv'] = 'Erzya'
language_dict['nah'] = 'Nahuatl'
language_dict['nai'] = 'North American Indian (Other)'
language_dict['nap'] = 'Neapolitan Italian'
language_dict['nau'] = 'Nauru'
language_dict['nav'] = 'Navajo'
language_dict['nbl'] = 'Ndebele (South Africa)'
language_dict['nde'] = 'Ndebele (Zimbabwe)'
language_dict['ndo'] = 'Ndonga'
language_dict['nds'] = 'Low German'
language_dict['nep'] = 'Nepali'
language_dict['new'] = 'Newari'
language_dict['nia'] = 'Nias'
language_dict['nic'] = 'Niger-Kordofanian (Other)'
language_dict['niu'] = 'Niuean'
language_dict['nno'] = 'Norwegian (Nynorsk)'
language_dict['nob'] = 'Norwegian (Bokmål)'
language_dict['nog'] = 'Nogai'
language_dict['non'] = 'Old Norse'
language_dict['nor'] = 'Norwegian'
language_dict['nqo'] = 'N\'Ko'
language_dict['nso'] = 'Northern Sotho'
language_dict['nub'] = 'Nubian languages'
language_dict['nwc'] = 'Newari, Old'
language_dict['nya'] = 'Nyanja'
language_dict['nym'] = 'Nyamwezi'
language_dict['nyn'] = 'Nyankole'
language_dict['nyo'] = 'Nyoro'
language_dict['nzi'] = 'Nzima'
language_dict['oci'] = 'Occitan (post-1500)'
language_dict['oji'] = 'Ojibwa'
language_dict['ori'] = 'Oriya'
language_dict['orm'] = 'Oromo'
language_dict['osa'] = 'Osage'
language_dict['oss'] = 'Ossetic'
language_dict['ota'] = 'Turkish, Ottoman'
language_dict['oto'] = 'Otomian languages'
language_dict['paa'] = 'Papuan (Other)'
language_dict['pag'] = 'Pangasinan'
language_dict['pal'] = 'Pahlavi'
language_dict['pam'] = 'Pampanga'
language_dict['pan'] = 'Panjabi'
language_dict['pap'] = 'Papiamento'
language_dict['pau'] = 'Palauan'
language_dict['peo'] = 'Old Persian (ca. 600-400 B.C.)'
language_dict['per'] = 'Persian'
language_dict['phi'] = 'Philippine (Other)'
language_dict['phn'] = 'Phoenician'
language_dict['pli'] = 'Pali'
language_dict['pol'] = 'Polish'
language_dict['pon'] = 'Pohnpeian'
language_dict['por'] = 'Portuguese'
language_dict['pra'] = 'Prakrit languages'
language_dict['pro'] = 'Provençal (to 1500)'
language_dict['pus'] = 'Pushto'
language_dict['que'] = 'Quechua'
language_dict['raj'] = 'Rajasthani'
language_dict['rap'] = 'Rapanui'
language_dict['rar'] = 'Rarotongan'
language_dict['roa'] = 'Romance (Other)'
language_dict['roh'] = 'Raeto-Romance'
language_dict['rom'] = 'Romani'
language_dict['rum'] = 'Romanian'
language_dict['run'] = 'Rundi'
language_dict['rup'] = 'Aromanian'
language_dict['rus'] = 'Russian'
language_dict['sad'] = 'Sandawe'
language_dict['sag'] = 'Sango (Ubangi Creole)'
language_dict['sah'] = 'Yakut'
language_dict['sai'] = 'South American Indian (Other)'
language_dict['sal'] = 'Salishan languages'
language_dict['sam'] = 'Samaritan Aramaic'
language_dict['san'] = 'Sanskrit'
language_dict['sao'] = 'Samoan'
language_dict['sas'] = 'Sasak'
language_dict['sat'] = 'Santali'
language_dict['scc'] = 'Serbian'
language_dict['scn'] = 'Sicilian Italian'
language_dict['sco'] = 'Scots'
language_dict['scr'] = 'Croatian'
language_dict['sel'] = 'Selkup'
language_dict['sem'] = 'Semitic (Other)'
language_dict['sga'] = 'Irish, Old (to 1100)'
language_dict['sgn'] = 'Sign languages'
language_dict['shn'] = 'Shan'
language_dict['sho'] = 'Shona'
language_dict['sid'] = 'Sidamo'
language_dict['sin'] = 'Sinhalese'
language_dict['sio'] = 'Siouan (Other)'
language_dict['sit'] = 'Sino-Tibetan (Other)'
language_dict['sla'] = 'Slavic (Other)'
language_dict['slo'] = 'Slovak'
language_dict['slv'] = 'Slovenian'
language_dict['sma'] = 'Southern Sami'
language_dict['sme'] = 'Northern Sami'
language_dict['smi'] = 'Sami'
language_dict['smj'] = 'Lule Sami'
language_dict['smn'] = 'Inari Sami'
language_dict['smo'] = 'Samoan'
language_dict['sms'] = 'Skolt Sami'
language_dict['sna'] = 'Shona'
language_dict['snd'] = 'Sindhi'
language_dict['snh'] = 'Sinhalese'
language_dict['snk'] = 'Soninke'
language_dict['sog'] = 'Sogdian'
language_dict['som'] = 'Somali'
language_dict['son'] = 'Songhai'
language_dict['sot'] = 'Sotho'
language_dict['spa'] = 'Spanish'
language_dict['srd'] = 'Sardinian'
language_dict['srn'] = 'Sranan'
language_dict['srp'] = 'Serbian'
language_dict['srr'] = 'Serer'
language_dict['ssa'] = 'Nilo-Saharan (Other)'
language_dict['sso'] = 'Sotho'
language_dict['ssw'] = 'Swazi'
language_dict['suk'] = 'Sukuma'
language_dict['sun'] = 'Sundanese'
language_dict['sus'] = 'Susu'
language_dict['sux'] = 'Sumerian'
language_dict['swa'] = 'Swahili'
language_dict['swe'] = 'Swedish'
language_dict['swz'] = 'Swazi'
language_dict['syc'] = 'Syriac'
language_dict['syr'] = 'Syriac, Modern'
language_dict['tag'] = 'Tagalog'
language_dict['tah'] = 'Tahitian'
language_dict['tai'] = 'Tai (Other)'
language_dict['taj'] = 'Tajik'
language_dict['tam'] = 'Tamil'
language_dict['tar'] = 'Tatar'
language_dict['tat'] = 'Tatar'
language_dict['tel'] = 'Telugu'
language_dict['tem'] = 'Temne'
language_dict['ter'] = 'Terena'
language_dict['tet'] = 'Tetum'
language_dict['tgk'] = 'Tajik'
language_dict['tgl'] = 'Tagalog'
language_dict['tha'] = 'Thai'
language_dict['tib'] = 'Tibetan'
language_dict['tig'] = 'Tigré'
language_dict['tir'] = 'Tigrinya'
language_dict['tiv'] = 'Tiv'
language_dict['tkl'] = 'Tokelauan'
language_dict['tlh'] = 'Klingon (Artificial language)'
language_dict['tli'] = 'Tlingit'
language_dict['tmh'] = 'Tamashek'
language_dict['tog'] = 'Tonga (Nyasa)'
language_dict['ton'] = 'Tongan'
language_dict['tpi'] = 'Tok Pisin'
language_dict['tru'] = 'Truk'
language_dict['tsi'] = 'Tsimshian'
language_dict['tsn'] = 'Tswana'
language_dict['tso'] = 'Tsonga'
language_dict['tsw'] = 'Tswana'
language_dict['tuk'] = 'Turkmen'
language_dict['tum'] = 'Tumbuka'
language_dict['tup'] = 'Tupi languages'
language_dict['tur'] = 'Turkish'
language_dict['tut'] = 'Altaic (Other)'
language_dict['tvl'] = 'Tuvaluan'
language_dict['twi'] = 'Twi'
language_dict['tyv'] = 'Tuvinian'
language_dict['udm'] = 'Udmurt'
language_dict['uga'] = 'Ugaritic'
language_dict['uig'] = 'Uighur'
language_dict['ukr'] = 'Ukrainian'
language_dict['umb'] = 'Umbundu'
language_dict['und'] = 'Undetermined'
language_dict['urd'] = 'Urdu'
language_dict['uzb'] = 'Uzbek'
language_dict['vai'] = 'Vai'
language_dict['ven'] = 'Venda'
language_dict['vie'] = 'Vietnamese'
language_dict['vol'] = 'Volapük'
language_dict['vot'] = 'Votic'
language_dict['wak'] = 'Wakashan languages'
language_dict['wal'] = 'Wolayta'
language_dict['war'] = 'Waray'
language_dict['was'] = 'Washoe'
language_dict['wel'] = 'Welsh'
language_dict['wen'] = 'Sorbian (Other)'
language_dict['wln'] = 'Walloon'
language_dict['wol'] = 'Wolof'
language_dict['xal'] = 'Oirat'
language_dict['xho'] = 'Xhosa'
language_dict['yao'] = 'Yao (Africa)'
language_dict['yap'] = 'Yapese'
language_dict['yid'] = 'Yiddish'
language_dict['yor'] = 'Yoruba'
language_dict['ypk'] = 'Yupik languages'
language_dict['zap'] = 'Zapotec'
language_dict['zbl'] = 'Blissymbolics'
language_dict['zen'] = 'Zenaga'
language_dict['zha'] = 'Zhuang'
language_dict['znd'] = 'Zande languages'
language_dict['zul'] = 'Zulu'
language_dict['zun'] = 'Zuni'
#language_dict['zxx'] = 'No linguistic content'
language_dict['zza'] = 'Zaza'

author_dict = {}
author_dict['abr'] = 'abridger'
author_dict['acp'] = 'art copyist'
author_dict['act'] = 'actor'
author_dict['adi'] = 'art director'
author_dict['adp'] = 'adapter'
author_dict['aft'] = 'author of afterword, colophon, etc.'
author_dict['anl'] = 'analyst'
author_dict['anm'] = 'animator'
author_dict['ann'] = 'annotator'
author_dict['ant'] = 'bibliographic antecedent'
author_dict['ape'] = 'appellee'
author_dict['apl'] = 'appellant'
author_dict['app'] = 'applicant'
author_dict['aqt'] = 'author in quotations or text abstracts'
author_dict['arc'] = 'architect'
author_dict['ard'] = 'artistic director'
author_dict['arr'] = 'arranger'
author_dict['art'] = 'artist'
author_dict['asg'] = 'assignee'
author_dict['asn'] = 'associated name'
author_dict['ato'] = 'autographer'
author_dict['att'] = 'attributed name'
author_dict['auc'] = 'auctioneer'
author_dict['aud'] = 'author of dialog'
author_dict['aui'] = 'author of introduction, etc.'
author_dict['aus'] = 'screenwriter'
author_dict['aut'] = 'author'
author_dict['bdd'] = 'binding designer'
author_dict['bjd'] = 'bookjacket designer'
author_dict['bkd'] = 'book designer'
author_dict['bkp'] = 'book producer'
author_dict['blw'] = 'blurb writer'
author_dict['bnd'] = 'binder'
author_dict['bpd'] = 'bookplate designer'
author_dict['brd'] = 'broadcaster'
author_dict['brl'] = 'braille embosser'
author_dict['bsl'] = 'bookseller'
author_dict['cas'] = 'caster'
author_dict['ccp'] = 'conceptor'
author_dict['chr'] = 'choreographer'
author_dict['clb'] = 'contributor'
author_dict['cli'] = 'client'
author_dict['cll'] = 'calligrapher'
author_dict['clr'] = 'colorist'
author_dict['clt'] = 'collotyper'
author_dict['cmm'] = 'commentator'
author_dict['cmp'] = 'composer'
author_dict['cmt'] = 'compositor'
author_dict['cnd'] = 'conductor'
author_dict['cng'] = 'cinematographer'
author_dict['cns'] = 'censor'
author_dict['coe'] = 'contestant-appellee'
author_dict['col'] = 'collector'
author_dict['com'] = 'compiler'
author_dict['con'] = 'conservator'
author_dict['cor'] = 'collection registrar'
author_dict['cos'] = 'contestant'
author_dict['cot'] = 'contestant-appellant'
author_dict['cou'] = 'court governed'
author_dict['cov'] = 'cover designer'
author_dict['cpc'] = 'copyright claimant'
author_dict['cpe'] = 'complainant-appellee'
author_dict['cph'] = 'copyright holder'
author_dict['cpl'] = 'complainant'
author_dict['cpt'] = 'complainant-appellant'
author_dict['cre'] = 'creator'
author_dict['crp'] = 'correspondent'
author_dict['crr'] = 'corrector'
author_dict['crt'] = 'court reporter'
author_dict['csl'] = 'consultant'
author_dict['csp'] = 'consultant to a project'
author_dict['cst'] = 'costume designer'
author_dict['ctb'] = 'contributor'
author_dict['cte'] = 'contestee-appellee'
author_dict['ctg'] = 'cartographer'
author_dict['ctr'] = 'contractor'
author_dict['cts'] = 'contestee'
author_dict['ctt'] = 'contestee-appellant'
author_dict['cur'] = 'curator'
author_dict['cwt'] = 'commentator for written text'
author_dict['dbp'] = 'distribution place'
author_dict['dfd'] = 'defendant'
author_dict['dfe'] = 'defendant-appellee'
author_dict['dft'] = 'defendant-appellant'
author_dict['dgg'] = 'degree granting institution'
author_dict['dgs'] = 'degree supervisor'
author_dict['dis'] = 'dissertant'
author_dict['dln'] = 'delineator'
author_dict['dnc'] = 'dancer'
author_dict['dnr'] = 'donor'
author_dict['dpc'] = 'depicted'
author_dict['dpt'] = 'depositor'
author_dict['drm'] = 'draftsman'
author_dict['drt'] = 'director'
author_dict['dsr'] = 'designer'
author_dict['dst'] = 'distributor'
author_dict['dtc'] = 'data contributor'
author_dict['dte'] = 'dedicatee'
author_dict['dtm'] = 'data manager'
author_dict['dto'] = 'dedicator'
author_dict['dub'] = 'dubious author'
author_dict['edc'] = 'editor of compilation'
author_dict['edm'] = 'editor of moving image work'
author_dict['edt'] = 'editor'
author_dict['ed'] = 'editor'
author_dict['ед'] = 'editor'
author_dict['egr'] = 'engraver'
author_dict['elg'] = 'electrician'
author_dict['elt'] = 'electrotyper'
author_dict['eng'] = 'engineer'
author_dict['enj'] = 'enacting jurisdiction'
author_dict['etr'] = 'etcher'
author_dict['evp'] = 'event place'
author_dict['exp'] = 'expert'
author_dict['fac'] = 'facsimilist'
author_dict['fds'] = 'film distributor'
author_dict['fld'] = 'field director'
author_dict['flm'] = 'film editor'
author_dict['fmd'] = 'film director'
author_dict['fmk'] = 'filmmaker'
author_dict['fmo'] = 'former owner'
author_dict['fmp'] = 'film producer'
author_dict['fnd'] = 'funder'
author_dict['fpy'] = 'first party'
author_dict['frg'] = 'forger'
author_dict['gis'] = 'geographic information specialist'
author_dict['grt'] = 'artist'
author_dict['his'] = 'host institution'
author_dict['hnr'] = 'honoree'
author_dict['hst'] = 'host'
author_dict['ill'] = 'illustrator'
author_dict['ilu'] = 'illuminator'
author_dict['ins'] = 'inscriber'
author_dict['inv'] = 'inventor'
author_dict['isb'] = 'issuing body'
author_dict['itr'] = 'instrumentalist'
author_dict['ive'] = 'interviewee'
author_dict['ivr'] = 'interviewer'
author_dict['jud'] = 'judge'
author_dict['jug'] = 'jurisdiction governed'
author_dict['lbr'] = 'laboratory'
author_dict['lbt'] = 'librettist'
author_dict['ldr'] = 'laboratory director'
author_dict['led'] = 'lead'
author_dict['lee'] = 'libelee-appellee'
author_dict['lel'] = 'libelee'
author_dict['len'] = 'lender'
author_dict['let'] = 'libelee-appellant'
author_dict['lgd'] = 'lighting designer'
author_dict['lie'] = 'libelant-appellee'
author_dict['lil'] = 'libelant'
author_dict['lit'] = 'libelant-appellant'
author_dict['lsa'] = 'landscape architect'
author_dict['lse'] = 'licensee'
author_dict['lso'] = 'licensor'
author_dict['ltg'] = 'lithographer'
author_dict['lyr'] = 'lyricist'
author_dict['mcp'] = 'music copyist'
author_dict['mdc'] = 'metadata contact'
author_dict['med'] = 'medium'
author_dict['mfp'] = 'manufacture place'
author_dict['mfr'] = 'manufacturer'
author_dict['mod'] = 'moderator'
author_dict['mon'] = 'monitor'
author_dict['mrb'] = 'marbler'
author_dict['mrk'] = 'markup editor'
author_dict['msd'] = 'musical director'
author_dict['mte'] = 'metal-engraver'
author_dict['mtk'] = 'minute taker'
author_dict['mus'] = 'musician'
author_dict['nrt'] = 'narrator'
author_dict['opn'] = 'opponent'
author_dict['org'] = 'originator'
author_dict['orm'] = 'organizer'
author_dict['osp'] = 'onscreen presenter'
author_dict['oth'] = 'other'
author_dict['own'] = 'owner'
author_dict['pan'] = 'panelist'
author_dict['pat'] = 'patron'
author_dict['pbd'] = 'publishing director'
author_dict['pbl'] = 'publisher'
author_dict['pdr'] = 'project director'
author_dict['pfr'] = 'proofreader'
author_dict['pht'] = 'photographer'
author_dict['plt'] = 'platemaker'
author_dict['pma'] = 'permitting agency'
author_dict['pmn'] = 'production manager'
author_dict['pop'] = 'printer of plates'
author_dict['ppm'] = 'papermaker'
author_dict['ppt'] = 'puppeteer'
author_dict['pra'] = 'praeses'
author_dict['prc'] = 'process contact'
author_dict['prd'] = 'production personnel'
author_dict['pre'] = 'presenter'
author_dict['prf'] = 'performer'
author_dict['prg'] = 'programmer'
author_dict['prm'] = 'printmaker'
author_dict['prn'] = 'production company'
author_dict['pro'] = 'producer'
author_dict['prp'] = 'production place'
author_dict['prs'] = 'production designer'
author_dict['prt'] = 'printer'
author_dict['prv'] = 'provider'
author_dict['pta'] = 'patent applicant'
author_dict['pte'] = 'plaintiff-appellee'
author_dict['ptf'] = 'plaintiff'
author_dict['pth'] = 'patent holder'
author_dict['ptt'] = 'plaintiff-appellant'
author_dict['pup'] = 'publication place'
author_dict['rbr'] = 'rubricator'
author_dict['rcd'] = 'recordist'
author_dict['rce'] = 'recording engineer'
author_dict['rcp'] = 'addressee'
author_dict['rdd'] = 'radio director'
author_dict['red'] = 'redaktor'
author_dict['ren'] = 'renderer'
author_dict['res'] = 'researcher'
author_dict['rev'] = 'reviewer'
author_dict['rpc'] = 'radio producer'
author_dict['rps'] = 'repository'
author_dict['rpt'] = 'reporter'
author_dict['rpy'] = 'responsible party'
author_dict['rse'] = 'respondent-appellee'
author_dict['rsg'] = 'restager'
author_dict['rsp'] = 'respondent'
author_dict['rsr'] = 'restorationist'
author_dict['rst'] = 'respondent-appellant'
author_dict['rth'] = 'research team head'
author_dict['rtm'] = 'research team member'
author_dict['sad'] = 'scientific advisor'
author_dict['sce'] = 'scenarist'
author_dict['scl'] = 'sculptor'
author_dict['scr'] = 'scribe'
author_dict['sds'] = 'sound designer'
author_dict['sec'] = 'secretary'
author_dict['sgd'] = 'stage director'
author_dict['sgn'] = 'signer'
author_dict['sht'] = 'supporting host'
author_dict['sll'] = 'seller'
author_dict['sng'] = 'singer'
author_dict['spk'] = 'speaker'
author_dict['spn'] = 'sponsor'
author_dict['spy'] = 'second party'
author_dict['srv'] = 'surveyor'
author_dict['std'] = 'set designer'
author_dict['stg'] = 'setting'
author_dict['stl'] = 'storyteller'
author_dict['stm'] = 'stage manager'
author_dict['stn'] = 'standards body'
author_dict['str'] = 'stereotyper'
author_dict['tcd'] = 'technical director'
author_dict['tch'] = 'teacher'
author_dict['ths'] = 'thesis advisor'
author_dict['tld'] = 'television director'
author_dict['tlp'] = 'television producer'
author_dict['trc'] = 'transcriber'
author_dict['trl'] = 'translator'
author_dict['tr'] = 'translator'
author_dict['tyd'] = 'type designer'
author_dict['tyg'] = 'typographer'
author_dict['uvp'] = 'university place'
author_dict['vac'] = 'voice actor'
author_dict['vdg'] = 'videographer'
author_dict['voc'] = 'singer'
author_dict['wac'] = 'writer of added commentary'
author_dict['wal'] = 'writer of added lyrics'
author_dict['wam'] = 'writer of accompanying material'
author_dict['wat'] = 'writer of added text'
author_dict['wdc'] = 'woodcutter'
author_dict['wde'] = 'wood engraver'
author_dict['win'] = 'writer of introduction'
author_dict['wit'] = 'witness'
author_dict['wpr'] = 'writer of preface'
author_dict['wst'] = 'writer of supplementary textual content'

collection_dict = {}
collection_dict['monograph'] = {'coverage_depth' : 'ebook', 'oclc_collection_name' : 'Center for Research Libraries (CRL) eResources, Monographs', 'oclc_collection_id' : 'customer.93175.5'}
collection_dict['newspaper'] = {'coverage_depth' : 'fulltext', 'oclc_collection_name' : 'Center for Research Libraries (CRL) eResources, Newspapers', 'oclc_collection_id' : 'customer.93175.10'}
collection_dict['serial'] = {'coverage_depth' : 'fulltext', 'oclc_collection_name' : 'Center for Research Libraries (CRL) eResources, Serials', 'oclc_collection_id' : 'customer.93175.8'}

file_types_dict = {'csv' : '.csv', 'tsv' : '.tsv', 'excel' : '.xlsx', 'unicode text' : '.txt'}

file_types = [('csv file', '.csv'), ('tsv file', '.tsv'), ('excel file', '.xlsx'), ('unicode text', '.txt')]

record_source_dict = {}
record_source_dict['millennium'] = {'button' : 'Millennium', 'label' : 'Bib number:\t'}
record_source_dict['worldcat'] = {'button' : 'Worldcat', 'label' : 'OCLC number:\t'}
record_source_dict['folio'] = {'button' : 'FOLIO', 'label' : ''}
record_source_dict['ddsnext'] = {'button' : 'DDSnext', 'label' : ''}
record_source_dict['eastview'] = {'button' : 'EastView', 'label' : ''}
record_source_dict['icon'] = {'button' : 'ICON', 'label' : ''}
record_source_dict['recap'] = {'button' : 'ReCAP', 'label' : ''}
record_source_dict['papr'] = {'button' : 'PAPR', 'label' : ''}

class excel_writer:
    def __init__(self, workbook, filename, name):
        self.row = 1
        self.column = 1
        self.workbook = workbook
        self.filename = filename
        self.workbook_sheet = workbook.active
        self.workbook_sheet.title = name
    #Append to row.
    def writerow(self, row_content):
#        print(row_content)
        for i in range(1, len(row_content) + 1):
            if row_content[i - 1] is 'None' or row_content[i - 1] is None:
                row_content[i - 1] = ''
#            print(row_content[i - 1])
            self.workbook_sheet.cell(row = self.row, column = i, value=row_content[i - 1])
        self.row += 1
    def save(self):
        self.workbook.save(filename = self.filename)
#        print(self.filename)
#        self.workbook.close()

class configuration:
    def __init__(self):
        self.config_folder = os.path.join(os.path.join(os.path.join(os.path.join(os.path.join('C:\\Users', os.getlogin()), 'AppData'), 'Local'), 'CRL'), 'Metadata')
        self.config = configparser.ConfigParser()
        self.config_file = os.path.join(self.config_folder, 'metadata_extraction_config.ini')
        self.config_data = {}
        self.read_config_file()
    #
    def read_config_file(self):
        check_or_create_dir(self.config_folder)
        # create a blank file if none exists
        if not os.path.isfile(self.config_file):
            self.write_config_file()
        self.config.read(self.config_file)
    #
    def write_config_file(self):
        with open(self.config_file, 'w') as config_out:
            self.config.write(config_out)
    #Check to see if a section with a given name exist.
    def section_exist(self, section):
        if section in self.config:
            return True
        return False
    
    def add_section(self, section):
        if not self.section_exist(section):
            self.config[section] = None
            self.write_config_file()
    
    def add_template_location(self, section, file_type):
        if not self.section_exist(section):
            self.config[section] = {}
        if file_type not in self.config[section]:
            self.config[section][file_type] = section + '.' + file_type
            self.config[self.config[section][file_type]] = {}
            self.write_config_file()
    #Modifies the file location in the config file.
    def modify_file_location(self, section, file_type=None, folder_name=None, file_name=None, file_extention=None, file_location=None):
        #
        if file_location is None:
            file_dict = OrderedDict({'folder_name' : folder_name, 'file_name' : file_name, 'file_extention' : file_extention})
        #
        else:
            file_dict = OrderedDict({'folder_name' : os.path.dirname(file_location), 'file_name' : os.path.splitext(os.path.basename(file_location))[0], 'file_extention' : os.path.splitext(os.path.basename(file_location))[1]})
#            print(file_location, file_dict)
        if not self.section_exist(section) or file_type not in self.config[section]:
            self.add_template_location(section, file_type)
            self.config[self.config[section][file_type]] = file_dict
            self.write_config_file()
        else:
            self.config[self.config[section][file_type]] = file_dict
#            print(self.config[self.config[section][file_type]])
            self.write_config_file()
    
    #Gets file location.
    def get_file_location(self, section, file_type):
        if self.section_exist(section) and self.config[section][file_type] in self.config:
            return os.path.join(self.config[self.config[section][file_type]]['folder_name'], self.config[self.config[section][file_type]]['file_name'] + self.config[self.config[section][file_type]]['file_extention'])
        return None


class export_dialog(tkinter.Toplevel):
    def __init__(self, parent, title = None):
        tkinter.Toplevel.__init__(self, parent)
        self.transient(parent)
        if title:
            self.title(title)
        self.parent = parent
        self.result = None
        self.title('Export')
        self.resizable(width='FALSE', height='FALSE')
        
        self.config = self.parent.config
        
        self.entry_descriptive_metadata_text = tkinter.StringVar()
        self.entry_title_metadata_text = tkinter.StringVar()
        self.entry_kbart_metadata_text = tkinter.StringVar()
        
        body = tkinter.Frame(self)
        self.initial_focus = self.body(body)
        self.grab_set()
        if not self.initial_focus:
            self.initial_focus = self
        self.protocol("WM_DELETE_WINDOW", self.teminate)
        self.geometry("+%d+%d" % (parent.winfo_rootx()+50, parent.winfo_rooty()+50))
        self.initial_focus.focus_set()
        self.wait_window(self)
    def body(self, master):
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.columnconfigure(1, weight = 10)
        self.rowconfigure(4, weight = 10)
        self.top_frame = tkinter.Frame(self)
        self.top_frame.grid(row=0, sticky="nwe")
        
        self.descriptive_metadata_label = tkinter.Label(self.top_frame, text='Descriptive Metadata File:\t', font=font.Font(size=self.parent.size))
        self.title_metadata_label = tkinter.Label(self.top_frame, text='Title Metadata File:\t', font=font.Font(size=self.parent.size))
        self.kbart_metadata_label = tkinter.Label(self.top_frame, text='Kbart Metadata File:\t', font=font.Font(size=self.parent.size))
        
        self.descriptive_metadata_button = tkinter.Button(self.top_frame, text='Browse', font=font.Font(size=self.parent.size), command=self.open_descriptive_metadata_file)
        self.title_metadata_button = tkinter.Button(self.top_frame, text='Browse', font=font.Font(size=self.parent.size), command=self.open_title_metadata_file)
        self.kbart_metadata_button = tkinter.Button(self.top_frame, text='Browse', font=font.Font(size=self.parent.size), command=self.open_kbart_metadata_file)
        
        self.save_button = tkinter.Button(self.top_frame, text='Export', font=font.Font(size=self.parent.size), command=self.export)
        self.close_button = tkinter.Button(self.top_frame, text='Close', font=font.Font(size=self.parent.size), command=self.destroy)
        
        self.entry_descriptive_metadata = tkinter.Entry(self.top_frame, textvariable=self.entry_descriptive_metadata_text, font=font.Font(size=self.parent.size))
        self.entry_title_metadata = tkinter.Entry(self.top_frame, textvariable=self.entry_title_metadata_text, font=font.Font(size=self.parent.size))
        self.entry_kbart_metadata = tkinter.Entry(self.top_frame, textvariable=self.entry_kbart_metadata_text, font=font.Font(size=self.parent.size))
        
        self.descriptive_metadata_label.grid(row=0, column=0, sticky='N'+'W')
        self.title_metadata_label.grid(row=1, column=0, sticky='N'+'W')
        self.kbart_metadata_label.grid(row=2, column=0, sticky='N'+'W')
        
        self.entry_descriptive_metadata.grid(row=0, column=1, columnspan=20, sticky='N'+'E'+'W')
        self.entry_title_metadata.grid(row=1, column=1, columnspan=20, sticky='N'+'E'+'W')
        self.entry_kbart_metadata.grid(row=2, column=1, columnspan=20, sticky='N'+'E'+'W')
        
        self.descriptive_metadata_button.grid(row=0, column=21, sticky='N'+'E'+'W')
        self.title_metadata_button.grid(row=1, column=21, sticky='N'+'E'+'W')
        self.kbart_metadata_button.grid(row=2, column=21, sticky='N'+'E'+'W')
        
        self.save_button.grid(row=4, column=20)
        self.close_button.grid(row=4, column=21)
        self.grid_rowconfigure(0, weight=1)
        col_count, row_count = self.top_frame.grid_size()
        for col in range(col_count):
            self.top_frame.grid_columnconfigure(col, minsize=20)
        for row in range(row_count):
            self.top_frame.grid_rowconfigure(row, minsize=20)
    def export(self):
        if self.entry_descriptive_metadata_text.get() != '':
#            descriptive_metadata_file = open(os.path.join(output_folder, 'ddsnext_descriptive_metadata_' + get_date() + file_types_dict[self.parent.file_types]),'wt',encoding='utf8', newline='')
#            print(self.config.get_file_location('file_locations', 'descriptive_metadata'))
#            descriptive_metadata_file = open(self.config.get_file_location('file_locations', 'descriptive_metadata'),'wt',encoding='utf8', newline='')
            if self.config.config['file_locations.descriptive_metadata']['file_extention'] == '.xlsx':
                descriptive_metadata_file = openpyxl.Workbook()
                dds_descriptive_writer = excel_writer(descriptive_metadata_file, self.config.get_file_location('file_locations', 'descriptive_metadata'), 'descriptive_metadata')
            elif self.config.config['file_locations.descriptive_metadata']['file_extention'] == '.txt':
                descriptive_metadata_file = open(self.config.get_file_location('file_locations', 'descriptive_metadata'),'wt',encoding='utf-16-le', newline='')
                dds_descriptive_writer = csv.writer(descriptive_metadata_file, delimiter='\t', dialect='excel-tab')
            else:
#            if file_types_dict[self.parent.file_types] == '.tsv':
                descriptive_metadata_file = open(self.config.get_file_location('file_locations', 'descriptive_metadata'),'wt',encoding='utf8', newline='')
                if self.config.config['file_locations.descriptive_metadata']['file_extention'] == '.tsv':
                    dds_descriptive_writer = csv.writer(descriptive_metadata_file, delimiter='\t', dialect='excel')
                else:
                    dds_descriptive_writer = csv.writer(descriptive_metadata_file, dialect='excel')
#            dds_descriptive_writer.writerow(['title_uuid' , 'field' , 'value'])
#            print(descriptive_items)
#            print(self.parent.descriptive_metadata)
#            for key in self.parent.descriptive_metadata:
#                for descriptive_metadata_field in self.parent.descriptive_metadata[key]:
##                    print(self.parent.descriptive_metadata)
##                    print(self.parent.descriptive_metadata[key])
##                    print(descriptive_metadata_field)
##                    print(self.parent.descriptive_metadata[key]['title_uuid'], self.parent.descriptive_metadata[key]['field'], self.parent.descriptive_metadata[key]['value'])
##                    dds_descriptive_writer.writerow([descriptive_metadata_field['title_uuid'], descriptive_metadata_field['field'], descriptive_metadata_field['value']])
##                    dds_descriptive_writer.writerow([self.parent.descriptive_metadata[key]['title_uuid'], self.parent.descriptive_metadata[key]['field'], self.parent.descriptive_metadata[key]['value']])
#                    dds_descriptive_writer.writerow([self.parent.descriptive_metadata[key][descriptive_metadata_field]['title_uuid'], self.parent.descriptive_metadata[key][descriptive_metadata_field]['field'], self.parent.descriptive_metadata[key][descriptive_metadata_field]['value']])

#            for key in self.parent.descriptive_metadata:
#                dds_descriptive_writer.writerow([self.parent.descriptive_metadata[key]['title_uuid'], self.parent.descriptive_metadata[key]['field'], self.parent.descriptive_metadata[key]['value']])

            for row in self.parent.template_spreadsheet['descriptive_metadata']:
                dds_descriptive_writer.writerow([self.parent.template_spreadsheet['descriptive_metadata'][row]['title_uuid'].get(), self.parent.template_spreadsheet['descriptive_metadata'][row]['field'].get(), self.parent.template_spreadsheet['descriptive_metadata'][row]['value'].get()])

            if self.config.config['file_locations.descriptive_metadata']['file_extention'] == '.xlsx':
                dds_descriptive_writer.save()
            descriptive_metadata_file.close()
        
        if self.entry_title_metadata_text.get() != '':
#            title_metadata_file = open(os.path.join(output_folder, 'ddsnext_title_metadata_' + get_date() + file_types_dict[self.parent.file_types]),'wt',encoding='utf8', newline='')
#            print(self.config.get_file_location('file_locations', 'title_metadata'))
            
#            title_metadata_file = open(self.config.get_file_location('file_locations', 'title_metadata'),'wt',encoding='utf8', newline='')
            
#            if file_types_dict[self.parent.file_types] == '.tsv':
            if self.config.config['file_locations.title_metadata']['file_extention'] == '.xlsx':
#                workbook = openpyxl.Workbook()
#                title_metadata_sheet = workbook.active
#                title_metadata_sheet.title  = 'title_metadata'
                title_metadata_file = openpyxl.Workbook()
                dds_title_writer = excel_writer(title_metadata_file, self.config.get_file_location('file_locations', 'title_metadata'), 'title_metadata')
                
#                dds_title_writer = csv.writer(title_metadata_file, delimiter='\t', dialect='excel')
            elif self.config.config['file_locations.title_metadata']['file_extention'] == '.txt':
                title_metadata_file = open(self.config.get_file_location('file_locations', 'title_metadata'),'wt',encoding='utf-16-le', newline='')
                dds_title_writer = csv.writer(title_metadata_file, delimiter='\t', dialect='excel-tab')
            else:
                title_metadata_file = open(self.config.get_file_location('file_locations', 'title_metadata'),'wt',encoding='utf8', newline='')
                if self.config.config['file_locations.title_metadata']['file_extention'] == '.tsv':
                    dds_title_writer = csv.writer(title_metadata_file, delimiter='\t', dialect='excel')
                else:
                    dds_title_writer = csv.writer(title_metadata_file, dialect='excel')
#            dds_title_writer.writerow(['Title Name', 'Title UUID', 'Title Material Type', 'Title Format', 'Title OCLC', 'Title Digital Holding Range', 'Title Resolution', 'Title Color Depth', 'Title Location Code', 'Title Catalog Link', 'Title External Link'])
#            print(self.parent.title_metadata)

#            for key in self.parent.title_metadata:
##                print(key, self.parent.title_metadata[key])
#                dds_title_writer.writerow([self.parent.title_metadata[key]['title_name'], self.parent.title_metadata[key]['title_uuid'], self.parent.title_metadata[key]['title_material_type'], self.parent.title_metadata[key]['title_format'], self.parent.title_metadata[key]['title_oclc'], self.parent.title_metadata[key]['title_digital_holding_range'], self.parent.title_metadata[key]['title_resolution'], self.parent.title_metadata[key]['title_color_depth'], self.parent.title_metadata[key]['title_location_code'], self.parent.title_metadata[key]['title_catalog_link'], self.parent.title_metadata[key]['title_external_link']])

            for row in self.parent.template_spreadsheet['title_metadata']:
                dds_title_writer.writerow([self.parent.template_spreadsheet['title_metadata'][row]['title_name'].get(), self.parent.template_spreadsheet['title_metadata'][row]['title_uuid'].get(), self.parent.template_spreadsheet['title_metadata'][row]['title_material_type'].get(), self.parent.template_spreadsheet['title_metadata'][row]['title_format'].get(), self.parent.template_spreadsheet['title_metadata'][row]['title_oclc'].get(), self.parent.template_spreadsheet['title_metadata'][row]['title_digital_holding_range'].get(), self.parent.template_spreadsheet['title_metadata'][row]['title_resolution'].get(), self.parent.template_spreadsheet['title_metadata'][row]['title_color_depth'].get(), self.parent.template_spreadsheet['title_metadata'][row]['title_location_code'].get(), self.parent.template_spreadsheet['title_metadata'][row]['title_catalog_link'].get(), self.parent.template_spreadsheet['title_metadata'][row]['title_external_link'].get()])

            if self.config.config['file_locations.title_metadata']['file_extention'] == '.xlsx':
                dds_title_writer.save()
            title_metadata_file.close()

        if self.entry_kbart_metadata_text.get() != '':
#            kbart_metadata_file = open(os.path.join(output_folder, 'worldshare_metadata_' + get_date() + file_types_dict[self.parent.file_types]),'wt',encoding='utf8', newline='')
#            print(self.config.get_file_location('file_locations', 'kbart_metadata'))
#            kbart_metadata_file = open(self.config.get_file_location('file_locations', 'kbart_metadata'),'wt',encoding='utf8', newline='')
#            if file_types_dict[self.parent.file_types] == '.tsv':
            if self.config.config['file_locations.kbart_metadata']['file_extention'] == '.xlsx':
                kbart_metadata_file = openpyxl.Workbook()
                worldshare_writer = excel_writer(kbart_metadata_file, self.config.get_file_location('file_locations', 'kbart_metadata'), 'kbart_metadata')
            elif self.config.config['file_locations.kbart_metadata']['file_extention'] == '.txt':
                kbart_metadata_file = open(self.config.get_file_location('file_locations', 'kbart_metadata'),'wt',encoding='utf-16-le', newline='')
                worldshare_writer = csv.writer(kbart_metadata_file, delimiter='\t', dialect='excel-tab')
            else:
                kbart_metadata_file = open(self.config.get_file_location('file_locations', 'kbart_metadata'),'wt',encoding='utf8', newline='')
                if self.config.config['file_locations.kbart_metadata']['file_extention'] == '.tsv':
                    worldshare_writer = csv.writer(kbart_metadata_file, delimiter='\t', dialect='excel')
                else:
                    worldshare_writer = csv.writer(kbart_metadata_file, dialect='excel')
#            worldshare_writer.writerow(['publication_title', 'print_identifier', 'online_identifier', 'date_first_issue_online', 'num_first_vol_online', 'num_first_issue_online', 'date_last_issue_online', 'num_last_vol_online', 'num_last_issue_online', 'title_url', 'first_author', 'title_id', 'embargo_info', 'coverage_depth', 'coverage_notes', 'publisher_name', 'location', 'title_notes', 'staff_notes', 'vendor_id', 'oclc_collection_name', 'oclc_collection_id', 'oclc_entry_id', 'oclc_linkscheme', 'oclc_number', 'ACTION'])
#            print(self.parent.kbart_metadata)

#            for key in self.parent.kbart_metadata:
##                print(key)
#                worldshare_writer.writerow([self.parent.kbart_metadata[key]['publication_title'], self.parent.kbart_metadata[key]['print_identifier'], self.parent.kbart_metadata[key]['online_identifier'], self.parent.kbart_metadata[key]['date_first_issue_online'], self.parent.kbart_metadata[key]['num_first_vol_online'], self.parent.kbart_metadata[key]['num_first_issue_online'], self.parent.kbart_metadata[key]['date_last_issue_online'], self.parent.kbart_metadata[key]['num_last_vol_online'], self.parent.kbart_metadata[key]['num_last_issue_online'], self.parent.kbart_metadata[key]['title_url'], self.parent.kbart_metadata[key]['first_author'], self.parent.kbart_metadata[key]['title_id'], self.parent.kbart_metadata[key]['embargo_info'], self.parent.kbart_metadata[key]['coverage_depth'], self.parent.kbart_metadata[key]['coverage_notes'], self.parent.kbart_metadata[key]['publisher_name'], self.parent.kbart_metadata[key]['location'], self.parent.kbart_metadata[key]['title_notes'], self.parent.kbart_metadata[key]['staff_notes'], self.parent.kbart_metadata[key]['vendor_id'], self.parent.kbart_metadata[key]['oclc_collection_name'], self.parent.kbart_metadata[key]['oclc_collection_id'], self.parent.kbart_metadata[key]['oclc_entry_id'], self.parent.kbart_metadata[key]['oclc_linkscheme'], self.parent.kbart_metadata[key]['oclc_number'], self.parent.kbart_metadata[key]['action']])

            for row in self.parent.template_spreadsheet['kbart_metadata']:
                worldshare_writer.writerow([self.parent.template_spreadsheet['kbart_metadata'][row]['publication_title'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['print_identifier'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['online_identifier'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['date_first_issue_online'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['num_first_vol_online'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['num_first_issue_online'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['date_last_issue_online'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['num_last_vol_online'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['num_last_issue_online'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['title_url'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['first_author'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['title_id'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['embargo_info'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['coverage_depth'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['coverage_notes'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['publisher_name'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['location'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['title_notes'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['staff_notes'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['vendor_id'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['oclc_collection_name'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['oclc_collection_id'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['oclc_entry_id'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['oclc_linkscheme'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['oclc_number'].get(), self.parent.template_spreadsheet['kbart_metadata'][row]['action'].get()])

            if self.config.config['file_locations.kbart_metadata']['file_extention'] == '.xlsx':
                worldshare_writer.save()
            kbart_metadata_file.close()

#        descriptive_metadata_file.close()
#        title_metadata_file.close()
#        kbart_metadata_file.close()
        self.teminate()
    #Opens file dialog
    def open_file(self, initialdir, confirmoverwrite, defaultextension, filetypes, initialfile, title):
        return asksaveasfilename(initialdir = initialdir, confirmoverwrite = confirmoverwrite, defaultextension = defaultextension, filetypes = filetypes, initialfile = initialfile, title = title)
    #Opens input file dialog, and prints it to input entry text field
    def open_descriptive_metadata_file(self):
#        self.config.config['file_locations.descriptive_metadata']['folder_name']
        file_location = self.open_file(initialdir = self.config.config['file_locations.descriptive_metadata']['folder_name'], confirmoverwrite = True, defaultextension = self.config.config['file_locations.descriptive_metadata']['file_extention'], filetypes = file_types, initialfile = self.config.config['file_locations.descriptive_metadata']['file_name'], title = 'Descriptive metadata')
        if file_location:
            self.entry_descriptive_metadata_text.set(file_location)
            self.config.modify_file_location('file_locations', 'descriptive_metadata', file_location=file_location)
    #Opens output file dialog, and prints it to output entry text field
    def open_title_metadata_file(self):
        file_location = self.open_file(initialdir = self.config.config['file_locations.title_metadata']['folder_name'], confirmoverwrite = True, defaultextension = self.config.config['file_locations.title_metadata']['file_extention'], filetypes = file_types, initialfile = self.config.config['file_locations.title_metadata']['file_name'], title = 'Title metadata')
        if file_location:
            self.entry_title_metadata_text.set(file_location)
            self.config.modify_file_location('file_locations', 'title_metadata', file_location=file_location)
    #Opens error file dialog, and prints it to error entry text field
    def open_kbart_metadata_file(self):
        file_location = self.open_file(initialdir = self.config.config['file_locations.kbart_metadata']['folder_name'], confirmoverwrite = True, defaultextension = self.config.config['file_locations.kbart_metadata']['file_extention'], filetypes = file_types, initialfile = self.config.config['file_locations.kbart_metadata']['file_name'], title = 'Kbart metadata')
        if file_location:
            self.entry_kbart_metadata_text.set(file_location)
            self.config.modify_file_location('file_locations', 'kbart_metadata', file_location=file_location)
    #Sets focus to parent application and closes dialog
    def teminate(self):
        self.parent.focus_set()
        self.destroy()

def convert_to_dict(input_list, header=False):
    output_dict = {}
#    print(input_list)
    row = 0
    for value in input_list:
        if value == 'header':
            output_dict['header'] = value
            header=False
        else:
            output_dict[row] = value
        row += 1
    return output_dict

class Application(tkinter.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.bib_num = ''
        self.uuid = ''
        self.size = 11
        self.master.title('Metadata Extraction')
        self.master.geometry('1000x700')
        self.grid(row=0, sticky='N'+'S'+'E'+'W')
        self.master.grid_rowconfigure(0, weight = 1)
        self.master.grid_columnconfigure(0, weight = 1)
        self.screen_text = tkinter.StringVar()
        self.entry_catalog_id_text = tkinter.StringVar()
        self.entry_uuid_text = tkinter.StringVar()
        self.entry_page_text = tkinter.StringVar()
        self.number_of_pages = tkinter.StringVar()
        self.number_of_pages.set(0)
        self.catalog_id_text = tkinter.StringVar()
        self.catalog_id_text.set('Bib number:\t')
        self.catalog_id_button_text = tkinter.StringVar()
        self.catalog_id_button_text.set('Millennium')
        self.record_source = 'millennium'
        self.file_types = 'csv'
        self.input_values = []
        self.import_type = 'batch'
        self.records = {}
        self.pages = {}
        self.spreadsheet = {}
        self.template_spreadsheet = {}
        self.template_dict = {}
        self.windows_id_dict = {}
        self.windows = ttk.Notebook(self)
        self.screen_frame = tkinter.Frame(self.windows)
        self.spreadsheet_frame = tkinter.Frame(self.windows)
        self.create_menu()
        self.create_widgets()
        self.create_spreadsheet()
        self.windows.add(self.spreadsheet_frame, text='Input')
        self.windows_id_dict['Input'] = 0
        self.windows.add(self.screen_frame, text='Records')
        self.windows_id_dict['Records'] = 1
        
        self.windows.grid(row=0, column=0, rowspan=8, columnspan=8, sticky='nswe')
        
        self.title_metadata = {}
        self.descriptive_metadata = {}
        self.kbart_metadata = {}
        
        global output_folder
        
        self.config = configuration()
#        self.config.add_section('file_locations')
#        self.config.add_template_location('file_locations', 'descriptive_metadata')
#        self.config.add_template_location('file_locations', 'title_metadata')
#        self.config.add_template_location('file_locations', 'kbart_metadata')
#        self.config.modify_file_location('file_locations', 'descriptive_metadata', output_folder, 'ddsnext_descriptive_metadata_' + get_date(), file_types_dict[self.file_types])
#        self.config.modify_file_location('file_locations', 'title_metadata', output_folder, 'ddsnext_title_metadata_' + get_date(), file_types_dict[self.file_types])
#        self.config.modify_file_location('file_locations', 'kbart_metadata', output_folder, 'worldshare_metadata_' + get_date(), file_types_dict[self.file_types])
        
        self.config.modify_file_location('file_locations', file_type='descriptive_metadata', folder_name=output_folder, file_name='ddsnext_descriptive_metadata_' + get_date(), file_extention=file_types_dict[self.file_types])
        self.config.modify_file_location('file_locations', file_type='title_metadata', folder_name=output_folder, file_name='ddsnext_title_metadata_' + get_date(), file_extention=file_types_dict[self.file_types])
        self.config.modify_file_location('file_locations', file_type='kbart_metadata', folder_name=output_folder, file_name='worldshare_metadata_' + get_date(), file_extention=file_types_dict[self.file_types])
        
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.p = None
        self.screen_p = None
        self.entry_input_text = tkinter.StringVar()
        self.entry_output_text = tkinter.StringVar()
        self.entry_error_text = tkinter.StringVar()
        self.current_record = None
    def create_menu(self):
        self.menubar = tkinter.Menu(self.master)
        self.file_menu = tkinter.Menu(self.menubar, tearoff=0)
        self.file_menu.add_command(label='Run', font=font.Font(size=self.size), command=self.run)
#        self.file_menu.add_command(label='Run', font=font.Font(size=self.size), command=partial(self.run, self.import_type))
#        self.file_menu.add_command(label='Export', state = 'disabled', command=self.export_to_file)
        self.file_menu.add_command(label='Export', font=font.Font(size=self.size), state = 'disabled', command=self.export)
#        self.file_menu.add_command(label='Batch', command=self.change_to_batch)
        self.file_menu.add_command(label='Quit', font=font.Font(size=self.size), command=self.teminate)
        self.menubar.add_cascade(label='File', font=font.Font(size=self.size), menu=self.file_menu)
        
        self.setting_menu = tkinter.Menu(self.menubar, tearoff=0)
        self.setting_menu.add_command(label='Font', font=font.Font(size=self.size), command=self.export)
#        self.setting_menu.add_command(label='Export', font=font.Font(size=self.size), state = 'disabled', command=self.export)
#        self.setting_menu.add_command(label='Quit', font=font.Font(size=self.size), command=self.teminate)
        self.menubar.add_cascade(label='Setting', font=font.Font(size=self.size), menu=self.setting_menu)
        #Displays the menu
        self.master.config(menu=self.menubar)
    def create_widgets(self):
#        self.screen_frame = tkinter.Frame(self)
        self.screen_frame.grid(row=0, column=0, rowspan=8, columnspan=8, sticky='nswe')
#        self.screen_frame.grid_rowconfigure(0, weight=1)
#        self.screen_frame.grid_columnconfigure(0, weight=1) 
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.screen_frame.grid_rowconfigure(1, weight=1)
        self.screen_frame.grid_columnconfigure(0, weight=1)
        
        self.screen_frame_outer = tkinter.Frame(self.screen_frame, background='black')
        #Sets up the screen canvas to enable scrolling.
        self.screen_canvas = tkinter.Canvas(self.screen_frame_outer, background='black', highlightthickness=0)
        self.screen_frame_inner = tkinter.Frame(self.screen_canvas, background='black')
        #Sets up the vertical and horizontal scrollbars.
        self.screen_scrollbar_vertical=tkinter.Scrollbar(self.screen_frame_outer, orient='vertical', command=self.screen_canvas.yview)
        self.screen_scrollbar_horizontal=tkinter.Scrollbar(self.screen_frame_outer, orient='horizontal', command=self.screen_canvas.xview)
        self.scrollbar_visible = False
#        self.screen_scrollbar_vertical.grid(row=0, column=2, rowspan=2, sticky='ns')
#        self.screen_scrollbar_horizontal.grid(row=1, column=0, sticky='we')
        self.screen_label = tkinter.Label(self.screen_frame_inner, background='black', textvariable=self.screen_text, font=font.Font(size=self.size), fg='white', anchor='nw', justify=tkinter.LEFT)
        self.screen_label.grid(row=0, column=0, columnspan=8, rowspan=6, sticky='nswe')
        self.screen_frame_inner.grid_rowconfigure(0, weight=1, minsize=461)
        self.screen_frame_inner.grid_columnconfigure(0, weight=1, minsize=620)
        #Adds space to the frame to prevent the vertical scrollbar from cutting off text.
        self.screen_frame_inner.grid_columnconfigure(9, minsize=15)
        self.screen_frame_inner.grid_rowconfigure(8, weight=1)
        self.screen_frame_inner.grid_columnconfigure(6, weight=1)
        
        #Finishes setting up the main canvas by connecting the scrollbar and adding the frame containing hot folders information.
        self.screen_canvas.configure(yscrollcommand=self.screen_scrollbar_vertical.set, xscrollcommand=self.screen_scrollbar_horizontal.set)
        self.canvas_frame = self.screen_canvas.create_window((0,0), window=self.screen_frame_inner, anchor='nw')
        
        #Binds the canvases to the configure event.
        self.screen_frame_outer.bind('<Configure>', self.set_up_canvas)
        self.screen_frame_outer.bind('<<change_page>>', self.set_up_records)
        
#        self.screen_canvas.grid(row=2, column=0, columnspan=8, rowspan=6, sticky='nswe', padx=6, pady=6)
#        self.screen_canvas.grid(row=2, column=0, columnspan=8, sticky='nswe', padx=6)
        self.screen_frame_outer.grid(row=1, column=0, columnspan=8, sticky='nswe', padx=6)
        self.screen_frame_outer.grid_rowconfigure(0, weight=1)
        self.screen_frame_outer.grid_columnconfigure(0, weight=1)
        
        self.screen_canvas.grid(row=0, column=0, columnspan=8, sticky='nswe', padx=6)
        
#        self.screen_label.grid(row=0, column=0, columnspan=8, rowspan=6, sticky='nswe', padx=6, pady=6)
        
#        self.screen_canvas.grid_rowconfigure(0, weight=1)
#        self.screen_canvas.grid_columnconfigure(0, weight=1)
#        self.screen_canvas.grid_columnconfigure(6, weight=1)
        
##        self.catalog_id_label = tkinter.Label(self.screen_frame, text='Bib number:\t')
#        self.screen_frame_top = tkinter.Frame(self.screen_frame)
#        self.catalog_id_label = tkinter.Label(self.screen_frame_top, textvariable=self.catalog_id_text, font=font.Font(size=self.size))
#        self.uuid_label = tkinter.Label(self.screen_frame_top, text='DDSnext UUID:\t', font=font.Font(size=self.size))
#        self.entry_bib_num = tkinter.Entry(self.screen_frame_top, textvariable=self.entry_catalog_id_text, font=font.Font(size=self.size))
#        self.entry_uuid = tkinter.Entry(self.screen_frame_top, textvariable=self.entry_uuid_text, font=font.Font(size=self.size))
#        self.run_button = tkinter.Button(self.screen_frame_top, text='Run', font=font.Font(size=self.size), command=self.run)
#        self.run_button = tkinter.Button(self.screen_frame_top, text='Run', font=font.Font(size=self.size), command=partial(self.run, 'single'))
#        #Sets up Button to switch between Millennium and Worldcat.
#        self.catalog_id_button = tkinter.Menubutton(self.screen_frame_top, textvariable=self.catalog_id_button_text, font=font.Font(size=self.size), relief=tkinter.RAISED)
#        self.catalog_id_button.grid()
#        self.catalog_id_button.menu = tkinter.Menu(self.catalog_id_button, tearoff=0)
#        self.catalog_id_button['menu'] =  self.catalog_id_button.menu
##        self.catalog_id_button.menu.add_command(label='Millennium', command=self.change_to_millennium)
##        self.catalog_id_button.menu.add_command(label='Worldcat', command=self.change_to_worldcat)
#        self.catalog_id_button.menu.add_command(label='Millennium', command=partial(self.change_record_source, 'millennium'))
#        self.catalog_id_button.menu.add_command(label='Worldcat', command=partial(self.change_record_source, 'worldcat'))
#        self.catalog_id_label.grid(row=0, column=0)
#        self.uuid_label.grid(row=1, column=0)
#        self.entry_bib_num.grid(row=0, column=1, columnspan=6, sticky='nswe')
#        self.entry_uuid.grid(row=1, column=1, columnspan=6, sticky='nswe')
#        self.catalog_id_button.grid(row=0, column=7)
#        self.run_button.grid(row=1, column=7)
#        self.screen_frame_top.grid(row=0, column=0, sticky='nswe')
#        self.screen_frame_top.grid_columnconfigure(3, weight=1)
        
        self.page_frame = tkinter.Frame(self.screen_frame)
        self.first_page_button = tkinter.Button(self.page_frame, text='<<', font=font.Font(size=8, weight='bold'), command=partial(self.change_record_page_navigation, first=True))
        self.previous_page_button = tkinter.Button(self.page_frame, text='<', font=font.Font(size=8, weight='bold'), command=partial(self.change_record_page_navigation, change=-1))
        self.entry_page = tkinter.Entry(self.page_frame, textvariable=self.entry_page_text, font=font.Font(size=self.size))
        self.page_divider = tkinter.Label(self.page_frame, text=' / ', font=font.Font(size=self.size))
        self.page_label = tkinter.Label(self.page_frame, textvariable=self.number_of_pages, font=font.Font(size=self.size))
        self.next_page_button = tkinter.Button(self.page_frame, text='>', font=font.Font(size=8, weight='bold'), command=partial(self.change_record_page_navigation, change=1))
        self.last_page_button = tkinter.Button(self.page_frame, text='>>', font=font.Font(size=8, weight='bold'), command=partial(self.change_record_page_navigation, last=True))
#        self.screen_label.grid(row=2, column=0, columnspan=8, rowspan=6, sticky='nswe', padx=6, pady=6)
#        self.page_frame.grid(row=8, column=4)
#        self.page_frame.grid(row=8, column=4)
        self.page_frame.grid(row=2, column=0)
        
        self.first_page_button.grid(row=0, column=1)
        self.previous_page_button.grid(row=0, column=2)
        self.entry_page.grid(row=0, column=3)
        self.page_divider.grid(row=0, column=4)
        self.page_label.grid(row=0, column=5)
        self.next_page_button.grid(row=0, column=6)
        self.last_page_button.grid(row=0, column=7)
#        self.grid_rowconfigure(2, weight=1)
#        self.grid_columnconfigure(3, weight=1)
#        self.grid_columnconfigure(6, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(3, weight=1)
        self.grid_columnconfigure(6, weight=1)
        
        self.entry_page.bind('<Return>', self.change_record_page_event_handler)
#        self.entry_page.bind('<<change_page>>', self.change_record_page_event_handler)
#        self.event_generate('<Configure>', when='tail')
    def create_spreadsheet(self):
#        self.spreadsheet_frame = tkinter.Frame(self)
        self.spreadsheet_frame_main = tkinter.Frame(self.spreadsheet_frame)
        self.spreadsheet_frame.grid_rowconfigure(0, weight=1)
        
        self.spreadsheet_frame.grid_columnconfigure(0, weight=1)
#                
        self.spreadsheet_frame_main.bind('<Configure>', self.set_up_canvas)
        
        #Sets up the screen canvas to enable scrolling.
        self.spreadsheet_screen_canvas = tkinter.Canvas(self.spreadsheet_frame_main, highlightthickness=0)
        self.spreadsheet_frame_inner = tkinter.Frame(self.spreadsheet_screen_canvas)
        
        #Sets up the vertical and horizontal scrollbars.
        self.spreadsheet_scrollbar_vertical=tkinter.Scrollbar(self.spreadsheet_frame_main, orient='vertical', command=self.spreadsheet_screen_canvas.yview)
        self.spreadsheet_scrollbar_horizontal=tkinter.Scrollbar(self.spreadsheet_frame_main, orient='horizontal', command=self.spreadsheet_screen_canvas.xview)
        self.spreadsheet_scrollbar_vertical.grid(row=0, column=8, sticky='ns')
        self.spreadsheet_scrollbar_horizontal.grid(row=1, column=0, columnspan=8, sticky='we')
#        self.spreadsheet_frame_inner.grid_rowconfigure(0, weight=1, minsize=461)
#        self.spreadsheet_frame_inner.grid_columnconfigure(0, weight=1, minsize=620)
        #Adds space to the frame to prevent the vertical scrollbar from cutting off text.
#        self.spreadsheet_frame_inner.grid_columnconfigure(9, minsize=15)
#        self.spreadsheet_frame_inner.grid_rowconfigure(8, weight=1)
#        self.spreadsheet_frame_inner.grid_columnconfigure(6, weight=1)
        
        #Finishes setting up the main canvas by connecting the scrollbar and adding the frame containing hot folders information.
        self.spreadsheet_screen_canvas.configure(yscrollcommand=self.spreadsheet_scrollbar_vertical.set, xscrollcommand=self.spreadsheet_scrollbar_horizontal.set)
        
        self.canvas_frame = self.spreadsheet_screen_canvas.create_window((0,0), window=self.spreadsheet_frame_inner, anchor='nw', tags='self.spreadsheet_frame_inner')
        
        #Binds the canvases to the configure event.
#        self.spreadsheet_frame.bind('<Configure>', self.set_up_spreadsheet_canvas)
#        self.spreadsheet_frame.bind('<<change_page>>', self.set_up_canvas)
        
#        self.spreadsheet_frame.grid(row=2, column=0, columnspan=8, sticky='nswe', padx=6)
        
#        self.spreadsheet_frame.grid_rowconfigure(0, weight=1)
#        self.spreadsheet_frame.grid_columnconfigure(0, weight=1)
        self.spreadsheet_frame_main.grid_rowconfigure(0, weight=1)
        self.spreadsheet_frame_main.grid_columnconfigure(0, weight=1)
        
        self.spreadsheet_screen_canvas.grid(row=0, column=0, columnspan=8, sticky='nswe')
        
        self.spreadsheet_catalog_id_label = tkinter.Label(self.spreadsheet_frame_inner, textvariable=self.catalog_id_text, font=font.Font(size=self.size))
        self.spreadsheet_uuid_label = tkinter.Label(self.spreadsheet_frame_inner, text='DDSnext UUID:\t', font=font.Font(size=self.size))
#        self.spreadsheet_collection_label = tkinter.Label(self.spreadsheet_frame_inner, text='Collection:\t', font=font.Font(size=self.size))
        #Sets up button to add collection info to spreadsheet.
        self.spreadsheet_collection_button = tkinter.Menubutton(self.spreadsheet_frame_inner, text='Collection:\t', font=font.Font(size=self.size), relief=tkinter.RAISED, anchor='w')
        self.spreadsheet_collection_button.grid()
        self.spreadsheet_collection_button.menu = tkinter.Menu(self.spreadsheet_collection_button, tearoff=0)
        self.spreadsheet_collection_button['menu'] =  self.spreadsheet_collection_button.menu
        self.spreadsheet_collection_button.menu.add_command(label='Monograph', font=font.Font(size=self.size), command=partial(self.change_collection, 'monograph'))
        self.spreadsheet_collection_button.menu.add_command(label='Newspaper', font=font.Font(size=self.size), command=partial(self.change_collection, 'newspaper'))
        self.spreadsheet_collection_button.menu.add_command(label='Serial', font=font.Font(size=self.size), command=partial(self.change_collection, 'serial'))



        self.spreadsheet_catalog_id_label.grid(row=0, column=0, sticky='w')
        self.spreadsheet_uuid_label.grid(row=0, column=4, sticky='w')
#        self.spreadsheet_collection_label.grid(row=0, column=8, sticky='w')
        self.spreadsheet_collection_button.grid(row=0, column=8, sticky='nswe')
        
        
        self.spreadsheet_bottom_frame = tkinter.Frame(self.spreadsheet_frame)
        self.spreadsheet_run_button = tkinter.Button(self.spreadsheet_bottom_frame, text='Run', font=font.Font(size=self.size), command=self.run)
#        self.spreadsheet_run_button = tkinter.Button(self.spreadsheet_bottom_frame, text='Run', font=font.Font(size=self.size), command=partial(self.run, 'batch'))
        self.spreadsheet_clear_button = tkinter.Button(self.spreadsheet_bottom_frame, text='Clear', font=font.Font(size=self.size), state = 'disabled', command=self.clear)
        
        self.spreadsheet_clear_options_button = tkinter.Menubutton(self.spreadsheet_bottom_frame, text='^', font=font.Font(size=5, weight='bold'), relief=tkinter.RAISED)
        self.spreadsheet_clear_options_button.menu = tkinter.Menu(self.spreadsheet_clear_options_button, tearoff=0)
        
        self.clear_value = tkinter.StringVar()
        self.clear_value.set('all')
        self.spreadsheet_clear_options_button['menu'] =  self.spreadsheet_clear_options_button.menu
        self.spreadsheet_clear_options_button.menu.add_radiobutton(label='Entries and Records', variable=self.clear_value, value='all', font=font.Font(size=self.size), command = self.get_clear_state)
        self.spreadsheet_clear_options_button.menu.add_radiobutton(label='Entries', variable=self.clear_value, value='entries', font=font.Font(size=self.size), command = self.get_clear_state)
        self.spreadsheet_clear_options_button.menu.add_radiobutton(label='Records', variable=self.clear_value, value='records', font=font.Font(size=self.size), command = self.get_clear_state)
        
        #Sets up Button to switch between Millennium and Worldcat.
        self.spreadsheet_catalog_id_button = tkinter.Menubutton(self.spreadsheet_bottom_frame, textvariable=self.catalog_id_button_text, font=font.Font(size=self.size), relief=tkinter.RAISED)
        self.spreadsheet_catalog_id_button.grid()
        self.spreadsheet_catalog_id_button.menu = tkinter.Menu(self.spreadsheet_catalog_id_button, tearoff=0)
        self.spreadsheet_catalog_id_button['menu'] =  self.spreadsheet_catalog_id_button.menu
#        self.spreadsheet_catalog_id_button.menu.add_command(label='Millennium', font=font.Font(size=self.size), command=self.change_to_millennium)
#        self.spreadsheet_catalog_id_button.menu.add_command(label='Worldcat', font=font.Font(size=self.size), command=self.change_to_worldcat)
        self.spreadsheet_catalog_id_button.menu.add_command(label='Millennium', font=font.Font(size=self.size), command=partial(self.change_record_source, 'millennium'))
        self.spreadsheet_catalog_id_button.menu.add_command(label='Worldcat', font=font.Font(size=self.size), command=partial(self.change_record_source, 'worldcat'))
        
##        self.spreadsheet_catalog_id_button_label = tkinter.Label(self.spreadsheet_bottom_frame, text='Record source', font=font.Font(size=self.size), bg='yellow')
#        self.spreadsheet_catalog_id_button_label = tkinter.Label(self.spreadsheet_frame_main, text='Record source', font=font.Font(size=self.size), bg='yellow')
#        self.spreadsheet_catalog_id_button.bind('<Enter>', partial(self.manage_catalog_id_popup, label=self.spreadsheet_catalog_id_button_label, type='enter'))
#        self.spreadsheet_catalog_id_button.bind('<Leave>', partial(self.manage_catalog_id_popup, label=self.spreadsheet_catalog_id_button_label, type='leave'))
        
        self.spreadsheet_frame_main.grid(row=0, column=0, sticky='nswe')
        self.spreadsheet_bottom_frame.grid(row=1, column=0)
#        self.spreadsheet_bottom_frame.rowconfigure(0, minsize=26)
#        , sticky='nswe'
        self.spreadsheet_run_button.grid(row=1, column=1)
        self.spreadsheet_clear_button.grid(row=1, column=2)
        self.spreadsheet_clear_options_button.grid(row=1, column=3, sticky='nw')
        self.spreadsheet_catalog_id_button.grid(row=1, column=0, sticky='nswe')
        
##        self.spreadsheet_catalog_id_button_label.grid(row=0, column=0, columnspan=2, sticky='nsw')
#        self.spreadsheet_catalog_id_button_label.grid(row=1, column=1, columnspan=2, sticky='nsw')
#        self.spreadsheet_catalog_id_button_label.grid_remove()
        
        self.set_up_spreadsheet()
        self.clipboard_content = ''
        
    
    #Changes record source.
    def change_record_source(self, record_source):
        self.catalog_id_button_text.set(record_source_dict[record_source]['button'])
        self.catalog_id_text.set(record_source_dict[record_source]['label'])
        self.record_source = record_source
    #Changes collection and fills entries in input spreadsheet.
    def change_collection(self, collection):
        for row in self.spreadsheet:
            self.spreadsheet[row]['collection'][0].set(collection)

#    #Changes to Millennium settings.
#    def change_to_millennium(self):
#        self.catalog_id_button_text.set('Millennium')
#        self.catalog_id_text.set('Bib number:\t')
#        self.record_source = 'millennium'
#    #Changes to Worldcat settings.
#    def change_to_worldcat(self):
#        self.catalog_id_button_text.set('Worldcat')
#        self.catalog_id_text.set('OCLC number:\t')
#        self.record_source = 'worldcat'
    #Sets settings
    def set_settings(self):
        settings = setting_dialog(self)
#    #
#    def run(self, import_type):
#        self.screen_text.set('')
#        self.input_values = []
#        if import_type == 'single':
#            self.input_values.append([self.entry_catalog_id_text.get(), self.entry_uuid_text.get()])
#        elif import_type == 'batch':
#            for row in self.spreadsheet:
##                if self.spreadsheet[row]['cat_id'][0].get() is not None and self.spreadsheet[row]['cat_id'][0].get() != '' and self.spreadsheet[row]['ddsnext_uuid'][0].get() is not None and self.spreadsheet[row]['ddsnext_uuid'][0].get() != '' and self.spreadsheet[row]['collection'][0].get() is not None and self.spreadsheet[row]['collection'][0].get() != '':
#                if self.spreadsheet[row]['cat_id'][0].get() is not None and self.spreadsheet[row]['cat_id'][0].get() != '' and self.spreadsheet[row]['ddsnext_uuid'][0].get() is not None and self.spreadsheet[row]['ddsnext_uuid'][0].get() != '':
#                    self.input_values.append([self.spreadsheet[row]['cat_id'][0].get(), self.spreadsheet[row]['ddsnext_uuid'][0].get(), self.spreadsheet[row]['collection'][0].get()])
##        self.records = process_records(self.record_source, self.file_types, self.input_values)
#        self.process_records()
#        self.pages = {}
#        page_num = 0
#        if self.records != {}:
#            for key in list(self.records):
#                page_num += 1
#                self.pages[page_num] = key
#            self.screen_text.set(self.records[self.pages[1]])
#            self.entry_page_text.set(1)
#            self.number_of_pages.set(page_num)
#            self.current_record = [1, self.pages[1]]
#            self.windows.select(1)
#            self.spreadsheet_clear_button['state'] = 'normal'
#            self.file_menu.entryconfig('Export', state= 'normal')
#            self.update()
#            self.screen_frame_outer.event_generate('<<change_page>>', when='tail')
    
    def remove_returns(self, text):
        while re.match('(.+)(?:\r)(.*$)', text):
            text = re.match('(.+)(?:\r)(.*$)', text).group(1) + re.match('(.+)(?:\r)(.*$)', text).group(2)
        return text
    
    #
    def run(self):
        self.screen_text.set('')
        self.input_values = []
        for row in self.spreadsheet:
            if self.spreadsheet[row]['cat_id'][0].get() is not None and self.spreadsheet[row]['cat_id'][0].get() != '' and self.spreadsheet[row]['ddsnext_uuid'][0].get() is not None and self.spreadsheet[row]['ddsnext_uuid'][0].get() != '':
                self.input_values.append([self.remove_returns(self.spreadsheet[row]['cat_id'][0].get()), self.remove_returns(self.spreadsheet[row]['ddsnext_uuid'][0].get()), self.remove_returns(self.spreadsheet[row]['collection'][0].get())])
        self.process_records()
        self.pages = {}
        page_num = 0
        if self.records != {}:
            for key in list(self.records):
                page_num += 1
                self.pages[page_num] = key
            self.screen_text.set(self.records[self.pages[1]])
            self.entry_page_text.set(1)
            self.number_of_pages.set(page_num)
            self.current_record = [1, self.pages[1]]
            self.windows.select(1)
            self.spreadsheet_clear_button['state'] = 'normal'
            self.file_menu.entryconfig('Export', state= 'normal')
            self.update()
            self.screen_frame_outer.event_generate('<<change_page>>', when='tail')
    
    def change_record_page_navigation(self, change=0, last=False, first=False):
        if first:
            self.entry_page_text.set(1)
            self.change_record_page()
        elif last:
           self.entry_page_text.set(self.number_of_pages.get())
           self.change_record_page()
        else:
            page_num = int(self.entry_page_text.get()) + change
            if page_num > 0 and page_num <= int(self.number_of_pages.get()):
                self.entry_page_text.set(page_num)
                self.change_record_page()
    
    #Change record by page.
    def change_record_page(self):
        self.screen_text.set('')
        page_num = int(self.entry_page.get())
        if self.pages != {} and self.records != {}:
            if page_num in self.pages:
                self.screen_text.set(self.records[self.pages[page_num]])
                self.entry_page_text.set(page_num)
            else:
                self.entry_page_text.set(self.current_record[0])
            self.update()
            self.screen_frame_outer.event_generate('<<change_page>>', when='tail')
    #Returns true if all spreadsheet entries are empty.  Returns false otherwise.
    def check_entries_empty(self):
        verify_empty = True
        for row in self.spreadsheet:
            if self.spreadsheet[row]['cat_id'][0].get() != '' or self.spreadsheet[row]['ddsnext_uuid'][0].get() != '' or self.spreadsheet[row]['collection'][0].get() != '':
                verify_empty = False
                break
        return verify_empty
    #
    def retry_get_clear_state(self, event):
        while self.attempt < 10:
            self.update()
            self.attempt += 1
            if not self.check_entries_empty():
                self.get_clear_state()
                break
        self.attempt = 1
        event.widget.unbind('<<retry_get_clear_state>>')
        self.get_clear_state()
    
#    def manage_catalog_id_popup(self, event, label, type):
#        if type == 'enter':
#            label.grid()
#        if type == 'leave':
#            label.grid_remove()
    
    #Event handler that determines if the clear button should be disabled or enabled.
    def check_entries_empty_handler(self, event):
        self.attempt = 1
        event.widget.bind('<<retry_get_clear_state>>', self.retry_get_clear_state)
        event.widget.event_generate('<<retry_get_clear_state>>', when='tail')
    #Determines if the clear button should be disabled or enabled.
    def get_clear_state(self):
        verify_empty = True
        #Checks if the spreadsheet entries are empty if clear_value is 'all' or 'entries'.
        if self.clear_value.get() == 'all' or self.clear_value.get() == 'entries':
            verify_empty = self.check_entries_empty()
        #Checks if there are any marc records if clear_value is 'all' or 'records'.
        if self.clear_value.get() == 'all' or self.clear_value.get() == 'records':
            if self.records != {}:
                verify_empty = False
        #If no valid check returned any values, disables the check button.
        if verify_empty:
            self.spreadsheet_clear_button['state'] = 'disabled'
        #If any valid check returned any values, enables the check button.
        else:
            self.spreadsheet_clear_button['state'] = 'normal'
    
    #Event handler that determines if the clear button should be disabled or enabled.
    def get_clear_state_handler(self, event):
        self.get_clear_state()
    
    #Clears screen, records, and spreadsheet.
    def clear(self):
        #Reset the screen text, record, and page info if clear_value is 'all' or 'records'.
        if self.clear_value.get() == 'all' or self.clear_value.get() == 'records':
            self.screen_text.set('')
            self.records = {}
            self.pages = {}
            self.entry_page_text.set(0)
            self.number_of_pages.set(0)
            self.screen_scrollbar_vertical.grid_forget()
            self.screen_scrollbar_horizontal.grid_forget()
            self.scrollbar_visible = False
        #Reset the spreadsheet entries if clear_value is 'all' or 'entries'.
        if self.clear_value.get() == 'all' or self.clear_value.get() == 'entries':
            for row in self.spreadsheet:
                self.spreadsheet[row]['cat_id'][0].set('')
                self.spreadsheet[row]['ddsnext_uuid'][0].set('')
                self.spreadsheet[row]['collection'][0].set('')
        #Disables the check button.
        self.spreadsheet_clear_button['state'] = 'disabled'
        self.file_menu.entryconfig('Export', state= 'disabled')
    
    #Clears template spreadsheet.
    def clear_template(self, template_name):
#        self.template_dict[template_name]
        for row in self.template_spreadsheet[template_name]:
            for item in self.template_spreadsheet[template_name][row]:
#                print(self.template_spreadsheet[template_name][row][item])
                if type(self.template_spreadsheet[template_name][row][item]) is tkinter.Entry:
                    self.set_entry(self.template_spreadsheet[template_name][row][item])
        #Disables the clear button.
#        self.spreadsheet_clear_button['state'] = 'disabled'
    
    def set_entry(self, given_entry, text=''):
        disable = False
        if given_entry['state'] == 'disabled':
            given_entry['state'] = 'normal'
            disable = True
        given_entry.delete(0, tkinter.END)
        given_entry.insert(0, text)
        if disable:
            given_entry['state'] = 'disabled'
    
    #Sets up the screen and spreadsheet canvas for scrolling.
    def set_up_canvas(self, event):
        width = event.width - 1
        self.screen_canvas.configure(scrollregion=self.screen_canvas.bbox('all'))
#        self.spreadsheet_screen_canvas.itemconfigure('self.spreadsheet_frame_inner', width=width)
        self.spreadsheet_screen_canvas.configure(scrollregion=self.spreadsheet_screen_canvas.bbox('all'))
    
    #Sets up the template canvases for scrolling.
    def set_up_template_canvas(self, event):
#        width = event.width - 1
#        print(self.winfo_width())
#        print(width)
        for template_name in self.template_dict:
#            print(template_name, width, self.winfo_width() - 15, self.winfo_width(), self.template_dict[template_name]['template_spreadsheet_frame'].winfo_width())
#            if self.template_dict[template_name]['defaut'] is None:
#                self.template_dict[template_name]['defaut'] = self.template_dict[template_name]['template_spreadsheet_frame'].winfo_width()
#            print(template_name, width, self.template_dict[template_name]['defaut'], self.template_dict[template_name]['defaut'] + width)
#            self.template_dict[template_name]['template_screen_canvas'].itemconfigure(template_name, width=self.template_dict[template_name]['defaut'] + width)
#            self.template_dict[template_name]['template_screen_canvas'].itemconfigure(template_name, width=self.template_dict[template_name]['template_spreadsheet_frame'].winfo_width() + (width - 619))
            self.template_dict[template_name]['template_screen_canvas'].configure(scrollregion=self.template_dict[template_name]['template_screen_canvas'].bbox('all'))
    
    #Sets up the screen canvas for scrolling.
    def set_up_records(self, event):
        self.screen_canvas.configure(scrollregion=self.screen_canvas.bbox('all'))
        if not self.scrollbar_visible:
            self.screen_scrollbar_vertical.grid(row=0, column=2, rowspan=2, sticky='ns')
            self.screen_scrollbar_horizontal.grid(row=1, column=0, sticky='we')
            self.scrollbar_visible = True
        self.screen_canvas.yview_moveto('0')
        self.screen_canvas.xview_moveto('0')
        
    #Sets up the spreadsheet.
    def set_up_spreadsheet(self):
        for row in range(1, 1000):
            self.spreadsheet[row] = {'cat_id' : [tkinter.StringVar(), None], 'ddsnext_uuid' : [tkinter.StringVar(), None], 'collection' : [tkinter.StringVar(), None]}
            self.spreadsheet[row]['cat_id'][1] = tkinter.Entry(self.spreadsheet_frame_inner, textvariable=self.spreadsheet[row]['cat_id'][0], font=font.Font(size=self.size))
            self.spreadsheet[row]['ddsnext_uuid'][1] = tkinter.Entry(self.spreadsheet_frame_inner, textvariable=self.spreadsheet[row]['ddsnext_uuid'][0], font=font.Font(size=self.size))
            self.spreadsheet[row]['collection'][1] = tkinter.Entry(self.spreadsheet_frame_inner, textvariable=self.spreadsheet[row]['collection'][0], font=font.Font(size=self.size))
#            self.spreadsheet[row]['collection'][1] = tkinter.Label(self.spreadsheet_frame_inner, textvariable=self.spreadsheet[row]['collection'][0], font=font.Font(size=self.size))
            self.spreadsheet[row]['cat_id'][1].grid(row=row, column=0, columnspan=4, sticky='nswe')
            self.spreadsheet[row]['ddsnext_uuid'][1].grid(row=row, column=4, columnspan=4, sticky='nswe')
            self.spreadsheet[row]['collection'][1].grid(row=row, column=8, columnspan=4, sticky='nswe')
            self.spreadsheet_frame_inner.grid_rowconfigure(row, weight=1, minsize=25)
            
            self.spreadsheet[row]['cat_id'][1].bind('<FocusIn>', self.enter)
            self.spreadsheet[row]['ddsnext_uuid'][1].bind('<FocusIn>', self.enter)
            self.spreadsheet[row]['collection'][1].bind('<FocusIn>', self.enter)
            self.spreadsheet[row]['cat_id'][1].bind('<FocusOut>', self.get_clear_state_handler)
            self.spreadsheet[row]['ddsnext_uuid'][1].bind('<FocusOut>', self.get_clear_state_handler)
            self.spreadsheet[row]['collection'][1].bind('<FocusOut>', self.get_clear_state_handler)
        self.spreadsheet_frame_inner.grid_columnconfigure(0, weight=1, minsize=330)
        self.spreadsheet_frame_inner.grid_columnconfigure(4, weight=1, minsize=330)
        self.spreadsheet_frame_inner.grid_columnconfigure(8, weight=1, minsize=330)
    
    #Sets up the spreadsheet for given template.
    def set_up_template(self, template_name, data):
        #Creates template window and frame if not the window dictionary.
        if not template_name in self.windows_id_dict:
            #Creates the template frame.
            template_frame = tkinter.Frame(self.windows)
            template_frame.grid(row = 0, column = 0, sticky='nswe')
            template_frame.grid_columnconfigure(0, weight = 1)
            template_frame.grid_rowconfigure(0, weight = 1)
            template_frame.bind('<Configure>', self.set_up_template_canvas)
            template_frame.bind('<<add_row>>', self.set_up_template_canvas)
            
            
            #Creates the frame containing the spreadsheet canvas.
            template_spreadsheet_frame_outer = tkinter.Frame(template_frame)
            template_spreadsheet_frame_outer.grid_rowconfigure(0, weight = 1)
            template_spreadsheet_frame_outer.grid_columnconfigure(0, weight = 1)
            template_spreadsheet_frame_outer.grid(row = 0, column = 0, sticky='nswe')
            
            
            #Creates the bottom frame of the template frame.
            template_bottom_frame = tkinter.Frame(template_frame)
            template_bottom_frame.grid(row = 2, column = 0)
            
            #Creates the (outer) canvas for the spreadsheet frame.
            template_screen_canvas = tkinter.Canvas(template_spreadsheet_frame_outer, highlightthickness = 0)
            template_screen_canvas.grid(row = 0, column = 0, sticky='nswe')
            
            #Creates the spreadsheet frame of the template frame.
            template_spreadsheet_frame_shell = tkinter.Frame(template_screen_canvas)
            template_spreadsheet_frame_shell.grid_rowconfigure(1, minsize = 26)
            
            #Creates the spreadsheet frame of the template frame.
            template_spreadsheet_frame = tkinter.Frame(template_spreadsheet_frame_shell)
            template_spreadsheet_frame.grid(row = 0, column = 0, sticky='nswe')
            
            
            
            #Sets up the vertical and horizontal scrollbars.
            spreadsheet_scrollbar_vertical = tkinter.Scrollbar(template_spreadsheet_frame_outer, orient = 'vertical', command = template_screen_canvas.yview)
            spreadsheet_scrollbar_horizontal = tkinter.Scrollbar(template_spreadsheet_frame_outer, orient = 'horizontal', command = template_screen_canvas.xview)
            spreadsheet_scrollbar_vertical.grid(row = 0, column = 2, sticky = 'ns')
            spreadsheet_scrollbar_horizontal.grid(row = 1, column = 0, sticky = 'we')
            
            
            #Finishes setting up the template canvas by connecting the scrollbar.
            template_screen_canvas.configure(yscrollcommand = spreadsheet_scrollbar_vertical.set, xscrollcommand = spreadsheet_scrollbar_horizontal.set)
            template_screen_canvas.create_window((0,0), window = template_spreadsheet_frame_shell, anchor='nw', tags = template_name)
            
            
            
#            self.template_dict[template_name] = {'template_frame' : template_frame, 'template_spreadsheet_frame' : template_spreadsheet_frame, 'template_bottom_frame' : template_bottom_frame, 'template_screen_canvas' : template_screen_canvas, 'defaut' : None}
#            self.template_dict[template_name] = {'template_frame' : template_frame, 'template_spreadsheet_frame' : template_spreadsheet_frame, 'template_bottom_frame' : template_bottom_frame, 'template_screen_canvas' : template_screen_canvas, 'last_column' : None}
            
            #
            self.template_dict[template_name] = {'template_frame' : template_frame, 'template_spreadsheet_frame' : template_spreadsheet_frame, 'template_bottom_frame' : template_bottom_frame, 'template_screen_canvas' : template_screen_canvas}
            
            self.template_spreadsheet[template_name] = {}
            self.template_spreadsheet[template_name]['header'] = {}
            
            self.windows.add(self.template_dict[template_name]['template_frame'], text=template_name)
            self.windows_id_dict[template_name] = len(self.windows_id_dict)
#            print(self.template_dict[template_name])
#            self.template_dict[template_name]['template_spreadsheet_frame'].bind('<Configure>', self.set_up_canvas)
            
        self.template_dict[template_name]['template_spreadsheet_frame'].grid_columnconfigure(0, weight=1)
        
        row = 0
        item_row = 0
        #Creates header row in spreedsheet.
        column=0
#        if template_name == 'kbart_metadata':
#            self.template_spreadsheet[template_name]['header'][column] = tkinter.Label(self.template_dict[template_name]['template_spreadsheet_frame'])
#            self.template_spreadsheet[template_name]['header'][column].grid(row=0, column=column, sticky='nswe')
#            column += 1
        for data_item in data['header']:
#            print(data_item)
#            if type(data['header'][data_item]) is dict:
#                print(data_item)
##                column = 0
##                if template_name == 'kbart_metadata':
##                    self.template_spreadsheet[template_name]['header'][column] = tkinter.Label(self.template_dict[template_name]['template_spreadsheet_frame'])
##                    column += 1
#                for data_item_item in data['header'][data_item]:
#                    self.template_spreadsheet[template_name]['header'][data_item_item] = tkinter.Entry(self.template_dict[template_name]['template_spreadsheet_frame'])
##                    self.template_spreadsheet[template_name]['header'][data_item_item].delete(0, tkinter.END)
#                    to_print = ''
#                    if data['header'][data_item][data_item_item] is not None:
#                        to_print = data['header'][data_item][data_item_item]
##                    self.template_spreadsheet[template_name]['header'][data_item_item].insert(0, to_print)
#                    self.set_entry(self.template_spreadsheet[template_name]['header'][data_item_item], text = to_print)
#                    self.template_spreadsheet[template_name]['header'][data_item_item].grid(row=0, column=column, sticky='nswe')
#                    self.template_dict[template_name]['template_spreadsheet_frame'].grid_columnconfigure(column, weight=1, minsize=300)
#                    self.template_spreadsheet[template_name]['header'][data_item_item]['state'] = 'disabled'
#                    column += 1
#                item_row += 1
#            
#            else:
#                self.template_spreadsheet[template_name]['header'][data_item] = tkinter.Entry(self.template_dict[template_name]['template_spreadsheet_frame'])
#                self.template_spreadsheet[template_name]['header'][data_item].delete(0, tkinter.END)
#                to_print = ''
#                if data['header'][data_item] is not None:
#                    to_print = data['header'][data_item]
#                self.template_spreadsheet[template_name]['header'][data_item].insert(0, to_print)
#                self.set_entry(self.template_spreadsheet[template_name]['header'][data_item], text = to_print)
#                self.template_spreadsheet[template_name]['header'][data_item].grid(row=0, column=column, sticky='nswe')
#                if template_name == 'descriptive_metadata':
#                    self.template_dict[template_name]['template_spreadsheet_frame'].grid_columnconfigure(column, weight=1, minsize=326)
#                else:
#                    self.template_dict[template_name]['template_spreadsheet_frame'].grid_columnconfigure(column, weight=1, minsize=240)
#                self.template_dict[template_name]['template_spreadsheet_frame'].grid_rowconfigure(row, weight=1, minsize=30)
#
#                self.template_spreadsheet[template_name]['header'][data_item]['state'] = 'disabled'
#                column += 1

            self.template_spreadsheet[template_name]['header'][data_item] = tkinter.Entry(self.template_dict[template_name]['template_spreadsheet_frame'], font=font.Font(size=self.size))
#            self.template_spreadsheet[template_name]['header'][data_item].delete(0, tkinter.END)
            to_print = ''
            if data['header'][data_item] is not None:
                to_print = data['header'][data_item]
#            self.template_spreadsheet[template_name]['header'][data_item].insert(0, to_print)
            self.set_entry(self.template_spreadsheet[template_name]['header'][data_item], text = to_print)
            self.template_spreadsheet[template_name]['header'][data_item].grid(row=0, column=column, sticky='nswe')
            if template_name == 'descriptive_metadata':
                self.template_dict[template_name]['template_spreadsheet_frame'].grid_columnconfigure(column, weight=1, minsize=326)
            else:
                self.template_dict[template_name]['template_spreadsheet_frame'].grid_columnconfigure(column, weight=1, minsize=240)
            self.template_dict[template_name]['template_spreadsheet_frame'].grid_rowconfigure(row, weight=1, minsize=25)

            self.template_spreadsheet[template_name]['header'][data_item]['state'] = 'disabled'
            column += 1
            
#        total_columns = column
#        self.template_dict[template_name]['last_column'] = column
        row += 1
#        row = 1
#        item_row = 1
        for key in data:
            if key != 'header':
                column = 0
#                if template_name == 'kbart_metadata':
#                    if row not in self.template_spreadsheet[template_name]:
#                        self.template_spreadsheet[template_name][row] = {}
#                    self.template_spreadsheet[template_name][row][column] = tkinter.Menubutton(self.template_dict[template_name]['template_spreadsheet_frame'], text = 'Collection')
#                    self.template_spreadsheet[template_name][row][column].grid()
#                    self.template_spreadsheet[template_name][row][column].menu = tkinter.Menu(self.template_spreadsheet[template_name][row][column], tearoff=0)
#                    self.template_spreadsheet[template_name][row][column]['menu'] =  self.template_spreadsheet[template_name][row][column].menu
#                    self.template_spreadsheet[template_name][row][column].menu.add_command(label='monograph', command=partial(self.get_collection_data, template_name, row = row, collection = 'monograph'))
#                    self.template_spreadsheet[template_name][row][column].menu.add_command(label='newspaper', command=partial(self.get_collection_data, template_name, row = row, collection = 'newspaper'))
#                    self.template_spreadsheet[template_name][row][column].menu.add_command(label='serial', command=partial(self.get_collection_data, template_name, row = row, collection = 'serial'))
#                    self.template_spreadsheet[template_name][row][column].grid(row=row, column=column, sticky='we')
#                    column += 1
#                for data_item in data[key]:
#                    if type(data_item) is dict:
#                        column = 0
#                        for data_item_item in data_item:
#                            self.template_spreadsheet[template_name][item_row] = {data_item_item: tkinter.Entry(self.template_dict[template_name]['template_spreadsheet_frame'])}
##                            self.template_spreadsheet[template_name][item_row][data_item_item].delete(0, tkinter.END)
#                            to_print = ''
#                            if data_item[data_item_item] is not None:
#                                to_print = data_item[data_item_item]
##                            self.template_spreadsheet[template_name][item_row][data_item_item].insert(0, to_print)
#                            self.set_entry(self.template_spreadsheet[template_name][item_row][data_item_item], text = to_print)
#                            self.template_spreadsheet[template_name][item_row][data_item_item].grid(row = item_row, column = column, sticky='nswe')
#                            column += 1
#                        item_row += 1
#                    elif type(data[key][data_item]) is dict:
#                        column = 0
#                        for data_item_item in data[key][data_item]:
#                            self.template_spreadsheet[template_name][item_row] = {data_item_item: tkinter.Entry(self.template_dict[template_name]['template_spreadsheet_frame'])}
##                            self.template_spreadsheet[template_name][item_row][data_item_item].delete(0, tkinter.END)
#                            to_print = ''
#                            if data[key][data_item][data_item_item] is not None:
#                                to_print = data[key][data_item][data_item_item]
##                            self.template_spreadsheet[template_name][item_row][data_item_item].insert(0, to_print)
#                            self.set_entry(self.template_spreadsheet[template_name][item_row][data_item_item], text = to_print)
#                            self.template_spreadsheet[template_name][item_row][data_item_item].grid(row = item_row, column = column, sticky='nswe')
#                            column += 1
#                        item_row += 1
#                    else:
#                        if row not in self.template_spreadsheet[template_name]:
#                            self.template_spreadsheet[template_name][row] = {}
#                        self.template_spreadsheet[template_name][row][data_item] = tkinter.Entry(self.template_dict[template_name]['template_spreadsheet_frame'], textvariable=data[key][data_item])
##                        self.template_spreadsheet[template_name][row][data_item].delete(0, tkinter.END)
#                        to_print = ''
#                        if data[key][data_item] is not None:
#                            to_print = data[key][data_item]
##                        self.template_spreadsheet[template_name][row][data_item].insert(0, to_print)
#                        self.set_entry(self.template_spreadsheet[template_name][row][data_item], text = to_print)
#                        self.template_spreadsheet[template_name][row][data_item].grid(row = row, column = column, sticky='nswe')
#                        column += 1
#                self.template_dict[template_name]['template_spreadsheet_frame'].grid_rowconfigure(row, weight=1, minsize=25)
#                row += 1
##                self.template_spreadsheet[template_name]['row'] = row
                for data_item in data[key]:
                    if row not in self.template_spreadsheet[template_name]:
                        self.template_spreadsheet[template_name][row] = {}
                    self.template_spreadsheet[template_name][row][data_item] = tkinter.Entry(self.template_dict[template_name]['template_spreadsheet_frame'], font=font.Font(size=self.size))
#                   self.template_spreadsheet[template_name][row][data_item].delete(0, tkinter.END)
                    to_print = ''
                    if data[key][data_item] is not None:
                        to_print = data[key][data_item]
#                    self.template_spreadsheet[template_name][row][data_item].insert(0, to_print)
                    self.set_entry(self.template_spreadsheet[template_name][row][data_item], text = to_print)
                    self.template_spreadsheet[template_name][row][data_item].grid(row = row, column = column, sticky='nswe')
                    column += 1
                self.template_dict[template_name]['template_spreadsheet_frame'].grid_rowconfigure(row, weight=1, minsize=25)
                row += 1


        last_row = self.get_template_last_row(template_name)
        if last_row < 24:
            for current_row in range(last_row, 24):
                self.add_template_row(template_name, last_row=current_row)
#        for i in range(total_columns):
#            temp = tkinter.Entry(self.template_dict[template_name]['template_spreadsheet_frame'], text=str(i))
#            temp.grid(row = row + item_row, column = i)
        template_run_button = tkinter.Button(self.template_dict[template_name]['template_bottom_frame'], text='Add row', font=font.Font(size=self.size), command=partial(self.add_template_row, template_name))
#        template_clear_button = tkinter.Button(self.template_dict[template_name]['template_bottom_frame'], text='Clear', state = 'disabled', command=partial(self.clear_template, template_name))
        template_clear_button = tkinter.Button(self.template_dict[template_name]['template_bottom_frame'], text='Clear', font=font.Font(size=self.size), command=partial(self.clear_template, template_name))
        
#        spreadsheet_clear_options_button = tkinter.Menubutton(self.template_dict[template_name]['template_bottom_frame'], text='^', font=font.Font(size=5, weight='bold'), relief=tkinter.RAISED)
#        spreadsheet_clear_options_button.menu = tkinter.Menu(spreadsheet_clear_options_button, tearoff=0)
#        
#        clear_value = tkinter.StringVar()
#        clear_value.set('all')
#        spreadsheet_clear_options_button['menu'] =  spreadsheet_clear_options_button.menu
#        spreadsheet_clear_options_button.menu.add_radiobutton(label='Entries and Records', variable=clear_value, value='all', command = self.get_clear_state)
#        spreadsheet_clear_options_button.menu.add_radiobutton(label='Entries', variable=clear_value, value='entries', command = self.get_clear_state)
#        spreadsheet_clear_options_button.menu.add_radiobutton(label='Records', variable=clear_value, value='records', command = self.get_clear_state)
        
        #Sets up Button to switch between Millennium and Worldcat.
#        spreadsheet_catalog_id_button = tkinter.Menubutton(self.template_dict[template_name]['template_bottom_frame'], textvariable=catalog_id_button_text, relief=tkinter.RAISED)
#        spreadsheet_catalog_id_button.grid()
#        spreadsheet_catalog_id_button.menu = tkinter.Menu(spreadsheet_catalog_id_button, tearoff=0)
#        spreadsheet_catalog_id_button['menu'] =  spreadsheet_catalog_id_button.menu
#        spreadsheet_catalog_id_button.menu.add_command(label='Millennium', command=change_to_millennium)
#        spreadsheet_catalog_id_button.menu.add_command(label='Worldcat', command=change_to_worldcat)
        
#        spreadsheet_frame_main.grid(row=0, column=0, sticky='nswe')
#        self.template_dict[template_name]['template_bottom_frame'].grid(row=9, column=0)
#        , sticky='nswe'
        template_run_button.grid(row=0, column=1)
        template_clear_button.grid(row=0, column=2)
#        spreadsheet_clear_options_button.grid(row=0, column=3, sticky='nw')
#        spreadsheet_catalog_id_button.grid(row=0, column=0, sticky='ns')

    #Sets up the spreadsheet for given template.
    def add_to_template(self, template_name, data):
        keys = []
        for key in self.template_spreadsheet[template_name]:
            keys = [key] + keys
        for row in self.template_spreadsheet[template_name]:
            if type(row) is int and row > last_row:
                last_row = row
        return last_row
    def get_template_last_row(self, template_name):
        last_row = 0
        for row in self.template_spreadsheet[template_name]:
            if type(row) is int and row > last_row:
                last_row = row
        return last_row
    def add_template_row(self, template_name, last_row=None):
        if last_row is None:
#            self.template_dict[template_name]
            last_row = 0
#            print(self.template_spreadsheet[template_name])
            for row in self.template_spreadsheet[template_name]:
                if type(row) is int and row > last_row:
                    last_row = row
        last_row += 1
#        row = self.template_spreadsheet[template_name]
#        self.template_spreadsheet[template_name]['column']

#        for column in range(self.template_dict[template_name]['last_column']):
#            if column == 0:
#                self.template_spreadsheet[template_name][last_row] = {}  
#            if column == 0 and template_name == 'kbart_metadata':
#                self.template_spreadsheet[template_name][last_row][column] = tkinter.Menubutton(self.template_dict[template_name]['template_spreadsheet_frame'], text = 'Collection')
#                self.template_spreadsheet[template_name][last_row][column].grid()
#                self.template_spreadsheet[template_name][last_row][column].menu = tkinter.Menu(self.template_spreadsheet[template_name][last_row][column], tearoff=0)
#                self.template_spreadsheet[template_name][last_row][column]['menu'] =  self.template_spreadsheet[template_name][last_row][column].menu
#                self.template_spreadsheet[template_name][last_row][column].menu.add_command(label='monograph', command=partial(self.get_collection_data, template_name, row = last_row, collection = 'monograph'))
#                self.template_spreadsheet[template_name][last_row][column].menu.add_command(label='newspaper', command=partial(self.get_collection_data, template_name, row = last_row, collection = 'newspaper'))
#                self.template_spreadsheet[template_name][last_row][column].menu.add_command(label='serial', command=partial(self.get_collection_data, template_name, row = last_row, collection = 'serial'))
#                self.template_spreadsheet[template_name][last_row][column].grid(row=last_row, column=column, sticky='we')
#            else:
#                self.template_spreadsheet[template_name][last_row][column] = tkinter.Entry(self.template_dict[template_name]['template_spreadsheet_frame'])
#                self.template_spreadsheet[template_name][last_row][column].grid(row = last_row, column = column, sticky='nswe')
        
        column = 0
        for key in self.template_spreadsheet[template_name]['header']:
            if column == 0:
                self.template_spreadsheet[template_name][last_row] = {}  
            self.template_spreadsheet[template_name][last_row][key] = tkinter.Entry(self.template_dict[template_name]['template_spreadsheet_frame'])
            self.template_spreadsheet[template_name][last_row][key].grid(row = last_row, column = column, sticky='nswe')
            column += 1
        
        self.template_dict[template_name]['template_spreadsheet_frame'].grid_rowconfigure(last_row, weight=1, minsize=25)
        
        self.template_dict[template_name]['template_frame'].event_generate('<<add_row>>', when='tail')
    
    def get_collection_data(self, template_name, row, collection):
        if 'coverage_depth' in self.template_spreadsheet[template_name][row]:
            self.set_entry(self.template_spreadsheet[template_name][row]['coverage_depth'], text=collection_dict[collection]['coverage_depth'])
        else:
            self.set_entry(self.template_spreadsheet[template_name][row][14], text=collection_dict[collection]['coverage_depth'])
        if 'oclc_collection_name' in self.template_spreadsheet[template_name][row]:
            self.set_entry(self.template_spreadsheet[template_name][row]['oclc_collection_name'], text=collection_dict[collection]['oclc_collection_name'])
        else:
            self.set_entry(self.template_spreadsheet[template_name][row][21], text=collection_dict[collection]['oclc_collection_name'])
        if 'oclc_collection_id' in self.template_spreadsheet[template_name][row]:
            self.set_entry(self.template_spreadsheet[template_name][row]['oclc_collection_id'], text=collection_dict[collection]['oclc_collection_id'])
        else:
            self.set_entry(self.template_spreadsheet[template_name][row][22], text=collection_dict[collection]['oclc_collection_id'])
        
    #Event handler that changes page if the page entry is the focus.
    def change_record_page_event_handler(self, event):
        if self.entry_page.focus_get():
            self.entry_page_text.set(self.entry_page.get())
            self.change_record_page()
    #
    def enter(self, event):
        event.widget.bind('<Control-KeyPress-v>', self.check_clipboard)
        event.widget.bind('<KeyPress>', self.check_entries_empty_handler)
    #
    def restore_clipboard(self, event):
#        self.clipboard_append(self.clipboard_content)
#        event.widget.unbind('<<restore_clipboard>>')
        
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(self.clipboard_content)
        win32clipboard.CloseClipboard()
        event.widget.unbind('<<restore_clipboard>>')
    #
    def check_clipboard(self, event):
        win32clipboard.OpenClipboard()
        self.clipboard_content = win32clipboard.GetClipboardData()
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()
        
#        self.clipboard_content = self.clipboard_get()
#        self.clipboard_clear()
        current_widget = event.widget
        for match in re.finditer('([^\n]+)(|\n|$)', self.clipboard_content, re.MULTILINE):
            match_text = match.group(1)
            tab_match = re.match('([^\t]+)(?:\t)(.+$)', match.group(1), re.MULTILINE)
            number_of_tabs = 0
            if tab_match:
                while tab_match:
                    current_widget.delete(0, tkinter.END)
                    current_widget.insert(0, tab_match.group(1))
                    current_widget = current_widget.tk_focusNext()
                    match_text = tab_match.group(2)
                    tab_match = re.match('([^\t]+)(?:\t)(.+$)', match_text, re.MULTILINE)
                    number_of_tabs += 1
                else:
                    current_widget.delete(0, tkinter.END)
                    current_widget.insert(0, match_text)
                    if re.match('(\n|$)' , match.group(2), re.MULTILINE):
                        if number_of_tabs == 1:
                            current_widget = current_widget.tk_focusNext().tk_focusNext()
                        elif number_of_tabs == 2:
                            current_widget = current_widget.tk_focusNext()
            else:
                current_widget.delete(0, tkinter.END)
                current_widget.insert(0, match_text)
                if re.match('(\n|$)' , match.group(2)):
                    current_widget = current_widget.tk_focusNext().tk_focusNext().tk_focusNext()
        self.get_clear_state()
        current_widget.bind('<<restore_clipboard>>', self.restore_clipboard)
        current_widget.event_generate('<<restore_clipboard>>', when='tail')
        current_widget.focus()
    
    #Three columns
#    def check_clipboard(self, event):
#        win32clipboard.OpenClipboard()
#        self.clipboard_content = win32clipboard.GetClipboardData()
#        win32clipboard.EmptyClipboard()
#        win32clipboard.CloseClipboard()
#        
##        self.clipboard_content = self.clipboard_get()
##        self.clipboard_clear()
#        current_widget = event.widget
#        for match in re.finditer('([^\n]+)(|\n|$)', self.clipboard_content, re.MULTILINE):
#            match_text = match.group(1)
#            tab_match = re.match('([^\t]+)(?:\t)(.+$)', match.group(1), re.MULTILINE)
#            number_of_tabs = 0
#            if tab_match:
#                while tab_match:
#                    current_widget.delete(0, tkinter.END)
#                    current_widget.insert(0, tab_match.group(1))
#                    current_widget = current_widget.tk_focusNext()
#                    match_text = tab_match.group(2)
#                    tab_match = re.match('([^\t]+)(?:\t)(.+$)', match_text, re.MULTILINE)
#                    number_of_tabs += 1
#                else:
#                    current_widget.delete(0, tkinter.END)
#                    current_widget.insert(0, match_text)
#                    if re.match('(\n|$)' , match.group(2), re.MULTILINE):
#                        if number_of_tabs == 1:
#                            current_widget = current_widget.tk_focusNext().tk_focusNext()
#                        elif number_of_tabs == 2:
#                            current_widget = current_widget.tk_focusNext()
#            else:
#                current_widget.delete(0, tkinter.END)
#                current_widget.insert(0, match_text)
#                if re.match('(\n|$)' , match.group(2)):
#                    current_widget = current_widget.tk_focusNext().tk_focusNext().tk_focusNext()
#        self.get_clear_state()
#        current_widget.bind('<<restore_clipboard>>', self.restore_clipboard)
#        current_widget.event_generate('<<restore_clipboard>>', when='tail')
#        current_widget.focus()
    
    def process_records(self):
        self.records = {}
        self.title_metadata = {}
        self.descriptive_metadata = {}
        self.kbart_metadata = {}
        self.title_metadata['header'] = {'title_name' : 'Title Name', 'title_uuid' : 'Title UUID', 'title_material_type' : 'Title Material Type', 'title_format' : 'Title Format', 'title_oclc' : 'Title OCLC', 'title_digital_holding_range' : 'Title Digital Holding Range', 'title_resolution' : 'Title Resolution', 'title_color_depth' : 'Title Color Depth', 'title_location_code' : 'Title Location Code', 'title_catalog_link' : 'Title Catalog Link', 'title_external_link' : 'Title External Link'}
#        self.descriptive_metadata['header'] = {'header' : {'title_uuid' : 'title uuid', 'field' : 'field', 'value' : 'value'}}
        self.descriptive_metadata['header'] = {'title_uuid' : 'title uuid', 'field' : 'field', 'value' : 'value'}
        self.kbart_metadata['header'] = {'publication_title' : 'publication_title', 'print_identifier' : 'print_identifier', 'online_identifier' : 'online_identifier', 'date_first_issue_online' : 'date_first_issue_online', 'num_first_vol_online' : 'num_first_vol_online', 'num_first_issue_online' : 'num_first_issue_online', 'date_last_issue_online' : 'date_last_issue_online', 'num_last_vol_online' : 'num_last_vol_online', 'num_last_issue_online' : 'num_last_issue_online', 'title_url' : 'title_url', 'first_author' : 'first_author', 'title_id' : 'title_id', 'embargo_info' : 'embargo_info', 'coverage_depth' : 'coverage_depth', 'coverage_notes' : 'coverage_notes', 'publisher_name' : 'publisher_name', 'location' : 'location', 'title_notes' : 'title_notes', 'staff_notes' : 'staff_notes', 'vendor_id' : 'vendor_id', 'oclc_collection_name' : 'oclc_collection_name', 'oclc_collection_id' : 'oclc_collection_id', 'oclc_entry_id' : 'oclc_entry_id', 'oclc_linkscheme' : 'oclc_linkscheme', 'oclc_number' : 'oclc_number', 'action' : 'ACTION'}
        title_items_row = 0
        descriptive_metadata_row = 0
        kbart_metadata_row = 0
        for input_value in self.input_values:
            title_items = None
            descriptive_items = None
            kbart_items = None
            if self.record_source == 'worldcat':
                if input_value[2] == '':
                    marc_record, title_items, descriptive_items, kbart_items = process_oclc(input_value[0], input_value[1])
                else:
                    marc_record, title_items, descriptive_items, kbart_items = process_oclc(input_value[0], input_value[1], input_value[2])
            else:
                if input_value[2] == '':
                    marc_record, title_items, descriptive_items, kbart_items = process_bib_num(input_value[0], input_value[1])
                else:
                    marc_record, title_items, descriptive_items, kbart_items = process_bib_num(input_value[0], input_value[1], input_value[2])
            self.records[input_value[0]] = marc_record
#            self.title_metadata[input_value[0]] = title_items
            self.title_metadata[title_items_row] = title_items
            title_items_row += 1
#            self.descriptive_metadata[input_value[0]] = descriptive_items
            for key in descriptive_items:
                self.descriptive_metadata[descriptive_metadata_row] = descriptive_items[key]
                descriptive_metadata_row += 1
#            self.kbart_metadata[input_value[0]] = kbart_items
            self.kbart_metadata[kbart_metadata_row] = kbart_items
            kbart_metadata_row += 1
#            print(self.kbart_metadata)

#        if not 'title_items' in self.windows_id_dict:
##            self.windows.add(self.screen_frame, text='title_items')
##            self.windows_id_dict['title_items'] = len(self.windows_id_dict)
#            self.set_up_template('title_metadata', self.title_metadata)
#        else:
#            self.add_to_template('title_metadata', self.title_metadata)
#        if not 'descriptive_metadata' in self.windows_id_dict:
##            self.windows.add(self.screen_frame, text='descriptive_metadata')
##            self.windows_id_dict['descriptive_metadata'] = len(self.windows_id_dict)
##            print(convert_to_dict(self.descriptive_metadata))
#            self.set_up_template('descriptive_metadata', self.descriptive_metadata)
#        if not 'kbart_metadata' in self.windows_id_dict:
##            self.windows.add(self.screen_frame, text='kbart_metadata')
##            self.windows_id_dict['kbart_metadata'] = len(self.windows_id_dict)
#            self.set_up_template('kbart_metadata', self.kbart_metadata)
##        print(self.windows_id_dict)
##        for 
##    
        self.set_up_template('title_metadata', self.title_metadata)
        self.set_up_template('descriptive_metadata', self.descriptive_metadata)
        self.set_up_template('kbart_metadata', self.kbart_metadata)
    def export(self):
        export_dialog(self)
    #Terminates process and closes window
    def teminate(self):
        root.destroy()

#Removes end punctuation and spaces.
def fix_end_char(text):
    output = text
    if re.search('(.*[^\\\.\s\;\,\:\/])([\\\.\s\;\,\:\/]+$)', text):
        output =  re.search('(.*[^\\\.\s\;\,\:\/])([\\\.\s\;\,\:\/]+$)', text).group(1)
    return output

#Removes duplicates from list while preserving order
def unique(items):
    found = set()
    keep = []
    for item in items:
        if item not in found:
            found.add(item)
            keep.append(item)
    return keep

#Removes the prefix before the numerical value
def remove_prefix(text):
    output = text
    while re.search('(?:\([Oo][Cc][Oo][Ll][Cc]\)|[Oo][Cc][Mm]|[Oo][Cc][Nn]|[Oo][Nn])', output):
        if re.search('(.*)(?:\([Oo][Cc][Oo][Ll][Cc]\)|[Oo][Cc][Mm]|[Oo][Cc][Nn]|[Oo][Nn])(.+$)', output):
            output = re.search('(.*)(?:\([Oo][Cc][Oo][Ll][Cc]\)|[Oo][Cc][Mm]|[Oo][Cc][Nn]|[Oo][Nn])(.+$)', output).group(1) + re.search('(.*)(?:\([Oo][Cc][Oo][Ll][Cc]\)|[Oo][Cc][Mm]|[Oo][Cc][Nn]|[Oo][Nn])(.+$)', output).group(2)
    return output

#Removes all subfields other than subfield a
def remove_subfields(input_field):
    out_field = pymarc.Field(tag = input_field.tag, indicators = [input_field.indicators[0], input_field.indicators[1]])
    sub_last = len(input_field.subfields)
    sub = 0
    while input_field.subfields != None and sub < sub_last:
        if re.search('(^[a]$)', input_field.subfields[sub][0]):
            inst = re.search('(^[a]$)', input_field.subfields[sub][0]).group(1)
            temp_subfield = input_field.get_subfields(inst)
            for occurrence in range(len(temp_subfield)):
                out_field.add_subfield(inst, temp_subfield[occurrence])
            for occurrence in temp_subfield:
                input_field.delete_subfield(inst)
            sub_last = len(input_field.subfields)
        else:
            sub += 1
    return out_field

#Removes given subfield
def remove_subfield(input_field, sub_id):
    out_field = pymarc.Field(tag = input_field.tag, indicators = [input_field.indicators[0], input_field.indicators[1]])
    sub_last = len(input_field.subfields)
    sub = 0
    while input_field.subfields != None and sub < sub_last:
        sub_search = re.search('(^' + sub_id+ '$)', input_field.subfields[sub][0])
        if sub_search:
            inst = sub_search.group(1)
            temp_subfield = input_field.get_subfields(inst)
            for occurrence in temp_subfield:
                input_field.delete_subfield(inst)
            sub_last = len(input_field.subfields)
        else:
            out_field.add_subfield(input_field.subfields[sub], input_field.subfields[sub + 1])
            sub += 2
    return out_field


#Modifies subfields e and 4
def format_author_field(input_field):
    out_field = pymarc.Field(tag = input_field.tag, indicators = [input_field.indicators[0], input_field.indicators[1]])
    sub_dups = []
    sub_last = len(input_field.subfields)
    sub = 0
    while input_field.subfields != None and sub < sub_last:
        if re.search('(^[e]$)', input_field.subfields[sub][0]):
            inst = re.search('(^[e]$)', input_field.subfields[sub][0]).group(1)
            temp_subfield = input_field.get_subfields(inst)
            for occurrence in range(len(temp_subfield)):
                if not re.search('(author)', temp_subfield[occurrence]):
                    fixed = fix_end_char(temp_subfield[occurrence]).lower()
                    if fixed in author_dict:
                        fixed = author_dict[fixed]
                    if fixed not in sub_dups:
                        out_field.add_subfield(inst, '(' + fixed + ')')
                        sub_dups.append(fixed)
            for occurrence in temp_subfield:
                input_field.delete_subfield(inst)
            sub_last = len(input_field.subfields)
        elif re.search('(^[4]$)', input_field.subfields[sub][0]):
            inst = re.search('(^[4]$)', input_field.subfields[sub][0]).group(1)
            temp_subfield = input_field.get_subfields(inst)
            for occurrence in range(len(temp_subfield)):
                if not re.match('(^aut)', temp_subfield[occurrence]):
                    fixed = author_dict[fix_end_char(temp_subfield[occurrence])]
                    if fixed not in sub_dups:
                        out_field.add_subfield(inst, '(' + fixed + ')')
                        sub_dups.append(fixed)
            for occurrence in temp_subfield:
                input_field.delete_subfield(inst)
            sub_last = len(input_field.subfields)
        else:
            if sub + 2 < sub_last:
                if re.search('([^A-Za-z]+[A-Za-z]\.[^A-Za-z\d]*$)', input_field.subfields[sub + 1]):
                   out_field.add_subfield(input_field.subfields[sub], fix_end_char(input_field.subfields[sub + 1]) + '.,')
                else:
                   out_field.add_subfield(input_field.subfields[sub], fix_end_char(input_field.subfields[sub + 1]) + ',')
            else:
                if re.search('([^A-Za-z]+[A-Za-z]\.[^A-Za-z\d]*$)', input_field.subfields[sub + 1]):
                   out_field.add_subfield(input_field.subfields[sub], input_field.subfields[sub + 1])
                else:
                   out_field.add_subfield(input_field.subfields[sub], fix_end_char(input_field.subfields[sub + 1]))
            sub += 2
    return out_field

def fix_245_field(input_field):
    out_field = pymarc.Field(tag = '245', indicators = [input_field.indicators[0], input_field.indicators[1]])
    sub_last = len(input_field.subfields)
    sub = 0
    while input_field.subfields != None and sub < sub_last:
        if re.search('(^[abpn]$)', input_field.subfields[sub][0]):
            if re.search('(^[a]$)', input_field.subfields[sub][0]) and re.search('(\")', input_field.subfields[sub + 1]):
                if input_field['h'] is not None:
                    for subf in input_field.get_subfields('h'):
                        if re.search('(\")', subf):
                            input_field.subfields[sub + 1] = input_field.subfields[sub + 1] + '"'
            inst = re.search('(^[abpn]$)', input_field.subfields[sub][0]).group(1)
            temp_subfield = input_field.get_subfields(inst)
            for occurrence in range(len(temp_subfield)):
                out_field.add_subfield(inst, temp_subfield[occurrence])
            for occurrence in temp_subfield:
                input_field.delete_subfield(inst)
        else:
            inst = re.search('(^.$)', input_field.subfields[sub][0]).group(1)
            temp_subfield = input_field.get_subfields(inst)
            for occurrence in temp_subfield:
                input_field.delete_subfield(inst)
        sub_last = len(input_field.subfields)
    return out_field

#Adds a leading zero if given a single digit number, returns the original number otherwise.
def pad(num):
    num_str = str(num)
    if len(num_str) == 1:
        num_str = '0' + num_str
    return num_str

#Gets the date with no separators: yyyymmdd
def get_date():
    output = str(time.localtime().tm_year) + pad(time.localtime().tm_mon) + pad(time.localtime().tm_mday)
    return output

def get_oclc_number(record):
    oclc_num = ''
    if record['001'] is not None:
        for f in record.get_fields('001'):
            if re.search('(?:\([Oo][Cc][Oo][Ll][Cc]\)|[Oo][Cc][Mm]|[Oo][Cc][Nn]|[Oo][Nn])', f.value()):
                oclc_num = remove_prefix(f.value())
                if re.search('(^\d+)(?:\s+$)', oclc_num):
                    oclc_num = re.search('(^\d+)(?:\s+$)', oclc_num).group(1)
                    break
                elif re.search('(^\d+$)', oclc_num):
                    break
            elif re.search('(^\d+\s*$)', f.value()):
                if re.search('(^\d+)(?:\s+$)', f.value()):
                    oclc_num = re.search('(^\d+)(?:\s+$)', f.value()).group(1)
                    break
                oclc_num = f.value()
                break
    elif record['035'] is not None and record['035']['a'] is not None:
        for f in record.get_fields('035'):
            if f['a'] is not None:
                if re.search('(?:\([Oo][Cc][Oo][Ll][Cc]\)|[Oo][Cc][Mm]|[Oo][Cc][Nn]|[Oo][Nn])', f['a']):
                    oclc_num = remove_prefix(f['a'])
                    if re.search('(^\d+)(?:\s+$)', oclc_num):
                        oclc_num = re.search('(^\d+)(?:\s+$)', oclc_num).group(1)
                        break
                    if re.search('(^\d+$)', oclc_num):
                        break
                elif re.search('(^\d+\s*$)', f['a']):
                    if re.search('(^\d+)(?:\s+$)', f['a']):
                        oclc_num = re.search('(^\d+)(?:\s+$)', f['a']).group(1)
                        break
                    oclc_num = f['a']
                    break
    return oclc_num

#Extract marc record for given oclc number from Worldcat.
def process_oclc(oclc_num, input_uuid, input_collection = None):
    marc_record = next(get_marc_records.get_marc_worldcat(marc_from_oclc(oclc_num)))
    return process_marc_file(marc_record, None, input_uuid, input_collection)

#Extract marc record for given bib number from  Millennium.
def process_bib_num(input_bib_num, input_uuid, input_collection = None):
#    print(input_bib_num)
    bib_num = ''
    if len(str(input_bib_num)) == 9:
        bib_num = input_bib_num[:8]
    elif len(str(input_bib_num)) == 8:
        bib_num = input_bib_num
    html_url = 'http://catalog.crl.edu/search~S1?/.' + bib_num + '/.' + bib_num + '/1%2C1%2C1%2CB/marc~' + bib_num
    req = urllib.request.Request(html_url, headers={'User-Agent': 'Mozilla/5.0'})
    webpage = urllib.request.urlopen(req).read()
    soup = BeautifulSoup(webpage, 'html.parser')
    marc = soup.find_all('pre')
    marc_record_millennium = ''
    for a in marc:
        if re.match('\nLEADER', a.getText()):
            marc_record_millennium = a.getText()
            break
    marc_record = next(get_marc_records.get_marc_millennium((marc_record_millennium)))
    return process_marc_file(marc_record, input_bib_num, input_uuid, input_collection)

def remove_ending_comma(text):
    if re.match('(.+)(,\s*$)', text):
        text = re.match('(.+)(,\s*$)', text).group(1)
    return text

def process_marc_file(marc_record, input_bib_num, input_uuid, input_collection = None, extention='.csv'):
    record = marc_record
#    descriptive_metadata = []
    ddsnext_uuid = ''
    issn_isbn = ''
    tid = ''
    collection = ''
    material_type = ''
    original_format = ''
    holdings = ''
    print_to = 'general'
    external_link = ''
    if input_bib_num is not None:
        catalog_link = 'http://catalog.crl.edu/record=' + input_bib_num[:8]
        bib_num_id = input_bib_num
        if re.match('(?:b)(.+$)', input_bib_num):
            bib_num_id = re.match('(?:b)(.+$)', input_bib_num).group(1)
    else:
        catalog_link = None
        bib_num_id = None
    location = ''
    related_formats = ''
    oclc_number = ''
    country = False
    languages = []
    subjects = []
    publisher = []
    row = 0
    original_marc = record.as_marc()
    ddsnext_uuid = input_uuid
    title_items = {'title_name' : None, 'title_uuid' : ddsnext_uuid, 'title_material_type' : None, 'title_format' : None, 'title_oclc' : None, 'title_digital_holding_range' : None, 'title_resolution' : None, 'title_color_depth' : None, 'title_location_code' : None, 'title_catalog_link' : catalog_link, 'title_external_link' : None}
    descriptive_items = {}
    kbart_items = {'publication_title' : None, 'print_identifier' : None, 'online_identifier' : None, 'date_first_issue_online' : None, 'num_first_vol_online' : None, 'num_first_issue_online' : None, 'date_last_issue_online' : None, 'num_last_vol_online' : None, 'num_last_issue_online' : None, 'title_url' : None, 'first_author' : None, 'title_id' : None, 'embargo_info' : None, 'coverage_depth' : None, 'coverage_notes' : None, 'publisher_name' : None, 'location' : None, 'title_notes' : None, 'staff_notes' : None, 'vendor_id' : None, 'oclc_collection_name' : None, 'oclc_collection_id' : None, 'oclc_entry_id' : bib_num_id, 'oclc_linkscheme' : None, 'oclc_number' : None, 'action' : 'raw'}
    if input_collection is not None:
        if re.match('(.+)(?:\r)(.*$)', input_collection):
            print(input_bib_num, input_uuid, input_collection)
            input_collection = re.match('(.+)(?:\r)(.*$)', input_collection).group(1) + re.match('(.+)(?:\r)(.*$)', input_collection).group(2)
        kbart_items['coverage_depth'] = collection_dict[input_collection]['coverage_depth']
        kbart_items['oclc_collection_name'] = collection_dict[input_collection]['oclc_collection_name']
        kbart_items['oclc_collection_id'] = collection_dict[input_collection]['oclc_collection_id']
        if re.match('(monograph$)', input_collection.lower()) and record['260'] is not None and record['260']['c'] is not None:
            if re.search('(\d{4})', record['260']['c']):
                kbart_items['date_first_issue_online'] = re.search('(\d{4})', record['260']['c']).group(1)
    #Extracts title
    if record['245'] is not None:
        output_title = fix_end_char(fix_245_field(record['245']).value())
        title_items['title_name'] = output_title
        kbart_items['publication_title'] = output_title
    #Extracts title
    elif record['222'] is not None:
#        if record['222']['b'] is not None:
#            record['222'].delete_subfield('b')
#        output_title = record['222'].value()
        output_title = remove_subfield(record['222'], 'b').value()
        title_items['title_name'] = output_title
        kbart_items['publication_title'] = output_title
    #Extracts and prints series title (CRL collection title variant)
    if record['246'] is not None:
        if record['246']['i'] is not None and record['246']['a'] is not None and re.search('CRL collection title', record['246']['i']):
#            dds_descriptive_writer.writerow([ddsnext_uuid, 'series', record['246']['i'] + ' ' + record['246']['a']])
#            descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'series', 'value' : record['246']['i'] + ' ' + record['246']['a']})
            descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'series', 'value' : record['246']['i'] + ' ' + record['246']['a']}
            row += 1
    #Extracts holdings and title url.
    if record['856'] is not None and record['856']['u'] is not None:
        for f in record.get_fields('856'):
            for sub in f.get_subfields('u'):
                if f['z'] is None or f['z'] is not None and not re.search('(?:[Gg][Uu][Ii][Dd][Ee])', f['z']):
                    if re.search('(?:.*ddsnext\.crl\.edu\/titles\/)(\d+)', sub):
                        kbart_items['title_url'] = sub
                        if f['z'] is not None and re.search('(?:.\:\s*)(.+)', f['z']):
                            holdings = re.search('(?:.\:\s*)(.+)', f['z']).group(1)
                            title_items['title_digital_holding_range'] = re.search('(?:.\:\s*)(.+)', f['z']).group(1)
                        elif f['3'] is not None and re.search('(?:.\:\s*)(.+)', f['3']):
                            holdings = re.search('(?:.\:\s*)(.+)', f['3']).group(1)
                            title_items['title_digital_holding_range'] = re.search('(?:.\:\s*)(.+)', f['3']).group(1)
    #Extracts and prints LCCN
    if record['010'] is not None:
#        dds_descriptive_writer.writerow([ddsnext_uuid, 'LCCN', record['010'].value()])
#        descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'LCCN', 'value' : record['010'].value()})
        descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'LCCN', 'value' : record['010'].value()}
        row += 1
    #Extracts and prints ISBN
    if record['020'] is not None:
#        dds_descriptive_writer.writerow([ddsnext_uuid, 'ISBN', record['020'].value()])
#        descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'ISBN', 'value' : record['020'].value()})
        descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'ISBN', 'value' : record['020'].value()}
        kbart_items['online_identifier'] = record['020']
        row += 1
    #Extracts and prints ISSN
    if record['022'] is not None and record['022']['a'] is not None:
#        dds_descriptive_writer.writerow([ddsnext_uuid, 'ISSN', record['022']['a']])
#        descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'ISSN', 'value' : record['022']['a']})
        descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'ISSN', 'value' : record['022']['a']}
        kbart_items['online_identifier'] = record['022']['a']
        row += 1
    #Extracts and prints ISSN
    if record['022'] is not None and record['022']['l'] is not None:
        kbart_items['print_identifier'] = record['022']['l']
    #Extracts and prints OCLC number
    oclc_number = get_oclc_number(record)
    if oclc_number != '':
        title_items['title_oclc'] = oclc_number
        kbart_items['oclc_number'] = oclc_number
    #Extracts and prints dissertation description
    if record['502'] is not None:
        description = ''
        #Extracts the dissertation description from the 502$a
        if record['502']['a'] is not None:
            description = record['502']['a']
            if re.match('(?:.*\()(.+)(?:\)\-\-))', record['502']['a']):
                description = description + re.match('(?:.*\()(.+)(?:\(\-\-))', record['502']['a']).group(1)
            if re.match('(?:.*\(\)\-\-)(.+)(?:\,)(\d{4})(?:\S*$))', record['502']['a']):
                if description != '':
                    description = description + ' '
                description = description + re.match('(?:.*\(\)\-\-)(.+)(?:\,)(\d{4})(?:\S*$))', record['502']['a']).group(1) + ' ' + re.match('(?:.*\(\)\-\-)(.+)(?:\,)(\d{4})(?:\S*$))', record['502']['a']).group(2)
            elif re.match('(?:.*\(\)\-\-)(.+)(?:\S*$))', record['502']['a']):
                if description != '':
                    description = description + ' '
                description = description + re.match('(?:.*\(\)\-\-)(.+)(?:\S*$))', record['502']['a']).group(1)
        #Extracts the dissertation description from the 502$b, 502$c, 502$d, and 502$g
        else:
            if record['502']['g'] is not None:
                description = record['502']['g']
            if record['502']['b'] is not None:
                if description != '':
                    description = description + ' '
                description = description + record['502']['b']
            if record['502']['c'] is not None:
                if description != '':
                    description = description + ' '
                description = description + record['502']['c']
            if record['502']['d'] is not None:
                if description != '':
                    description = description + ' '
                description = description + record['502']['d']
#        dds_descriptive_writer.writerow([ddsnext_uuid , 'description', description])
#        descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'description', 'value' : description})
        descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'description', 'value' : description}
        row += 1
    #Extracts and prints description
    if record['520'] is not None and record['520']['a'] is not None:
        description = record['520']['a']
        if record['520']['b'] is not None:
            if description != '':
                description = description + ' '
            description = description + record['520']['b']
#        dds_descriptive_writer.writerow([ddsnext_uuid , 'description', description])
#        descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'description', 'value' : description})
        descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'description', 'value' : description}
        row += 1
    #Extracts language
    if record['041'] is not None:
        if record['041']['a'] is not None:
            for sub in record['041'].get_subfields('a'):
                if sub in language_dict:
                    languages.append(language_dict[sub])
    #Extracts country
    #Extracts language
    if record['998'] is not None:
        if record['998']['g'] is not None:
            if fix_end_char(record['998']['g']) in country_dict and fix_end_char(record['998']['g']) != 'xx':
#                dds_descriptive_writer.writerow([ddsnext_uuid , 'country', country_dict[fix_end_char(record['998']['g'])]])
#                descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'country', 'value' : country_dict[fix_end_char(record['998']['g'])]})
                descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'country', 'value' : country_dict[fix_end_char(record['998']['g'])]}
                row += 1
                country = True
        if record['998']['f'] is not None:
            if record['998']['f'] in language_dict:
                if record['041'] is None or record['041'] is not None and (record['041']['a'] is None or (record['041']['a'] is not None and record['998']['f'] not in record['041'].get_subfields('a'))):
                    languages.append(language_dict[record['998']['f']])
    #Extracts and prints country if not found in the 998 field
    #Extracts language
    if record['008'] is not None:
        if not country:
            if re.search('([a-zA-Z]+)', record['008'].value()[15:18]):
                country_code = re.search('([a-zA-Z]+)', record['008'].value()[15:18]).group(1)
            if fix_end_char(country_code) in country_dict and fix_end_char(country_code) != 'xx':
#                dds_descriptive_writer.writerow([ddsnext_uuid , 'country', country_dict[fix_end_char(country_code)]])
#                descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'country', 'value' : country_dict[fix_end_char(country_code)]})
                descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'country', 'value' : country_dict[fix_end_char(country_code)]}
                row += 1
        if record['008'].value()[35:38] in language_dict:
            languages.append(language_dict[record['008'].value()[35:38]])
    #Deduplicate languages
    if languages != []:
        languages = unique(languages)
    #Prints languages
    for language in languages:
#        dds_descriptive_writer.writerow([ddsnext_uuid , 'language' , language])
#        descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'language', 'value' : language})
        descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'language', 'value' : language}
        row += 1
    #Extracts and prints coverage (country)
    if record['752'] is not None and record['752']['a'] is not None:
#        dds_descriptive_writer.writerow([ddsnext_uuid , 'coverage' , record['752']['a']])
#        descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'coverage', 'value' : record['752']['a']})
        descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'coverage', 'value' : record['752']['a']}
        row += 1
    #Extracts and prints call number if record is an electronic Millennium record.
    if record['099'] is not None:
        if record['099']['a'] is not None:
            for sub in record['099'].get_subfields('a'):
                if sub != 'Internet resource' and sub != 'ediss' and sub != 'MF' and sub != 'Electronic version' and sub != 'Electronic resource/e' and sub != 'TOSS' and not re.search('[Ee][Ll][Ee][Cc]?[Tt][Rr][Oo][Nn][Ii][Cc]', sub):
#                    dds_descriptive_writer.writerow([ddsnext_uuid, 'resource_identifier', sub])
#                    descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'resource_identifier', 'value' : sub})
                    descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'resource_identifier', 'value' : sub}
                    row += 1
    #Extracts and prints creator
    #Extracts subject (title)
    if record['100'] is not None:
        for f in record.get_fields('100'):
            if f['t'] is not None:
                subjects.append(f['t'])
                f.delete_subfield('t')
            author = remove_ending_comma(remove_subfield(format_author_field(f), '6').value())
#            if re.match('(.+)(,\s*$)', author):
#                author = re.match('(.+)(,\s*$)', author).group(1)
#            dds_descriptive_writer.writerow([ddsnext_uuid , 'creator' , author])
#            descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'creator', 'value' : author})
            descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'creator', 'value' : author}
            row += 1
            kbart_items['first_author'] = remove_ending_comma(record['100']['a'])
    #Extracts and prints creator
    #Extracts subject (title)
    if record['110'] is not None:
        for f in record.get_fields('110'):
            if f['t'] is not None:
                subjects.append(f['t'])
                f.delete_subfield('t')
            author = remove_subfield(format_author_field(f), '6').value()
#            dds_descriptive_writer.writerow([ddsnext_uuid , 'creator' , author])
#            descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'creator', 'value' : author})
            descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'creator', 'value' : author}
            row += 1
    #Extracts and prints creator
    #Extracts subject (title)
    if record['111'] is not None:
        for f in record.get_fields('111'):
            if f['t'] is not None:
                subjects.append(f['t'])
                f.delete_subfield('t')
            author = remove_subfield(format_author_field(f), '6').value()
#            dds_descriptive_writer.writerow([ddsnext_uuid , 'creator' , author])
#            descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'creator', 'value' : author})
            descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'creator', 'value' : author}
            row += 1
    #Extracts and prints creator
    #Extracts subject (title)
    if record['700'] is not None:
        for f in record.get_fields('700'):
            if f['t'] is not None:
                subjects.append(f['t'])
                f.delete_subfield('t')
            author = remove_ending_comma(remove_subfield(format_author_field(f), '6').value())
#            if re.match('(.+)(,\s*$)', author):
#                author = re.match('(.+)(,\s*$)', author).group(1)
#            dds_descriptive_writer.writerow([ddsnext_uuid , 'creator' , author])
#            descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'creator', 'value' : author})
            descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'creator', 'value' : author}
            row += 1
    #Extracts and prints creator
    #Extracts subject (title)
    if record['710'] is not None:
        for f in record.get_fields('710'):
            if f['t'] is not None:
                subjects.append(f['t'])
                f.delete_subfield('t')
            author = remove_subfield(format_author_field(f), '6').value()
#            dds_descriptive_writer.writerow([ddsnext_uuid , 'creator' , author])
#            descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'creator', 'value' : author})
            descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'creator', 'value' : author}
            row += 1
    #Extracts and prints creator
    #Extracts subject (title)
    if record['711'] is not None:
        for f in record.get_fields('711'):
            if f['t'] is not None:
                subjects.append(f['t'])
                f.delete_subfield('t')
            author = remove_subfield(format_author_field(f), '6').value()
#            dds_descriptive_writer.writerow([ddsnext_uuid , 'creator' , author])
#            descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'creator', 'value' : author})
            descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'creator', 'value' : author}
            row += 1
    #Extracts publishers
    if record['260'] is not None:
        for f in record.get_fields('260'):
            if f['b'] is not None:
                for sub in f.get_subfields('b'):
                    if not (re.search('(\[?publisher not identified\]?)', sub) or re.search('(\[?[Ss]\.n\.\??\]?)', sub)):
                        publisher.append(fix_end_char(sub))
    #Extracts publishers
    elif record['264'] is not None:
        for f in record.get_fields('264'):
            if f.indicator2 == '1' and f['b'] is not None:
                for sub in f.get_subfields('b'):
                    if not (re.search('(\[?publisher not identified\]?)', sub) or re.search('(\[?s\.n\.\]?)', sub)):
                        publisher.append(fix_end_char(sub))
    #Deduplicate publishers
    if publisher != []:
        publisher = unique(publisher)
    #Prints publishers
    for pub in publisher:
#        dds_descriptive_writer.writerow([ddsnext_uuid , 'publisher' , pub])
#        descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'publisher', 'value' : pub})
        descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'publisher', 'value' : pub}
        row += 1
        if kbart_items['publisher_name'] is None:
            kbart_items['publisher_name'] = pub
    #Extracts and prints series
    #Extracts and prints author
    if record['800'] is not None:
        for f in record.get_fields('800'):
            f = remove_subfield(f, '6')
            if f['t'] is not None:
#                dds_descriptive_writer.writerow([ddsnext_uuid , 'series' , f['a']])
#                descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'series', 'value' : f['a']})
                descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'series', 'value' : f['a']}
                row += 1
                f.delete_subfield('t')
            if f['v'] is not None:
                f.delete_subfield('v')
            author = remove_ending_comma(remove_subfield(format_author_field(f), '6').value())
#            if re.match('(.+)(,\s*$)', author):
#                author = re.match('(.+)(,\s*$)', author).group(1)
#            dds_descriptive_writer.writerow([ddsnext_uuid , 'creator' , author])
#            descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'creator', 'value' : author})
            descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'creator', 'value' : author}
            row += 1
    #Extracts and prints series
    #Extracts and prints author
    if record['810'] is not None:
        for f in record.get_fields('810'):
            f = remove_subfield(f, '6')
            if f['t'] is not None:
#                dds_descriptive_writer.writerow([ddsnext_uuid , 'series' , f['a']])
#                descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'series', 'value' : f['a']})
                descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'series', 'value' : f['a']}
                row += 1
                f.delete_subfield('t')
            if f['v'] is not None:
                f.delete_subfield('v')
            author = remove_subfield(format_author_field(f), '6').value()
#            dds_descriptive_writer.writerow([ddsnext_uuid , 'creator' , author])
#            descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'creator', 'value' : author})
            descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'creator', 'value' : author}
            row += 1
    #Extracts and prints series
    #Extracts and prints author
    if record['811'] is not None:
        for f in record.get_fields('811'):
            f = remove_subfield(f, '6')
            if f['t'] is not None:
#                dds_descriptive_writer.writerow([ddsnext_uuid , 'series' , f['a']])
#                descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'series', 'value' : f['a']})
                descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'series', 'value' : f['a']}
                row += 1
                f.delete_subfield('t')
            if f['v'] is not None:
                f.delete_subfield('v')
            author = remove_subfield(format_author_field(f), '6').value()
#            dds_descriptive_writer.writerow([ddsnext_uuid , 'creator' , author])
#            descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'creator', 'value' : author})
            descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'creator', 'value' : author}
            row += 1
    #Extracts and prints series
    if record['830'] is not None:
        for f in record.get_fields('830'):
            f = remove_subfield(f, '6')
            if f['a'] is not None:
#                dds_descriptive_writer.writerow([ddsnext_uuid , 'series' , f['a']])
#                descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'series', 'value' : f['a']})
                descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'series', 'value' : f['a']}
                row += 1
    #Extracts and prints subject (author)
    #Extracts and prints subject (title)
    if record['600'] is not None:
        for f in record.get_fields('600'):
            f = remove_subfield(f, '6')
            if f.indicator2 == '0':
                if f['t'] is not None:
                    subjects.append(f['t'])
                    f.delete_subfield('t')
                if f['v'] is not None:
                    subjects.append(f['v'])
                    f.delete_subfield('v')
                subjects.append(format_author_field(f).value())
    #Extracts and prints subject (author)
    #Extracts and prints subject (title)
    if record['610'] is not None:
        for f in record.get_fields('610'):
            f = remove_subfield(f, '6')
            if f.indicator2 == '0':
                if f['t'] is not None:
                    subjects.append(f['t'])
                    f.delete_subfield('t')
                if f['v'] is not None:
                    subjects.append(f['v'])
                    f.delete_subfield('v')
                subjects.append(format_author_field(f).value())
    #Extracts and prints subject (author)
    #Extracts and prints subject (title)
    if record['611'] is not None:
        for f in record.get_fields('611'):
            f = remove_subfield(f, '6')
            if f.indicator2 == '0':
                if f['t'] is not None:
                    subjects.append(f['t'])
                    f.delete_subfield('t')
                if f['v'] is not None:
                    subjects.append(f['v'])
                    f.delete_subfield('v')
                subjects.append(format_author_field(f).value())
    #Extracts and prints subject
    if record['630'] is not None:
        for f in record.get_fields('630'):
            f = remove_subfield(f, '6')
            if f.indicator2 == '0':
                for sub in range(len(f.subfields)):
                    if sub % 2 != 0:
                        subjects.append(f.subfields[sub])
    #Extracts and prints subject
    if record['650'] is not None:
        for f in record.get_fields('650'):
            f = remove_subfield(f, '6')
            if f.indicator2 == '0':
                for sub in range(len(f.subfields)):
                    if sub % 2 != 0:
                        subjects.append(f.subfields[sub])
    #Extracts and prints subject
    if record['651'] is not None:
        for f in record.get_fields('651'):
            f = remove_subfield(f, '6')
            if f.indicator2 == '0':
                for sub in range(len(f.subfields)):
                    if sub % 2 != 0:
                        subjects.append(f.subfields[sub])
    #Extracts and prints original author and original title
    #Extracts subject (title)
    if record['880'] is not None:
        for f in record.get_fields('880'):
            if f['6'] is not None:
                for sub in f.get_subfields('6'):
                    #Extracts and prints original author
                    if re.search('(^[17][01][01]-)', sub):
                        if f['t'] is not None:
                            subjects.append(f['t'])
                            f.delete_subfield('t')
                        if f['v'] is not None:
                            f.delete_subfield('v')
                        orig_author = remove_subfield(format_author_field(f), '6').value()
                        if re.match('(.+)(,\s*$)', orig_author):
                            orig_author = re.match('(.+)(,\s*$)', orig_author).group(1)
#                        dds_descriptive_writer.writerow([ddsnext_uuid , 'orig_author' , orig_author])
#                        descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'orig_author', 'value' : orig_author})
                        descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'orig_author', 'value' : orig_author}
                        row += 1
                    #Extracts and prints original title
                    if re.search('(^245-)', sub):
                        orig_title = fix_end_char(fix_245_field(f).value())
#                        dds_descriptive_writer.writerow([ddsnext_uuid , 'orig_title' , orig_title])
#                        descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'orig_title', 'value' : orig_title})
                        descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'orig_title', 'value' : orig_title}
                        row += 1
    for n in range(len(subjects)):
        subjects[n] = fix_end_char(subjects[n])
    #Deduplicate subjects
    if subjects != []:
        subjects = unique(subjects)
    #Prints subjects
    for s in subjects:
        if not ((material_type == '2' or (material_type == '5' and (location == 'fdocs' or location == 'fogse'))) and fix_end_char(s) == 'Periodicals') and not (material_type == '3' and fix_end_char(s) == 'Newspapers'):
#            dds_descriptive_writer.writerow([ddsnext_uuid , 'subject' , fix_end_char(s)])
#            descriptive_items.append({'title_uuid' : ddsnext_uuid, 'field' : 'subject', 'value' : fix_end_char(s)})
            descriptive_items[row] = {'title_uuid' : ddsnext_uuid, 'field' : 'subject', 'value' : fix_end_char(s)}
            row += 1
    return [marc_record, title_items, descriptive_items, kbart_items]

#def process_records(record_source, file_types, input_values):
#    marc_record_dict = {}
#    title_metadata = {}
#    descriptive_metadata = {}
#    kbart_metadata = {}
#    print(input_values)
#    for input_value in input_values:
#        title_items = None
#        descriptive_items = None
#        kbart_items = None
#        if record_source == 'worldcat':
#            marc_record, title_items, descriptive_items, kbart_items = process_oclc(input_value[0], input_value[1])
#        else:
#            print(input_value[0], input_value[1])
#            print(process_bib_num(input_value[0], input_value[1]))
#            marc_record, title_items, descriptive_items, kbart_items = process_bib_num(input_value[0], input_value[1])
#            print(descriptive_items)
#        marc_record_dict[input_value[0]] = marc_record
#        title_metadata[input_value[0]] = title_items
#        descriptive_metadata[input_value[0]] = descriptive_items
#        kbart_metadata[input_value[0]] = kbart_items
#    #
#    descriptive_metadata_file = open(os.path.join(output_folder, 'ddsnext_descriptive_metadata_' + get_date() + file_types_dict[file_types]),'wt',encoding='utf8', newline='')
#    if file_types_dict[file_types] == '.tsv':
#        dds_descriptive_writer = csv.writer(descriptive_metadata_file, delimiter='\t', dialect='excel')
#    else:
#        dds_descriptive_writer = csv.writer(descriptive_metadata_file, dialect='excel')
#    dds_descriptive_writer.writerow(['title_uuid' , 'field' , 'value'])
#    print(descriptive_items)
#    for key in descriptive_metadata:
#        for descriptive_metadata_field in descriptive_metadata[key]:
#            dds_descriptive_writer.writerow([descriptive_metadata_field['title_uuid'], descriptive_metadata_field['field'], descriptive_metadata_field['value']])
#    
#    title_metadata_file = open(os.path.join(output_folder, 'ddsnext_title_metadata_' + get_date() + file_types_dict[file_types]),'wt',encoding='utf8', newline='')
#    if file_types_dict[file_types] == '.tsv':
#        dds_title_writer = csv.writer(title_metadata_file, delimiter='\t', dialect='excel')
#    else:
#        dds_title_writer = csv.writer(title_metadata_file, dialect='excel')
#    dds_title_writer.writerow(['Title Name', 'Title UUID', 'Title Material Type', 'Title Format', 'Title OCLC', 'Title Digital Holding Range', 'Title Resolution', 'Title Color Depth', 'Title Location Code', 'Title Catalog Link', 'Title External Link'])
#    print(title_metadata)
#    for key in title_metadata:
#        print(key, title_metadata[key])
#        dds_title_writer.writerow([title_metadata[key]['title_name'], title_metadata[key]['title_uuid'], title_metadata[key]['title_material_type'], title_metadata[key]['title_format'], title_metadata[key]['title_oclc'], title_metadata[key]['title_digital_holding_range'], title_metadata[key]['title_resolution'], title_metadata[key]['title_color_depth'], title_metadata[key]['title_location_code'], title_metadata[key]['title_catalog_link'], title_metadata[key]['title_external_link']])
#        
#    kbart_metadata_file = open(os.path.join(output_folder, 'worldshare_metadata_' + get_date() + file_types_dict[file_types]),'wt',encoding='utf8', newline='')
#    if file_types_dict[file_types] == '.tsv':
#        worldshare_writer = csv.writer(kbart_metadata_file, delimiter='\t', dialect='excel')
#    else:
#        worldshare_writer = csv.writer(kbart_metadata_file, dialect='excel')
#    worldshare_writer.writerow(['publication_title', 'print_identifier', 'online_identifier', 'date_first_issue_online', 'num_first_vol_online', 'num_first_issue_online', 'date_last_issue_online', 'num_last_vol_online', 'num_last_issue_online', 'title_url', 'first_author', 'title_id', 'embargo_info', 'coverage_depth', 'coverage_notes', 'publisher_name', 'location', 'title_notes', 'staff_notes', 'vendor_id', 'oclc_collection_name', 'oclc_collection_id', 'oclc_entry_id', 'oclc_linkscheme', 'oclc_number', 'ACTION'])
#    print(kbart_metadata)
#    for key in kbart_metadata:
#        worldshare_writer.writerow([kbart_metadata[key]['publication_title'], kbart_metadata[key]['print_identifier'], kbart_metadata[key]['online_identifier'], kbart_metadata[key]['date_first_issue_online'], kbart_metadata[key]['num_first_vol_online'], kbart_metadata[key]['num_first_issue_online'], kbart_metadata[key]['date_last_issue_online'], kbart_metadata[key]['num_last_vol_online'], kbart_metadata[key]['num_last_issue_online'], kbart_metadata[key]['title_url'], kbart_metadata[key]['first_author'], kbart_metadata[key]['title_id'], kbart_metadata[key]['embargo_info'], kbart_metadata[key]['coverage_depth'], kbart_metadata[key]['coverage_notes'], kbart_metadata[key]['publisher_name'], kbart_metadata[key]['location'], kbart_metadata[key]['title_notes'], kbart_metadata[key]['staff_notes'], kbart_metadata[key]['vendor_id'], kbart_metadata[key]['oclc_collection_name'], kbart_metadata[key]['oclc_collection_id'], kbart_metadata[key]['oclc_entry_id'], kbart_metadata[key]['oclc_linkscheme'], kbart_metadata[key]['oclc_number'], kbart_metadata[key]['action']])
#        
#    for key in list(cell_location):
#        if cell_location[key] is not None:
#            insert_into_table.append(line[cell_location[key]])
#        else:
#            insert_into_table.append(None)
#    insert_into_table = tuple(insert_into_table)
#    
#    for key in list(title_metadata[input_bib_num]):
#        
#    
#    descriptive_metadata_file.close()
#    title_metadata_file.close()
#    kbart_metadata_file.close()
#    return marc_record_dict

#process_bib_num()

if __name__ == "__main__":
    root = tkinter.Tk()
    app = Application(master=root)
    app.mainloop()
