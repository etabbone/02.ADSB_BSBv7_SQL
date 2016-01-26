#!/usr/bin/env python
from __future__ import print_function
import os, sys, time, getopt, json, sqlite3, xlrd, datetime, decimal, shutil, smtplib, csv, locale
from xlrd import XLRDError
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

os_win = False
if os.name == 'nt':
    os_win = True
if sys.version_info<(3,0,0):
    from urllib import FancyURLopener
else:
    from urllib.request import FancyURLopener
    
# Script date
script_date = "2015/11/05"
# Script time
script_time = "09:45"
# Script version
script_version = "1.7"
# Author
Author = 'Eric TABBONE'
# Author email
Author_email = 'etabbone@gmail.com'
# Author linkedin
Linkedin = 'https://br.linkedin.com/in/etabbone'
#Email sender
email_fromaddr = '___@gmail.com'
#Email password
email_password = '___'
# email recipients
email_toaddr = '___@gmail.com'
#Email subject
email_subject = 'ADSB Automatic report: '
#Email body
email_body = 'Last month automatic report'
#Email server address
email_server = 'smtp.gmail.com'
#Email port
email_port = 587
# Version
version = script_version + " - " + script_date + " " + script_time
# Script name
prog_name = 'DUMP1090 to SQLite parser'
# Script name (not dump1090sql mode)
prog_name_report = 'DUMP1090 Report utility'
# Script filename
exe_name = os.path.basename(sys.argv[0])
# default DUMP1090 JSON IP server 
server_IP = '127.0.0.1'
# default DUMP1090 JSON PORT number
server_PORT = '8080'
# Must filter datas before writing in txt file or sql?
exclude = False
# Config filename
config_file = exe_name.split('.')[0] + ".ini"
# Must write datas in sql base?
write2sql = False
# Look for date of the day
now = time.strftime("%Y%m%d%H%M%S", time.localtime())
today = time.strftime("%Y%m%d", time.localtime())
month = time.strftime("%Y%m", time.localtime())
year = time.strftime("%Y", time.localtime())
local_path = os.getcwd()
# Output txt filename
output_file = today + '_Flights.txt'
# database filename (sqlite3)
DB_file = month + '_Flights.db'
# Must write to file?
write2file = False
# Input xls 01DB type file
input_file = 'N/A'
# Must load 01DB file into database?
load2sql = False
# Be quiet! Do the job and show nothing
quiet = False
# Show all informations on screen
interactive = False
# Do not make pause after errors, warnings...etc
pause = False
# JSON Server connexion ip/port/file
net_HTTP = 'http://' + server_IP + ':' + server_PORT + '/data.json'
# JSON streaming
stream = ''
# First time JSON is reading?
first_JSON = True
# Latitude (DUMP1090 system)
latitude = 'N/A'
# Longitude (DUMP1090 system)
longitude = 'N/A'
# Altitude (DUMP1090 system)
altitude = 'N/A'
# Name station
station_name = 'N/A'
# Current index of new record in infos
idx_infos = 0
# Use metric units
metric = False
# log all actions on file
log = False
# logfile filename
log_file = now + '_logs.txt'
# SQL Connector
conn = ''      
# SQL cursor
cursor = ''           
# XLS workbook
wb = ''
# write logs in buffer before changing (or not) path
log_buffer = ''
# write screen messages in buffer before changing (or not) path
msg_buffer = ''
# Try again after connexion lost?
cnx_try = False
# Delay before trying again if connexion lost
cnx_delay = 0
# Should try to read stream again?
try_again = False
# first time generic indicator
first_time = False
# Test connexion with server
cnx_test = False
# zip
zip_backup = False
# zip filename
zip_filename = 'N/A'
# email
email_backup = False
# zip month of backup
zip_month = ''
# zip year of backup
zip_year = ''
# Compression ok?
zip_ok = False
# remove backup after zip success
remove = False
# make report with db and xls
report = False
# output report file
output_report = ''
# Maximum altitude to consider flight
max_altitude = 0
# Maximum distance from DUO station
max_distance = 0
# Change from ft to metric in command line
change_to_metric = False
# Open csv file
open_csv = False
# use external ddb flights
extDB = False
# external ddb flights filename
extDB_filename = 'N/A'
# Use old report to add flight
ireport = False
# Set ireport filename
ireport_filename = 'N/A'
# is dump1090sql
is_dump1090sql = False
# old sql access
onesql_report = False

def parseCmdLine(argv):
# Read cmd line and check options
    global prog_name
    global prog_name_report
    global exe_name
    global server_IP 
    global server_PORT
    global exclude
    global DB_file
    global write2sql
    global output_file
    global write2file
    global input_file
    global load2sql
    global quiet
    global interactive
    global version
    global net_HTTP
    global pause
    global do_help
    global do_version
    global latitude
    global longitude
    global altitude
    global station_name
    global metric
    global log
    global log_file
    global log_buffer
    global msg_buffer  
    global cnx_try   
    global cnx_delay 
    global cnx_test
    global zip_backup
    global zip_filename
    global email_backup
    global email_toaddr
    global month
    global remove
    global report
    global max_altitude
    global max_distance
    global change_to_metric
    global open_csv
    global extDB
    global extDB_filename
    global ireport
    global ireport_filename
    global is_dump1090sql
    global onesql_report
    global local_path
    new_report = False
    old_report = False
    try:
        if is_dump1090sql:
            opts, args = getopt.getopt(argv,"phvestiq",['pause','help','version','exclude','net-http-ip=','ip=','net-http-port=','port=','sql','sqlfile=','txt','txtfile=','ifile=','latitude=','lat=','longitude=','lon=','altitude=','alt=','sta=','station=','metric','quiet','interactive','log','delay=','zip=','email','remove','report','oldreport','maxalt=','maxdist=','opencsv','extdb=','ireport='])
        else:
            opts, args = getopt.getopt(argv,"phvestiq",['pause','help','version','sqlfile=','ifile=','metric','quiet','log','report','oldreport', 'maxalt=','maxdist=','opencsv','extdb=','ireport='])
    except getopt.GetoptError as Err:
    # Error reading cmd line (option does not exist...)
        print('Usage:' , exe_name , '[options]')
        print('      ' , exe_name , '-h for help')
        print('')
        print(exe_name + ': FATAL ERROR: ', str(Err))
        checkPause()
        sys.exit(2)
    for opt, arg in opts:
    # pause on?
        if opt in ("-p","--pause"):
            if not quiet:
                pause = True  
        if opt in ("--log"):
            log = True  
        if opt in ("--zip"):
            zip_backup = True  
            zip_filename = arg
    log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseCmdLine(): cmd line options ' + str(opts) + '\n'
    log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseCmdLine(): start parsing command line options\n'
    msg_buffer = msg_buffer + 'Start parsing command line options\n'
    for opt, arg in opts:
    # Show help or version then exit?
        if opt in ("--help", "-h"):
            if is_dump1090sql:
                print(prog_name , '- version' , version)
                print('')
                print('Usage:')
                print(exe_name , '[options]')
                print('')
                print('Options:')
                print('  -h, --help                 Show this help message and exit')
                print('      --ip "ip"              Set HTTP server IP (default: 127.0.0.1)')
                print('      --port "port"          Set HTTP server port (default: 8080)') 
                print('      --delay "seconds"      Set delay before reconnecting on lost connexion')
                print('  -s, --sql                  Insert into SQLite database')
                print('      --sqlfile "database"   Set SQLite database filename ')
                print('  -t, --txt                  Write to txt file')
                print('      --txtfile "file"       Set txt output filename')
                print('  -e, --exclude              Exclude inconsistent data') 
                print('      --maxalt "altitude"    Set maxalt altitude flight')
                print('      --maxdist "distance"   Set maximum distance between flight and station')
                print('      --ifile "file"         Load Excel "file" into database (need --sqlfile)')
                print('      --report               Make report from database (need --sqlfile)')
                print('      --extdb "database"     Add flights infos (need --report or --ireport)')
                print('      --ireport "file"       Add flights infos to report "file" (need --extdb)') 
                print('      --opencsv              Open csv report file (need --report)')
                print('      --zip "file"           Compress all files in one "zip file"')
                print('      --email                Send end of month email, (need --zip)')
                print('      --remove               Remove files after compression (need --zip)')
                print('      --lat "latitude"       Set latitude of local station (DD)')
                print('      --lon "longitude"      Set longitude of local station (DD)')
                print('      --alt "altitude"       Set altitude of local station (meters)')
                print('      --sta "name"           Set local station name')
                print('      --metric               Use metric units (meters, km/h, ...)')
                print('  -q, --quiet                Disable all output')
                print('  -i, --interactive          Show all informations')
                print('      --pause                Enable pauses during process')
                print('      --log                  Create a log file')
                print('  -v, --version              Show software version')
                print('')
                print('Type CTRL+C to exit')
            else:
                print(prog_name_report , '- version' , version)
                print('')
                print('Usage:')
                print(exe_name , '[options]')
                print('')
                print('Options:')
                print('  -h, --help                 Show this help message and exit')
                print('      --sqlfile "database"   Set SQLite database filename ')
                print('      --maxalt "altitude"    Set maxalt altitude flight')
                print('      --maxdist "distance"   Set maximum distance between flight and station')
                print('      --ifile "file"         Load Excel "file" into database (need --sqlfile)')
                print('      --report               Make report from database (need --sqlfile)')
                print('      --extdb "database"     Add flights infos (need --report or --ireport)')
                print('      --ireport "file"       Add flights infos to report "file" (need --extdb)') 
                print('      --opencsv              Open csv report file (need --report)')
                print('      --metric               Use metric units (meters, km/h, ...)')
                print('  -q, --quiet                Disable all output')
                print('      --pause                Enable pauses during process')
                print('      --log                  Create a log file')
                print('  -v, --version              Show software version')
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(log_buffer)
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseCmdLine(): show help\n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseCmdLine(): sys.exit(0): Normal E-o-P\n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseCmdLine(): END LOG FILE\n')
            checkPause()
            sys.exit(0)
        elif opt in ("-v", "--version"):
            if is_dump1090sql:
                print(prog_name , '- version', version)
            else:
                print(prog_name_report , '- version', version)
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(log_buffer)
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseCmdLine(): show version\n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseCmdLine(): sys.exit(0): Normal E-o-P\n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseCmdLine(): END LOG FILE\n')
            checkPause()
            sys.exit(0)
    sql_file = False
    for opt, arg in opts:
    # check all options
        if opt in ("--net-http-ip", '--ip'):
            server_IP = arg
        elif opt in ("--net-http-port", '--port'):
            server_PORT = arg
        elif opt in ("-e","--exclude"):
            exclude = True
        elif opt in ("-s","--sql"):
            write2sql = True
        elif opt in ("--sqlfile"):
            DB_file = arg
            sql_file = True
        elif opt in ("-t", "--txt"):
            write2file = True
        elif opt in ("--txtfile"):
            output_file = arg
        elif opt in ("--lat", "--latitude"):
            latitude = arg
        elif opt in ("--lon", "--longitude"):
            longitude = arg
        elif opt in ("--alt","--altitude"):
            altitude = arg
        elif opt in ("--sta","--station"):
            station_name = arg
        elif opt in ("--metric"):
            metric = True
            change_to_metric = True
        elif opt in ("-q", "--quiet"):
            quiet = True
            pause = False
            interactive = False
        elif opt in ("-i", "--interactive"):
            if not quiet:
                interactive = True
        elif opt in ("--delay"):
            cnx_try = True
            cnx_delay = arg
        elif opt in ("--maxalt"):
            max_altitude = arg
        elif opt in ("--maxdist"):
            max_distance = arg
        elif opt in ("--opencsv"):
            open_csv = True
        elif opt in ("--email"):
            if not zip_backup:
                msg_buffer = msg_buffer + "Command line: WARNING: No --zip option, send email option will be ignored\n"
            else:
                email_backup = True
        elif opt in ("--remove"):
            if not zip_backup:
                msg_buffer = msg_buffer + "Command line: WARNING: No --zip option, remove option will be ignored\n"
            else:
                remove = True
        elif opt in ("--extdb"):
            extDB_filename = arg
            extDB = True
    for opt, arg in opts:
        if opt in ("--report"):
            if not sql_file:
                msg_buffer = msg_buffer + "Command line: WARNING: No --sqlfile option, --report option will be ignored\n"
            else:
                report = True
                new_report = True
        if opt in ("--oldreport"):
            if not sql_file:
                msg_buffer = msg_buffer + "Command line: WARNING: No --sqlfile option, --oldreport option will be ignored\n"
            else:
                report = True
                onesql_report = True
                old_report = True
        if opt in ("--ifile"):
            if not sql_file:
                msg_buffer = msg_buffer + "Command line: WARNING: No --sqlfile option, --ifile option will be ignored\n"
            else:
                input_file = arg
                load2sql = True
        if opt in ("--ireport"):
            if not extDB:
                msg_buffer = msg_buffer + "Command line: WARNING: No --extDB option, --ireport option will be ignored\n"
                ireport = False
                ireport_filename = 'N/A'
            else:
                ireport = True
                ireport_filename = arg
    if old_report and new_report:
        old_report = False
        msg_buffer = msg_buffer + "Command line: WARNING: --report and --oldreport options sets, --report option will be ignored\n"
    if report or load2sql:
        option_msg = ''
        if report:
            if onesql_report:
                option_msg = "--oldreport "
            else:
                option_msg = "--report "
        if report and load2sql:
            option_msg = option_msg + "& "
        if load2sql:
            option_msg = option_msg + "--ifile "
        msg_buffer = msg_buffer + "Command line: WARNING: " + option_msg + "option(s) set(s), other option will be ignored\n"
        write2sql = False
        write2file = False
        interactive = False
    if report and ireport:
        if onesql_report:
            msg_buffer = msg_buffer + "Command line: WARNING: --oldReport option set, --ireport option will be ignored\n"
        else:
            msg_buffer = msg_buffer + "Command line: WARNING: --report option set, --ireport option will be ignored\n"
        ireport = False
        ireport_filename = 'N/A'
    if extDB:
        if not report and not ireport:
            msg_buffer = msg_buffer + "Command line: WARNING: No --report or --ireport option, --extDB option will be ignored\n"
            extDB = False
            extDB_filename = 'N/A'
    if not report and not ireport:
        open_csv = False
    net_HTTP = 'http://' + server_IP + ':' + server_PORT + '/data.json'
    log_buffer = log_buffer + msg_buffer + '\n'  
    if ((DB_file[0] == '.') or (output_file[0] == '.') or (input_file[0] == '.') or (zip_filename[0] == '.') or (extDB_filename[0] == '.') or (ireport_filename[0] == '.')):
        print('')
        print(exe_name + ": FATAL ERROR: filename in command line should use full pathname")
        print(exe_name + ":              filename can't start with '.'")
        checkPause()
        sys.exit(2)
        
def showOptions():
# Show all options on screen 
    global prog_name
    global prog_name_report
    global exe_name
    global os_win
    global net_HTTP
    global DB_file
    global exclude
    global max_altitude
    global max_distance
    global write2sql
    global output_file
    global write2file
    global input_file
    global load2sql
    global quiet
    global pause
    global interactive
    global latitude
    global longitude
    global altitude
    global station_name
    global metric
    global change_to_metric
    global log
    global log_file
    global local_path
    global cnx_try
    global cnx_delay
    global zip_backup
    global zip_filename
    global email_backup
    global email_toaddr
    global remove
    global month
    global report
    global open_csv
    global extDB
    global extDB_filename
    global ireport
    global ireport_filename
    global is_dump1090sql
    global onesql_report
    cnx_delay_msg = str(cnx_delay)
    if cnx_delay == 0:
        cnx_delay_msg = 'N/A'
    output_file_msg = output_file
    if not write2file:
        output_file_msg = 'N/A'
    zip_filename_msg = 'N/A'
    if zip_backup:
        zip_filename_msg = zip_filename + '_' + month + '.zip'
    email_toaddr_msg = email_toaddr
    email_toaddr_msg = email_toaddr_msg.replace(",","\n                                ")
    if not email_backup:
        email_toaddr_msg = 'N/A'
    if max_altitude == 0:
        max_altitude_msg = 'N/A'
    else:
        if not metric:
            max_altitude_msg = str(max_altitude) + 'ft/'
            max_altitude_meters = int(max_altitude * 0.3048)
            max_altitude_msg = max_altitude_msg + str(max_altitude_meters) + 'm'
        else:   
            max_altitude_meters = max_altitude
            max_altitude = int(max_altitude / 0.3048)
            max_altitude_msg = str(max_altitude) + 'ft/' + str(max_altitude_meters) + 'm'
    if max_distance == 0:
        max_distance_msg = 'N/A'
    else:
        if not metric:
            max_distance_msg = str(max_distance) + 'ft/'
            max_distance_meters = int(max_distance * 0.3048)
            max_distance_msg = max_distance_msg + str(max_distance_meters) + 'm'
        else:   
            max_distance_meters = max_distance
            max_distance = int(max_distance / 0.3048)
            max_distance_msg = str(max_distance) + 'ft/' + str(max_distance_meters) + 'm'
    metric_msg = ''
    if ((max_distance > 0) or (max_altitude > 0)) and change_to_metric:
        metric_msg = 'WARNING: unit was changed in command line'
    if not write2sql and not load2sql and not report:
        DB_file = 'N/A'
    open_csv_msg = open_csv  
    if not report:
        open_csv_msg = 'N/A'
    alt_msg = 'N/A'
    if (altitude != 'N/A'):
        falt=int(float(altitude))
        falt_ft = int(falt / 0.3048)
        alt_msg = str(falt_ft) + 'ft/' + str(falt) + 'm'
    DB_file = os.path.normpath(DB_file)
    output_file_msg = os.path.normpath(output_file_msg)
    input_file = os.path.normpath(input_file)
    extDB_filename = os.path.normpath(extDB_filename)
    ireport_filename = os.path.normpath(ireport_filename)
    zip_filename_msg = os.path.normpath(zip_filename_msg)
    if is_dump1090sql:
        print('')
        print('OPTIONS:')
        print('This script filename:            ', exe_name)
        print('Windows OS:                      ', os_win)
        print('Local folder:                    ', local_path)
        print('Save file to folder:             ', os.getcwd())
        print('')
        print('DUMP1090 json source:            ', net_HTTP)
        print('Try to reconnect on lost cnx:    ', cnx_try)
        print('Delay before reconnecting (s):   ', cnx_delay_msg) 
        print('Write to database:               ', write2sql)
        print('Database filename:               ', DB_file)
        print('Write to file:                   ', write2file)
        print('Output filename:                 ', output_file_msg)
        print('Load 01DB file into database:    ', load2sql)
        print('01DB xls filename:               ', input_file)
        print('Make report:                     ', report)
        print('Find flight in external database:', extDB)
        print('External database filename:      ', extDB_filename)
        print('Add flights infos to old report: ', ireport)
        print('Old report filename:             ', ireport_filename)
        print('Open report in Excel:            ', open_csv_msg)
        print('Make compressed backup:          ', zip_backup)
        print('Compressed backup filename:      ', zip_filename_msg)
        print('Remove files after compression:  ', remove)
        print('Send email with backup file:     ', email_backup)
        print('Email recipient(s):              ', email_toaddr_msg)
        print('Exclude inconsistent data:       ', exclude)
        if change_to_metric:
            print(metric_msg)
        print('Maximum altitude flight:         ', max_altitude_msg)
        print('Maximum distance from station:   ', max_distance_msg)
        print('Latitude (local station):        ', latitude)
        print('Longitude (local station):       ', longitude)
        print('Altitude (local station):        ', alt_msg)
        print('Station name (local station):    ', station_name)
        print('Metric units:                    ', metric)
        print('Quiet mode:                      ', quiet)
        print('Enable pauses:                   ', pause)
        print('Interactive mode:                ', interactive)
        print('Log file:                        ', log)
        print('')
        print('Starting process, please wait...')
    else:
        print('')
        print('OPTIONS:')
        print('This script filename:            ', exe_name)
        print('Windows OS:                      ', os_win)
        print('Local folder:                    ', local_path)
        print('Save file to folder:             ', os.getcwd())
        print('')
        print('Database filename:               ', DB_file)
        print('Load 01DB file into database:    ', load2sql)
        print('01DB xls filename:               ', input_file)
        print('Make report:                     ', report)
        print('Find flight in external database:', extDB)
        print('External database filename:      ', extDB_filename)
        print('Add flights infos to old report: ', ireport)
        print('Old report filename:             ', ireport_filename)
        print('Open report in Excel:            ', open_csv_msg)
        if change_to_metric:
            print(metric_msg)
        print('Maximum altitude flight:         ', max_altitude_msg)
        print('Maximum distance from station:   ', max_distance_msg)
        print('Metric units:                    ', metric)
        print('Quiet mode:                      ', quiet)
        print('Enable pauses:                   ', pause)
        print('Log file:                        ', log)
        print('')
        print('Starting process, please wait...')
    #time.sleep(3)
    if log:
        with open(log_file,'a') as target_file:
            target_file.write('Windows OS:' + str(os_win) + '\n')
            target_file.write('Local folder:' + local_path + '\n')
            target_file.write('Save file to folder:' + os.getcwd() + '\n')
            target_file.write('DUMP1090 json source:' + net_HTTP + '\n')
            target_file.write('Try to reconnect:' + str(cnx_try) + '\n')
            target_file.write('Delay before reconnect:' + cnx_delay_msg + '\n')
            target_file.write('Database filename:' + DB_file + '\n')
            target_file.write('Write to database:' + str(write2sql) + '\n')
            target_file.write('Output filename:' + output_file_msg + '\n')
            target_file.write('Write to file:' + str(write2file) + '\n')
            target_file.write('Load 01DB file into database:' + str(load2sql) + '\n')
            target_file.write('Make report:' + str(report) + '\n')
            target_file.write('One SQL request:' + str(onesql_report) + '\n')
            target_file.write('Use external ddb:' + str(extDB) + '\n')
            target_file.write('External ddb filename:' + extDB_filename + '\n')
            target_file.write('Use old report:' + str(ireport) + '\n')
            target_file.write('Old report filename:' + ireport_filename + '\n')
            target_file.write('Open report in Excel:' + str(open_csv_msg) + '\n')
            target_file.write('01DB xls filename:' + input_file + '\n')
            target_file.write('Make report:' + str(report) + '\n')
            target_file.write('Backup:' + str(zip_backup) + '\n')
            target_file.write('Backup filename:' + zip_filename_msg + '\n')
            target_file.write('Remove file:' + str(remove) + '\n')
            target_file.write('Send email: ' + str(email_backup) + '\n')
            target_file.write('Email recipient(s):' + email_toaddr + '\n')
            target_file.write('Exclude inconsistent data:' + str(exclude) + '\n')
            if change_to_metric:
                target_file.write(metric_msg + '\n')
            target_file.write('Maximum altitude:' + max_altitude_msg + '\n')
            target_file.write('Maximum distance:' + max_distance_msg + '\n')
            target_file.write('Latitude:' + latitude + '\n')
            target_file.write('Longitude:' + longitude + '\n')
            target_file.write('Altitude:' + alt_msg + '\n')
            target_file.write('Station name:' + station_name + '\n')
            target_file.write('Metric units:' + str(metric) + '\n')
            target_file.write('Quiet mode:' + str(quiet) + '\n')
            target_file.write('Enable pauses:' + str(pause) + '\n')
            target_file.write('Interactive mode:' + str(interactive) + '\n')   
            target_file.write('Log file:' + str(log) + '\n')            
    checkPause()
    
def readJSON():
# connect to DUMP1090 server
    global exe_name
    global stream
    global net_HTTP
    global opener
    global interactive
    global log
    global log_file
    global quiet
    global cnx_delay
    global cnx_try
    global try_again
    global first_time
    global write2sql
    global load2sql
    if log:
        with open(log_file,'a') as target_file:
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readJSON(): opening stream ' + net_HTTP + '\n')
    try:
        # read JSON stream
        stream = opener.open(net_HTTP)
        if try_again:
            if not quiet:
                print("Server '" + net_HTTP + "' is now responding")
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readJSON(): Server ' + net_HTTP + ' is now responding\n')
        try_again = False
        first_time = True
    except (OSError, IOError) as e:
        if not cnx_try:
            # Error during connexion and not going to reconnect
            print('')
            print(exe_name + ": FATAL ERROR: Trying to connect to '" + net_HTTP + "'")
            print('')
            print(str(e))
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readJSON(): FATAL ERROR: Trying to connect to ' + net_HTTP + '\n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readJSON(): OSError, IOError: ' + str(e) + '\n')
            # If database was previously open, close it
            # load2sql?
            if write2sql:
                conn.close()
                print('')
                print("Closing database '" + DB_file + "'")
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readJSON(): Closing database ' + DB_file + '\n')
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readJSON(): sys.exit(2): E-o-P \n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readJSON(): END LOG FILE\n')
            checkPause()
            sys.exit(2)
        else:
            # Trying to reconnect after x seconds
            if first_time:
                first_time = False
                if not quiet:
                    print(exe_name + ": WARNING: Server '" + net_HTTP + "' not responding")
                    print('Trying to reconnect every', str(cnx_delay), 'second(s)')
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readJSON(): WARNING: Server ' + net_HTTP + ' not responding\n')
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readJSON(): trying to reconnect every ' + str(cnx_delay) + 'second(s)\n')
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readJSON(): OSError, IOError: ' + str(e) + '\n')
            waittime = float(cnx_delay)
            try_again = True
            time.sleep(waittime)       
            pass
            
def pingServer():
    global server_IP
    global os_win
    global exe_name
    global quiet
    global cnx_delay
    global log
    global log_file
    try_again = False
    if os_win:
        cmdline = 'ping -n 1 ' + server_IP + ' | find "TTL=" >NUL'
    else:
        cmdline = 'ping -q -c 1 ' + server_IP + ' > /dev/null'
    if not quiet:
        print('')
        print("Contacting server '" + server_IP + "'")
    if log:
        with open(log_file,'a') as target_file:
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'pingServer(): contacting server ' + server_IP + '\n')
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'pingServer(): ' + cmdline + '\n')
    if os.system(cmdline) != 0:
        if not cnx_try:
            if not quiet:
                print('')
                print(exe_name + ": FATAL ERROR: Server '" + server_IP + "' not responding")
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'pingServer(): FATAL ERROR: Trying to connect to server ' + server_IP + '\n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'pingServer(): ' + cmdline + ' fail\n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'pingServer(): sys.exit(2): E-o-P \n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'pingServer(): END LOG FILE\n')        
            checkPause()
            sys.exit(2)       
        else:
            try_again = True
            if not quiet:
                print(exe_name + ': WARNING: Server not responding')
                print('Trying to reconnect every', str(cnx_delay), 'second(s)')
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'pingServer(): WARNING: server not responding (ping)\n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'pingServer(): trying to reconnect every ' + str(cnx_delay) + 'second(s)\n')
    while try_again:
        waittime = float(cnx_delay)
        time.sleep(waittime)
        if os.system(cmdline) == 0:
            try_again = False
            if not quiet:
                print(exe_name + ': Server is now ok')
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'pingServer(): server is now ok\n')
    
def parseJSON():
# Read JSON streaming and parse it
    global stream
    global write2file
    global output_file
    global interactive
    global cursor
    global first_JSON
    global idx_infos
    global metric
    global log
    global log_file
    global today
    global max_altitude
    # Read all JSON data (1 JSON data = x JSON informations)
    line = stream.read()
    # Clean data
    line = line.decode('utf-8', 'ignore')
    if log:
        line_log = line
        replace_list = [',', ']\n', '{', '}', '"', '[\n']
        for i in replace_list:
            line_log = line_log.replace(i, '')
        with open(log_file,'a') as target_file:
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseJSON(): JSON datas\n' + line_log)
    # Format ALL data => parsedLine
    parsedLine = json.loads(line)
    # if --interactive, clear screen (Windows or other OS version) and show title bar
    if interactive:
        line = '--------+--------+-----+---+---+----------+-----------+----------+-------- '
        if exclude:
            line = line + 'F'
        else:
            line = line + ' '
        
        if metric:
            line = line + 'M'
        else:
            line = line + ' '
        if write2sql:
            line = line + 'S'
        else:
            line = line + ' '
        if write2file:
            line = line + 'T'
        else:
            line = line + ' '
        if os_win:
            os.system('cls')
        else:
            os.system('clear')
        print('  ICAO  | FLIGHT | ALT |SPD|Hdg|   LAT    |    LON    |   DATE   |  TIME  ')
        print(line)
    for result in parsedLine:
    # Read and format all previous formatted JSON datas
        parsedRow = json.loads(json.dumps(result))
        # Write everythings or filtered datas?
        # Do not save if missing altitude, latitude, longitude, position is not valid
        # Do not save if we received more 60 times the same information
        if (((not exclude) or (exclude and parsedRow['altitude'] != 0 and parsedRow['lat'] != 0.0 and parsedRow['lon'] != 0.0 and parsedRow['validposition'] != 0 and parsedRow['seen'] < 600)) and not ((max_altitude > 0) and (parsedRow['altitude'] > max_altitude))):
            first_JSON = False
            # YYYY/MM/DD;HH:MM:SS (save date/time in this special format for compatibility with old Excel macro)
            nday = time.strftime("%Y/%m/%d", time.localtime())
            # format date/time for sql comatibility
            fday = time.strftime("%Y-%m-%d", time.localtime())
            ftime = time.strftime("%H:%M:%S", time.localtime())
            fdaytime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            # New formatted line to write in txt file
            line2file = parsedRow['hex'].ljust(8).upper() + ';' + parsedRow['flight'].ljust(8).upper() + ';' + str(parsedRow['altitude']).rjust(5) + ';' + str(parsedRow['speed']).rjust(3) + ';' + str(parsedRow['track']).rjust(3) + ';' + str(parsedRow['lat']).rjust(10) + ';' + str(parsedRow['lon']).rjust(11) + ';' + nday + ";" + ftime
            # If --sql option, do an INSERT into database with formatted datas
            ntoday = time.strftime("%Y%m%d", time.localtime())
            # Date changed?
            new_date = False
            if ntoday != today:
                new_date = True
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseJSON(): FILTERED (NEW DATE): ICAO:' + parsedRow['hex'] + ' valid pos:' + str(parsedRow['validposition']) +' seen:' + str(parsedRow['seen']) + ' lat:' + str(parsedRow['lat']) + ' lon:' + str(parsedRow['lon']) + ' alt:' + str(parsedRow['altitude']) + '\n')
            if write2sql and not new_date:
                try:
                    cursor.execute("""INSERT INTO aircrafts (icao, flight, altitude, speed, heading, latitude, longitude, date, time, datetime, idx_infos) values(?,?,?,?,?,?,?,?,?,?, ?)""", (parsedRow['hex'].upper(), parsedRow['flight'].upper(), parsedRow['altitude'], parsedRow['speed'], parsedRow['track'], parsedRow['lat'], parsedRow['lon'], fday, ftime, fdaytime, idx_infos))
                except sqlite3.Error as e:
                    # Problem during insertion into database, quit.
                    print('')
                    print(exe_name + ": FATAL ERROR: Trying to insert datas into database '" + DB_file + "'")
                    print('')
                    print(str(e))
                    if log:
                        with open(log_file,'a') as target_file:
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseJSON(): FATAL ERROR: Trying to insert datas into database ' + DB_file + ', table aircrats\n')
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseJSON(): sqlite3.Error: ' + str(e) + '\n')
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseJSON(): sys.exit(2): E-o-P\n')
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseJSON(): END LOG FILE\n')
                    checkPause()
                    sys.exit(2)
            # If --txt option, WRITE to file with formatted datas
            if write2file and not new_date:
                with open(output_file,'a') as target_file:
                    target_file.write(line2file + '\n')
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseJSON(): writing to file ' + output_file + '\n' + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseJSON(): ' + line2file + '\n')
            # if --interactive, show formatted datas on screen   
            if interactive:
                seen = ''
                if parsedRow['seen'] > 59:
                    seen = '+'
                line2screen = parsedRow['hex'].ljust(8).upper() + ' ' + parsedRow['flight'].ljust(8).upper() + ' ' + str(parsedRow['altitude']).rjust(5) + ' ' + str(parsedRow['speed']).rjust(3) + ' ' + str(parsedRow['track']).rjust(3) + ' ' + str(parsedRow['lat']).rjust(10) + ' ' + str(parsedRow['lon']).rjust(11) + ' ' + nday + ' ' + ftime + ' ' + seen
                if metric:
                    # format speed (metric)
                    line2screen = parsedRow['hex'].ljust(8).upper() + ' ' + parsedRow['flight'].ljust(8).upper() + ' ' + str(int(round(parsedRow['altitude'] * 0.3048))).rjust(5) + ' ' + str(int(round(parsedRow['speed'] * 1.852))).rjust(3) + ' ' + str(parsedRow['track']).rjust(3) + ' ' + str(parsedRow['lat']).rjust(10) + ' ' + str(parsedRow['lon']).rjust(11) + ' ' + nday + ' ' + ftime + ' ' + seen
                # Show datas on screen
                print(line2screen)
        else:
            # Filtered datas
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseJSON(): FILTERED (NEW DATE): ICAO:' + parsedRow['hex'] + ' valid pos:' + str(parsedRow['validposition']) +' seen:' + str(parsedRow['seen']) + ' lat:' + str(parsedRow['lat']) + ' lon:' + str(parsedRow['lon']) + ' alt:' + str(parsedRow['altitude']) + '\n')
    # Do an commint on database after inserted alls JSON datas
    if write2sql:
        conn.commit()

def testDate():
# test if the date change
    global today
    global month
    global output_file
    global conn    
    global write2sql
    global write2file
    global DB_file
    global log_file
    global log
    global quiet
    global version
    global log_buffer
    global year
    global zip_month
    global zip_year
    global remove
    global email_backup
    global zip_ok
    # Zip all txt and db file, then send email
    zip_backup = False
    # Get day
    ntoday = time.strftime("%Y%m%d", time.localtime())
    # Date changed?
    if ntoday != today:
        # New day
        if not quiet:
            print('')
            print('New date detected')
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'testDate(): new date detected\n') 
        # New output_file name if txtfile option not selected
        if write2file:
            output_datetime_file = today + '_Flights.txt'
            if output_file == output_datetime_file:
                output_file = ntoday + '_Flights.txt'
                if not quiet:
                    print("New txt output filename: '" + output_file + "'")
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'testDate(): new output_file name: ' + output_file + '\n')
        nmonth = time.strftime("%Y%m", time.localtime())
        SQL_datetime_file = month + '_Flights.db'
        # if month changed and if --sqlfile option not selected, create new database
        if (nmonth != month):
            # New month, zip and send email
            zip_backup = True
            zip_month = month
            zip_year = year
            if not quiet:
                print('New month detected')
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'testDate(): new month detected.\n')
            # if previous database was open, close it
            if write2sql:
                conn.close()
                if not quiet:
                    print("Closing database '" + DB_file + "'")
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'testDate(): closing database ' + DB_file + '\n')
                # Change database name if option sqlfile not selected
                if DB_file == SQL_datetime_file:
                    DB_file = nmonth + '_Flights.db'
                if not quiet:
                    print("New database name in new directory: '" + DB_file + "'")
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'testDate(): new database name in new directory: ' + DB_file + '\n')
            month = nmonth
            year = time.strftime("%Y", time.localtime())
            # Close actual log file
            if not quiet:
                print('Creating new directory structure with new date')
            now = time.strftime("%Y%m%d%H%M%S", time.localtime())
            new_log_file = now + '_logs.txt'
            old_log_file = log_file
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'testDate(): creating new directory structure with new date\n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'testDate(): closing log file ' + old_log_file + ', opening ' + new_log_file + '\n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'testDate(): END LOG FILE\n')
            # Go back to intitial path
            os.chdir("..\..")
            createFolders()
            # Open new log file
            log_file = new_log_file
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'testDate(): START LOG FILE\n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'testDate(): script ' + exe_name + ' v' + str(version) + '\n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'testDate(): python ' + str(sys.version_info) + '\n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'testDate(): log file ' + old_log_file + ' was closed, ' + new_log_file + ' is now open\n')
                    if log_buffer != '': 
                        target_file.write(log_buffer)
                log_buffer = ''
            # Create database (if needed)
            if write2sql:
               # Open (or create new) database
                openBase()
                # Create all tables (if needed)
                createTables()
                # Create index on aircrafts table (if needed)
                createIndex()
                # Save std informations (date/time/GPS coord...)
                insertInfos()
                # Get new record index (from std infos)
                selectMaxInfos()
        else:
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'testDate(): END LOG FILE\n')
            now = time.strftime("%Y%m%d%H%M%S", time.localtime())
            log_file = now + '_logs.txt'
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'testDate(): START LOG FILE\n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'testDate(): script ' + exe_name + ' v' + str(version) + '\n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'testDate(): python ' + str(sys.version_info) + '\n')
        today = ntoday
    # New month = backup
    if zip_backup:
        # Compress all files (txt and db)
        zipAllFiles()
        
def zipAllFiles():
    global zip_month
    global zip_year
    global local_path
    global zip_filename
    global quiet
    global log
    global log_file
    global remove
    global email_backup
    global email_subject_name
    #Define folder to compress
    zippath = local_path + '/' + zip_year + '/' + zip_month
    zippath = os.path.normpath(zippath)
    #Define zip filename
    zip_filename_msg = local_path + '/' + zip_year + '/' + zip_filename + '_' + zip_month
    zip_filename_msg = os.path.normpath(zip_filename_msg)
    try:
        #Compress all files
        result = shutil.make_archive(zip_filename_msg, 'zip', zippath)
        zip_ok = True
        if not quiet:
            print("New backup file '" + result + "'")
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'zipAllFiles(): backup file:' + result + '\n')
        # if --remove, delete all files and remove folder
        if remove:
            try:
                shutil.rmtree(zippath)  
                if not quiet:
                    print("Deleting all files and database in '" + zippath + "'")
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'zipAllFiles(): deleting all files/database in ' + zippath + '\n')
            except:
                if not quiet:
                    print('')
                    print(exe_name + ": WARNING: Can't delete '", zippath, "' folder")
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "zipAllFiles(): WARNING: Can't delete " + zippath + " folder\n")
        # send email
        if email_backup:
            sendEmail()
    except:
        zip_ok = False
        if not quiet:
            print('')
            print(exe_name + ': WARNING: Problem during compression, zip file was not created')
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'zipAllFiles(): WARNING: Problem during compression, zip file was not created\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "zipAllFiles(): zippath:" + zippath + " zip_filename_msg:" + zip_filename_msg + " command: shutil.make_archive(" + zip_filename_msg + ", 'zip', " + zippath + ")\n")
        checkPause()
        pass
    checkPause()
    
def sendEmail():
    global quiet
    global log
    global log_file
    global email_backup
    global email_toaddr
    global email_fromaddr
    global email_password
    global email_subject
    global zip_month
    global zip_filename
    global zip_year
    global email_body
    global local_path
    global email_server
    global email_port
    # can send to multiples recipients using "email1@domain,email2@domain..."
    if email_backup:
        toaddr = email_toaddr.split(',')
        msg = MIMEMultipart()
        #Email header
        msg['From'] = email_fromaddr
        msg['To'] = email_toaddr
        msg['Subject'] = email_subject + zip_filename + '_' + zip_month
        body = email_body
        #Email body
        msg.attach(MIMEText(body, 'plain'))
        #Define attachment backup file 
        filename = zip_filename + '_' + zip_month + '.zip' 
        filename_path = local_path + '/' + zip_year + '/' + zip_filename + '_' + zip_month + '.zip'
        filename_path = os.path.normpath(filename_path)
        attachment = open(filename_path, "rb")
        #Convert attachment to base64 file type
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        #Join attachment 
        msg.attach(part)
        #Send message 
        server = smtplib.SMTP(email_server, email_port)
        server.ehlo()
        server.starttls()
        #Log to server
        try:
            server.login(email_fromaddr, email_password)
            text = msg.as_string()
            try:
                #send email
                server.sendmail(email_fromaddr, toaddr, text)
                if not quiet:
                    print("Sending backup email from '" + email_fromaddr + "' to '" + email_toaddr + "'")
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'sendEmail(): Sending email from ' + email_fromaddr + ' to ' + email_toaddr + '\n')
            except Exception as e:
                if not quiet:
                    print('')
                    print(exe_name + ": WARNING: Can't send email")
                    print(str(e))
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "sendEmail(): WARNING: Can't email from " + email_fromaddr + " to " + email_toaddr + "\n")
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + str(e) + "\n")
        except Exception as e:
            if not quiet:
                print('')
                print(exe_name + ": WARNING: Can't log to server '" + email_server + "'")
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "sendEmail(): WARNING: Can't log to server " + email_server + " port " + str(email_port) + ". Login:" + email_fromaddr + "\n")
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + str(e) + "\n")
        # Close connexion with email server
        server.quit()
        if not quiet:
            print('Going back to work')
        
def getNewJSON():
    global try_again
    global first_time
# Read JSON streaming and parse datas
    while True:
        first_time = True
        readJSON()
        # On connexion error, try again if option --delay selected
        while try_again:
            readJSON()
        parseJSON()
        if write2file or write2sql:
            testDate()
        time.sleep(1)
        
def checkPause():
# If --nopause is NOT set, make a pause after all (main) process, warning or errors
    global pause
    global os_win
    global quiet
    if pause and not quiet:
        print('')
        # Depending on witch OS, make a pause
        if os_win:
            os.system('pause')
        else:
            os.system('read -s -n 1 -p "Press any key to continue . . ."')

def createFolders():
# if not existes, create new folders YEAR/YEARMONTH
    global year
    global month
    global log
    global log_buffer
    global exe_name
    global quiet
    if not os.path.exists(year):
        try:
            os.makedirs(year)
            if not quiet:
                print("Creating folder '" + str(year) + "'")
            log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createFolders(): creating folder ' + str(year) + '\n'
        except OSError  as e:
            if not quiet:
                print('')
                print(exe_name + ": WARNING: Cannot create folder '" + str(year) + "'")
                print(str(e))
            log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createFolders(): WARNING: cannot create folder ' + str(year) + '\n'
            log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createFolders(): OSError: ' + str(e) + '\n'
            pass
    try:
        os.chdir(year)
    except OSError as e: 
        if not quiet:
            print('')
            print(exe_name + ": WARNING: Cannot change path to '" + str(year) + "'")
            print(str(e))
        log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createFolders(): WARNING: cannot change path to ' + str(year) + '\n'
        log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createFolders(): OSError: ' + str(e) + '\n'
        pass
    if not os.path.exists(month):
        try:
            os.makedirs(month)
            if not quiet:
                print("Creating folder '" + str(month) + "'")
            log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createFolders(): creating folder ' + str(month) + '\n'
        except OSError  as e:
            if not quiet:
                print('')
                print(exe_name + ": WARNING: Cannot create folder '" + str(month) + "'")
                print(str(e))
            log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createFolders(): WARNING: cannot create folder ' + str(month) + '\n'
            log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createFolders(): OSError: ' + str(e) + '\n'
            pass
    try:
        os.chdir(month)
        if not quiet:
            print('Path changed with success')
    except OSError as e: 
        if not quiet:
            print('')
            print(exe_name + ": WARNING: Cannot change path to '" + str(month) + "'")
            print(str(e))
        log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createFolders(): WARNING: cannot change path to ' + str(month) + '\n'
        log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createFolders(): OSError: ' + str(e) + '\n'
        pass
        
def changeFolders():
# if not existes, create new folders YEAR/YEARMONTH
    global year
    global month
    global log_buffer
    global exe_name
    global quiet
    try:
        os.chdir(year)
    except OSError as e: 
        if not quiet:
            print('')
            print(exe_name + ": WARNING: Cannot change path to '" + str(year) + "'")
            print(str(e))
        log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'changeFolders(): WARNING: cannot change path to ' + str(year) + '\n'
        log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'changeFolders(): OSError: ' + str(e) + '\n'
        pass
    try:
        os.chdir(month)
        if not quiet:
            print('Path changed with success')
    except OSError as e: 
        if not quiet:
            print('')
            print(exe_name + ": WARNING: Cannot change path to '" + str(month) + "'")
            print(str(e))
        log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'changeFolders(): WARNING: cannot change path to ' + str(month) + '\n'
        log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'changeFolders(): OSError: ' + str(e) + '\n'
        pass
    
# Open database (or new database)
def openBase():
    global conn
    global log_file
    global DB_file
    global exe_name
    global log
    try:
        conn = sqlite3.connect(DB_file)
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'openBase(): connexion to database ' + DB_file + ' success\n')
    except sqlite3.Error as e:
        # Problem during connexion with database, quit.
        print('')
        print(exe_name + ": FATAL ERROR: Trying to open database '" + DB_file + "'")
        print('')
        print(str(e))
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'openBase(): FATAL ERROR: Trying to open database ' + DB_file + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'openBase(): sqlite3.Error: ' + str(e) + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'openBase(): sys.exit(2): E-o-P\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'openBase(): END LOG FILE\n')
        checkPause()
        sys.exit(2)

# Open external database (flight)
def openExtDB():
    global conn
    global log_file
    global extDB_filename
    global exe_name
    global log
    try:
        conn = sqlite3.connect(extDB_filename)
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'openExtDB(): connexion to database ' + extDB_filename + ' success\n')
    except sqlite3.Error as e:
        # Problem during connexion with database, quit.
        print('')
        print(exe_name + ": FATAL ERROR: Trying to open database '" + extDB_filename + "'")
        print('')
        print(str(e))
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'openExtDB(): FATAL ERROR: Trying to open external database ' + extDB_filename + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'openExtDB(): sqlite3.Error: ' + str(e) + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'openExtDB(): sys.exit(2): E-o-P\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'openExtDB(): END LOG FILE\n')
        checkPause()
        sys.exit(2)
        
# Create all tables (if needed)
def createTables():
    global conn
    global cursor
    global log
    global log_file
    global DB_file
    global write2sql
    global load2sql
    global exe_name
    global quiet
    try:
        # try to create table if it's a new database
        cursor = conn.cursor()
        # infos table (GPS coord, date/time)
        cursor.execute("""CREATE TABLE IF NOT EXISTS infos (idx_infos INTEGER PRIMARY KEY AUTOINCREMENT, date DATE, time TIME, latitude REAL, longitude REAL, altitude INTEGER, station_name VARCHAR);""")
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createTables(): create tables aircrafts and infos into database ' + DB_file + '\n')
        if write2sql:
            # aircrafts table: all infos abouts planes and flights
            cursor.execute("""CREATE TABLE IF NOT EXISTS aircrafts (idx_aircrafts INTEGER PRIMARY KEY AUTOINCREMENT, icao STRING (6), flight VARCHAR (8), altitude INT, speed INT, heading INTEGER, latitude REAL, longitude REAL, date DATE, time TIME, datetime DATETIME, idx_infos INTEGER);""") 
        if load2sql:
            # infos_01DB table: generics infos about sound station (DUO / CUBE...)
            cursor.execute("""CREATE TABLE IF NOT EXISTS infos_01DB (idx_infos_01DB INTEGER PRIMARY KEY AUTOINCREMENT, filename VARCHAR, location VARCHAR, latitude REAL, longitude REAL, altitude INTEGER, source VARCHAR, type_of_datas STRING (16), weighting CHAR, unit STRING (4), start DATETIME, stop DATETIME);""")
            # events_01DB table : all infos abouts sound's events from soudn station
            cursor.execute("""CREATE TABLE IF NOT EXISTS events_01DB (idx_events_01DB INTEGER PRIMARY KEY AUTOINCREMENT, datetime DATETIME, duration TIME, leq REAL, lmax REAL, lmax_time DATETIME, sel REAL, idx_infos_01DB INTEGER);""")
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createTables(): create tables infos_01DB and events_01DB into database ' + DB_file + '\n')
        conn.commit()
    except conn.Error as e:
    # Problem during table creation
        print('')
        print(exe_name + ": FATAL ERROR: Trying to create table infos_01DB, events_01DB into '" + DB_file + "'")
        print('')
        print(str(e))
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createTables(): closing database ' + DB_file + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createTables(): FATAL ERROR: Trying to create tables aircrafts, infos, infos_01DB and events_01DB into ' + DB_file + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createTables(): conn.Error: ' + str(e) + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createTables(): sys.exit(2): E-o-P\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createTables(): END LOG FILE\n')
        conn.close()
        if not quiet:
            print('')
            print("Closing database '" + DB_file + "'")
        checkPause()
        sys.exit(2)

# Create index (if needed)
def createIndex():
    global conn
    global cursor
    global log
    global log_file
    global DB_file
    global write2sql
    global load2sql
    global exe_name
    global quiet
    if not quiet:
        print('')
        print("Creating index if needed on database " + DB_file + ". Please wait.")
    try:
        # try to create table if it's a new database
        cursor = conn.cursor()
        # create new index on datetime in aircrafts
        cursor.execute("""CREATE INDEX IF NOT EXISTS timeindex ON aircrafts(datetime);""")
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createTables(): create new index on datetime in aircrafts, database ' + DB_file + '\n')
        conn.commit()
    except conn.Error as e:
    # Problem during table creation
        print('')
        print(exe_name + ": FATAL ERROR: create new index on datetime in aircrafts, database '" + DB_file + "'")
        print('')
        print(str(e))
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createTables(): closing database ' + DB_file + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createTables(): FATAL ERROR: Trying create new index on datetime in aircrafts, database ' + DB_file + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createTables(): conn.Error: ' + str(e) + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createTables(): sys.exit(2): E-o-P\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'createTables(): END LOG FILE\n')
        conn.close()
        if not quiet:
            print('')
            print("Closing database '" + DB_file + "'")
        checkPause()
        sys.exit(2)

        
# Save std informations (date/time/GPS coord...)
def insertInfos():
    global conn
    global cursor
    global log
    global log_file
    global DB_file
    global exe_name
    global quiet
    global latitude
    global longitude
    global altitude
    global station_name
    fday = time.strftime("%Y-%m-%d", time.localtime())
    ftime = time.strftime("%H:%M:%S", time.localtime())
    if latitude == 'N/A':
        latitude = 0.0
    if longitude == 'N/A':
        longitude = 0.0
    if altitude == 'N/A':
        altitude = 0
    if station_name == 'N/A':
        station_name = ''
    try:
        # Write date/time and GPS coord. into database
        sql = "INSERT INTO infos (date, time, latitude, longitude, altitude, station_name) values('" + fday + "','" + ftime + "','" + str(latitude) + "','" + str(longitude) + "','" + str(altitude) + "','" + station_name + "')"
        cursor.execute(sql)
        if not quiet:
            print('')
            print("Saving informations into database '" + DB_file + "'")
        if log:
            with open(log_file,'a') as target_file: 
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'insertInfos(): saving informations into database ' + DB_file + ', table infos\n')
        conn.commit()
    except sqlite3.Error as e:
        # Problem during insertion into database
        print('')
        print(exe_name + ": WARNING: Trying to insert datas into database '" + DB_file + "'")
        print('')
        print(str(e))
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'insertInfos(): WARNING: trying to insert datas into database ' + DB_file + ', table infos\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'insertInfos(): ' + sql + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'insertInfos(): sqlite3.Error: ' + str(e) + '\n')
        pass

# Get new record index (from std infos)
def selectMaxInfos():
    global conn
    global cursor
    global log
    global log_file
    global DB_file
    global exe_name
    global quiet
    global idx_infos
    # Get new idx_infos
    try:
        cursor.execute("""SELECT max(idx_infos) FROM infos""")
        idx_infos = cursor.fetchone()[0]
        if log:
            with open(log_file,'a') as target_file: 
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'selectMaxInfos(): select max(idx_infos) from infos. idx_infos: ' + str(idx_infos) + '\n')
    except sqlite3.Error as e:
        # Problem during reading
        print('')
        print(exe_name + ": WARNING: Trying to select max from database '" + DB_file + "'")
        print('')
        print(str(e))
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'selectMaxInfos(): WARNING: trying to select max from database ' + DB_file + ', table infos\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'selectMaxInfos(): SELECT max(idx_infos) FROM infos\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'selectMaxInfos(): sqlite3.Error: ' + str(e) + '\n')
        checkPause()
        pass
        
# Check if csv file exist
def checkCSV():
    global output_report
    global local_path
    global quiet
    global exe_name
    global log
    global log_file
    global DB_file
    global conn
    input_file = os.path.normpath(output_report)
    if not os.path.isfile(input_file):
        if not quiet:
            print(exe_name + ": FATAL ERROR: CSV file '" + input_file + "' not found")
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "checkCSV(): FATAL ERROR: CSV file " + input_file + " not found\n")
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkCSV(): sys.exit(1): E-o-P\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkCSV(): END LOG FILE\n')
        checkPause()
        sys.exit(1)
    if not quiet:
        print("Found csv report file '" + input_file)
    if log:
        with open(log_file,'a') as target_file:
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkCSV(): found csv report file ' + input_file + '\n')
            
# Check if xls file exist 
def checkXLSfile():
    global input_file
    global local_path
    global quiet
    global exe_name
    global log
    global log_file
    global DB_file
    global conn
    input_file = os.path.normpath(input_file)
    if not os.path.isfile(input_file):
        if not quiet:
            print(exe_name + ": FATAL ERROR: Excel file '" + input_file + "' not found")
            print('')
            print("Closing database '" + DB_file + "'")
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkXLSfile(): FATAL ERROR: Excel file ' + input_file + ' not found\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkXLSfile(): closing database: ' + DB_file + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkXLSfile(): sys.exit(1): E-o-P\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkXLSfile(): END LOG FILE\n')
        conn.close()
        checkPause()
        sys.exit(1)
    if not quiet:
        print("Found xls file '" + input_file)
    if log:
        with open(log_file,'a') as target_file:
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkXLSfile(): found xls file ' + input_file + '\n')
          
def openXLSfile():
    # Open Excel file if exist
    global input_file
    global local_path
    global quiet
    global exe_name
    global log
    global log_file
    global DB_file
    global conn
    global wb
    # Open Excel file
    try:
        wb = xlrd.open_workbook(input_file)
        if log:
            with open(log_file,'a') as target_file: 
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'openXLSfile(): opening Excel file ' + input_file + '\n')
    except Exception as e:
        print('')
        print(exe_name + ": FATAL ERROR: Trying to open Excel file '" + input_file + "'")
        print('')
        print(str(e))
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'openXLSfile(): closing database: ' + DB_file + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'openXLSfile(): FATAL ERROR: trying to open Excel file ' + input_file + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'openXLSfile(): Exception: ' + str(e) + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'openXLSfile(): sys.exit(2): E-o-P\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'openXLSfile(): END LOG FILE\n')
        conn.close()
        if not quiet:
            print('')
            print("Closing database '" + DB_file + "'")
        checkPause()
        sys.exit(2) 

def readSheets():
    # Read all sheets from Excel workbook, save header, parse datas and store into database
    global wb
    global quiet
    global log
    global log_file
    global cursor
    global exe_name
    global DB_file
    global input_file
    # Select all sheets
    worksheets = wb.sheet_names()
    for worksheet_name in worksheets:
        sheet = wb.sheet_by_name(worksheet_name)
        if not quiet:
            print("Opening Excel sheet '" + sheet.name + "'")
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): opening Excel sheet ' + sheet.name + '\n')
        empty_sheet = False
        try:
            filename = sheet.cell_value(0,1)
        except:
            empty_sheet = True
            if not quiet:
                print("Sheet '" + sheet.name + "' is empty")
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): sheet ' + sheet.name + ' is empty\n')
            pass
        if not empty_sheet:
            try:
                latitude_label = sheet.cell_value(2,0)
                latitude_label = latitude_label.upper()
                if latitude_label != 'LATITUDE':
                    empty_sheet = True
                    if not quiet:
                        print("Sheet '" + sheet.name + "' is not valid")
                    if log:
                        with open(log_file,'a') as target_file:
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): sheet ' + sheet.name + ' is not valid\n')
            except:
                empty_sheet = True
                if not quiet:
                    print("Sheet '" + sheet.name + "' is not valid")
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): sheet ' + sheet.name + ' is not valid\n')
                pass
        if not empty_sheet:
            if log:
                with open(log_file,'a') as target_file: 
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): start reading header from Excel file ' + input_file + ', sheet ' + sheet.name + '\n')
            # Get info from header (description) and put everythings in database
            location = sheet.cell_value(1,1)
            latitude = sheet.cell_value(2,1)
            longitude = sheet.cell_value(3,1)
            altitude = sheet.cell_value(4,1)
            source = sheet.cell_value(5,1)
            type_of_datas = sheet.cell_value(6,1)
            weighting = sheet.cell_value(7,1)
            unit = sheet.cell_value(8,1)
            start = sheet.cell_value(9,1)
            stop = sheet.cell_value(10,1)
            # New entry?
            idx_infos_01DB = -1
            result = 0
            try:
                # Look for old record in database
                sql = "SELECT count(*) FROM infos_01DB WHERE (filename = '" + filename + "' AND location = '" + location + "' AND source = '" + source + "' AND type_of_datas = '" + type_of_datas + "' AND weighting = '" + weighting + "' AND unit = '" + unit + "' AND start = DATETIME('1899-12-30', '+" + str(start) + " days') AND stop = DATETIME('1899-12-30', '+" + str(stop) + " days'))"
                cursor.execute(sql)
                result = cursor.fetchone()[0]
                if log:
                    with open(log_file,'a') as target_file: 
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): select count(*) from  infos_01DB. result: ' + str(result) + '\n')
            except sqlite3.Error as e:
                # Problem during insertion into database
                print('')
                print(exe_name + ": WARNING: Trying to execute count(*) from database '" + DB_file + "'")
                print(str(e))
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): trying to execute count(*) from database ' + DB_file + ', table infos_01DB\n')
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): ' + sql + '\n')
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): sqlite3.Error: ' + str(e) + '\n')
                checkPause()
                pass
            if result > 0:  
                try:
                    # Look for old record in database
                    sql = "SELECT idx_infos_01DB FROM infos_01DB WHERE (filename = '" + filename + "' AND location = '" + location + "' AND source = '" + source + "' AND type_of_datas = '" + type_of_datas + "' AND weighting = '" + weighting + "' AND unit = '" + unit + "' AND start = DATETIME('1899-12-30', '+" + str(start) + " days') AND stop = DATETIME('1899-12-30', '+" + str(stop) + " days'))"
                    cursor.execute(sql)
                    idx_infos_01DB = cursor.fetchone()[0]
                    if log:
                        with open(log_file,'a') as target_file: 
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): select idx_infos_01DB from infos_01DB. idx_infos_01DB: ' + str(idx_infos_01DB) + '\n' + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): updating records from database' + DB_file + ', table infos_01DB\n')
                    if not quiet:
                        print("Updating records from database '" + DB_file + "'")
                except sqlite3.Error as e:
                    # Problem during insertion into database
                    print('')
                    print(exe_name + ": WARNING: Trying to read idx_infos_01DB from database '" + DB_file + "'")
                    print(str(e))
                    if log:
                        with open(log_file,'a') as target_file:
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): WARNING: trying to read idx_infos_01DB from database ' + DB_file + ', table infos_01DB\n')
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): ' + sql + '\n')
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): sqlite3.Error: ' + str(e) + '\n')
                    checkPause()
                    pass
            else:
                if not quiet:
                    print("Inserting new records into database '" + DB_file + "'")
                    if log:
                        with open(log_file,'a') as target_file: 
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): inserting new records into database' + DB_file + ', table infos_01DB\n')
            # If records allready exists, delete all records from infos_01DB and events_01DB
            if idx_infos_01DB != -1:
                try:
                    sql1 = "DELETE FROM infos_01DB WHERE (idx_infos_01DB = '" + str(idx_infos_01DB) + "')"
                    cursor.execute(sql1)
                    sql2 = "DELETE FROM events_01DB WHERE (idx_infos_01DB = '" + str(idx_infos_01DB) + "')"
                    cursor.execute(sql2)
                    if log:
                        with open(log_file,'a') as target_file:
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): delete from infos_01DB and events_01DB. idx_infos_01DB: ' + str(idx_infos_01DB) + '\n')
                except sqlite3.Error as e:
                    # Problem during insertion into database
                    print('')
                    print(exe_name + ": WARNING: Trying to delete records from database '" + DB_file + "'")
                    print(str(e))
                    if log:
                        with open(log_file,'a') as target_file:
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): delete from infos_01DB and events_01DB. idx_infos_01DB: ' + str(idx_infos_01DB) + '\n')
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): ' + sql1 + '\n')
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): ' + sql2 + '\n')
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): sqlite3.Error: ' + str(e) + '\n')
                    checkPause()
                    pass
            try:
                # Convert Excel date format to SQL date format and write all informations into database
                sql = "INSERT INTO infos_01DB (filename, location, latitude, longitude, altitude, source, type_of_datas, weighting, unit, start, stop) values('" + filename + "','" + location + "','" + str(latitude) + "','" + str(longitude) + "','" + str(altitude) + "','" + source + "','" + type_of_datas + "','" + weighting + "','" + unit + "', DATETIME('1899-12-30', '+" + str(start) + " days'), DATETIME('1899-12-30', '+" + str(stop) + " days'))"
                cursor.execute(sql)
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): insert into database ' + DB_file + ', table infos_01DB sound records from Excel file ' + input_file + '\n')
            except sqlite3.Error as e:
                # Problem during insertion into database
                print('')
                print(exe_name + ": WARNING: Trying to insert datas into database '" + DB_file + "'")
                print(str(e))
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): WARNING: trying to insert datas into database ' + DB_file + ', table infos_01DB\n')
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): ' + sql + '\n')
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): sqlite3.Error: ' + str(e) + '\n')
                checkPause()
                pass
            conn.commit()
            # Get last idx_infos_01DB
            try:
                cursor.execute("""SELECT max(idx_infos_01DB) FROM infos_01DB""")
                idx_infos_01DB = cursor.fetchone()[0]
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): select max idx_infos_01DB from infos_01DB\n')
            except sqlite3.Error as e:
                # Problem during reading
                print('')
                print(exe_name + ": WARNING: Trying to select max from database '" + DB_file + "'")
                print(str(e))
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): trying to select max idx_infos_01DB from database ' + DB_file + ', table infos_01DB\n')
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): SELECT max(idx_infos_01DB) FROM infos_01DB\n')
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): sqlite3.Error: ' + str(e) + '\n')
                checkPause()
                pass
            if log:
                with open(log_file,'a') as target_file: 
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): stop reading header from Excel file ' + input_file + ', sheet ' + sheet.name + '\n' + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): start reading datas from Excel file ' + input_file + ', sheet ' + sheet.name + '\n')
            # read all datas from file starting after header
            for rownum in range(12,sheet.nrows-1):
                date_time = sheet.cell_value(rownum,0)
                duration = sheet.cell_value(rownum,1)
                leq = sheet.cell_value(rownum,2)
                lmax = sheet.cell_value(rownum,3)
                lmax_time = sheet.cell_value(rownum,4)
                try:
                    sel = sheet.cell_value(rownum,5)
                except:
                    sel = 0
                try:
                    # Convert Excel date and time formats to SQL formats and write all informations into database
                    sql="INSERT INTO events_01DB (datetime, duration, leq, lmax, lmax_time, sel, idx_infos_01DB) values(DATETIME('1899-12-30', '+" + str(date_time) + " days'),strftime('00:%M:%S'," + str(duration) + "),'" + str(leq) + "','" + str(lmax) + "',DATETIME('1899-12-30', '+" + str(lmax_time) + " days'),'" + str(sel) + "'," + str(idx_infos_01DB) + ")"
                    cursor.execute(sql)
                    if log:
                        with open(log_file,'a') as target_file: 
                            target_file.write('writing ' + input_file + ', ' + sheet.name + ', line ' + str(rownum) + ' to database ' + DB_file + ', table events_01DB\n')
                except sqlite3.Error as e:
                    # Problem during insertion into database, quit.
                    print('')
                    print(exe_name + ": WARNING: Trying to insert datas into database '" + DB_file + "'")
                    print(str(e))
                    if log:
                        with open(log_file,'a') as target_file: 
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): WARNING: trying to insert datas from Excel file ' + input_file + ' into database ' + DB_file + ', table events_01DB\n')
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): ' + sql + '\n')
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): sqlite3.Error: ' + str(e) + '\n')
                    checkPause()
                    pass
            conn.commit()
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readSheets(): stop reading datas from Excel file ' + input_file + ', sheet ' + sheet.name + '\n')
    if not quiet:
        print("Closing file '" + input_file + "'")
    if log:
        with open(log_file,'a') as target_file: 
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "readSheets(): closing file " + input_file + "\n")
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "readSheets(): stop reading datas from Excel file " + input_file + "\n")

def closeOnLoadOnly():
    # Exit program after loading xls file, if nothing else to do
    global write2sql
    global conn
    global DB_file
    global log
    global log_file
    global write2file
    global write2sql
    global interactive
    global quiet
    global exe_name 
    if (not write2sql):
        conn.close()
        if not quiet:
            print("Closing database '" + DB_file + "'")
        if log:
            with open(log_file,'a') as target_file: 
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'closeOnLoadOnly(): closing database ' + DB_file + '\n')
    if (not write2file) and (not write2sql) and (not interactive):
        if not quiet:
            print('')
            print('Nothing more to do')
        if log:
            with open(log_file,'a') as target_file: 
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'closeOnLoadOnly(): Nothing more to do\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'closeOnLoadOnly(): sys.exit(O): Normal E-o-P\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'closeOnLoadOnly(): END LOG FILE\n')
        checkPause()
        sys.exit(0)
            
def showWhatToDo():
    # show what to do depending on options
    global quiet
    global write2file
    global write2sql
    global output_file
    global DB_file
    global log
    global log_file 
    global interactive
    if not quiet:
        if write2file and write2sql:
            print('')
            print("Start writing JSON stream to file '" + output_file + "' and database '" + DB_file + "'")
            if log:
                with open(log_file,'a') as target_file: 
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'showWhatToDo(): start writing JSON stream to file ' +  output_file + ' and database '+ DB_file + '\n')
        elif write2file:
            print('')
            print("Start writing JSON stream to file '" + output_file + "'")
            if log:
                with open(log_file,'a') as target_file: 
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'showWhatToDo(): start writing JSON stream to file ' + output_file + '\n')
        elif write2sql:
            print('')
            print("Start saving JSON stream to database '" + DB_file + "'")
            if log:
                with open(log_file,'a') as target_file: 
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'showWhatToDo(): start saving JSON stream to database ' + DB_file + '\n')
        if interactive:
            print('')
            print('Display JSON stream on screen')
            if log:
                with open(log_file,'a') as target_file: 
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'showWhatToDo(): display JSON stream on screen\n')
        checkPause()
            
def exitPgm():
    # close database then exit
    global quiet
    global exe_name
    global log
    global log_file
    global write2sql
    global conn
    global DB_file
    if not quiet:
        print('')
        print(exe_name + ': WARNING: Program interrupted by user')
    # Close database if needed and exit program
    if log:
        with open(log_file,'a') as target_file: 
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'exitPgm(): WARNING: KeyboardInterrupt: CTRL+C detected. Program interrupted by user\n')
    try:
        if write2sql:
            conn.close()
            if not quiet:
                print('')
                print("Closing database '" + DB_file + "'")
        if log:
            with open(log_file,'a') as target_file: 
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'exitPgm(): closing database ' + DB_file + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'exitPgm(): sys.exit(O): Normal E-o-P\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'exitPgm(): END LOG FILE\n')
        checkPause()
        sys.exit(0)
    except SystemExit:
        if write2sql:
            conn.close()
            if not quiet:
                print('')
                print("Closing database '" + DB_file + "'")
        if log:
            with open(log_file,'a') as target_file: 
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'exitPgm(): closing database ' + DB_file + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'exitPgm(): os._exit(0): Normal E-o-P\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'exitPgm(): END LOG FILE\n')
        checkPause()
        os._exit(0)      
   
def readConfigFile():
    global config_file
    global quiet
    global interactive
    global write2file 
    global write2sql
    global metric
    global pause
    global exclude
    global log
    global server_PORT
    global server_IP
    global output_file
    global DB_file
    global latitude
    global longitude
    global altitude
    global station_name
    global net_HTTP
    global load2sql
    global input_file
    global log_buffer
    global msg_buffer
    global cnx_delay
    global cnx_try
    global zip_backup
    global zip_filename
    global month
    global remove
    global email_backup
    global email_toaddr
    global email_fromaddr
    global email_password
    global email_subject
    global email_body
    global email_server
    global email_port
    global report
    global max_altitude
    global max_distance
    global open_csv
    global extDB
    global extDB_filename
    global ireport
    global ireport_filename
    sql_file = False
    config_file = local_path + '/' + config_file
    config_file = os.path.normpath(config_file)
    if not os.path.isfile(config_file):
        msg_buffer = msg_buffer + "\n" + "No config file '" + config_file + "' found\n"
        log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readConfigFile(): No config file ' + config_file + ' found\n'
    else:
        msg_buffer = msg_buffer + "Loading config file '" + config_file + "'\n"
        with open(config_file) as f:
            content = f.readlines()
        for current in content:
            current = current.replace('\n',' ')
            current = current.replace('"','')
            current = current.replace("'","")
            option = current.lower()
            option = option.rstrip()
            if option == 'interactive':
                if not quiet:
                    interactive = True   
                    msg_buffer = msg_buffer + 'Config file: Interactive mode\n'
            if option == 'txt':
                write2file = True
                msg_buffer = msg_buffer + 'Config file: Write datas to text file\n'
            if option == 'sql':
                write2sql = True
                msg_buffer = msg_buffer + 'Config file: Write datas into database\n'
            if option == 'metric':
                metric = True
                msg_buffer = msg_buffer + 'Config file: Use metric system\n'
            if option == 'pause':
                if not quiet:
                    pause = True
                    msg_buffer = msg_buffer + 'Config file: Make pauses\n'
            if option == 'exclude':
                exclude = True
                msg_buffer = msg_buffer + 'Config file: Exclude inconsistents datas\n'
            if option == 'log':
                log = True
                msg_buffer = msg_buffer + 'Config file: Create a log file\n'
            option = current.split('=')[0]
            option = option.lower()
            try:
                arg = current.split('=')[1]
                arg = arg[:-1]
                if option == "sqlfile":
                    DB_file = arg
                    sql_file = True
                    msg_buffer = msg_buffer + 'Config file: Database filename ' + DB_file + '\n'
                if option == "txtfile":
                    output_file = arg
                    write2file = True
                    msg_buffer = msg_buffer + 'Config file: Write datas to text file ' + output_file + '\n' 
                if option == "ip":
                    server_IP = arg
                    msg_buffer = msg_buffer + 'Config file: Get JSON from server IP ' + server_IP + '\n'
                if option == "port":
                    server_PORT = arg
                    msg_buffer = msg_buffer + 'Config file: Get JSON from server port ' + server_PORT + '\n'
                if option == "delay":
                    cnx_delay = arg
                    cnx_try = True
                    msg_buffer = msg_buffer + 'Config file: Delay before reconnecting: ' + cnx_delay + ' second(s)\n'
                if option == "lat":
                    latitude = arg
                    msg_buffer = msg_buffer + 'Config file: Set latitude to ' + latitude + '\n'
                if option == "lon":
                    longitude = arg
                    msg_buffer = msg_buffer + 'Config file: Set longitude to ' + longitude + '\n'
                if option == "alt":
                    altitude = arg
                    msg_buffer = msg_buffer + 'Config file: Set altitude to ' + altitude + 'm\n'
                if option == "sta":
                    station_name = arg
                    msg_buffer = msg_buffer + 'Config file: Set station name to ' + station_name + '\n'
                if option == "zip":
                    zip_backup = True
                    zip_filename = arg
                    msg_buffer = msg_buffer + 'Config file: Make compressed backup file ' + zip_filename + '.zip\n'
                if option == "emailfrom":
                    email_fromaddr = arg
                    msg_buffer = msg_buffer + 'Config file: Email sender: ' + email_fromaddr + '\n'
                if option == "emailto":
                    email_toaddr = arg
                    msg_buffer = msg_buffer + 'Config file: Email recipient(s): ' + email_toaddr + '\n'
                if option == "emailpassword":
                    email_password = arg
                    msg_buffer = msg_buffer + 'Config file: Email new login password\n'
                if option == "emailsubject":
                    email_subject = arg
                    msg_buffer = msg_buffer + 'Config file: Email subject: ' + email_subject + '\n'
                if option == "emailbody":
                    email_body = arg
                    msg_buffer = msg_buffer + 'Config file: Email body: ' + email_body + '\n'
                if option == "emailserver":
                    email_server = arg
                    msg_buffer = msg_buffer + 'Config file: Email server: ' + email_server + '\n'
                if option == "emailport":
                    email_port = arg
                    msg_buffer = msg_buffer + 'Config file: Email port: ' + email_port + '\n'
            except: 
                pass
        for current in content:
            current = current.replace('\n','')
            current = current.replace('"','')
            current = current.replace("'","")
            option = current.lower()
            option = option.rstrip()
            if option == "report":
                if not sql_file:
                    msg_buffer = msg_buffer + "Config file: WARNING: No 'sqlfile' option, 'report' option will be ignored\n"
                else:
                    report = True
                    msg_buffer = msg_buffer + 'Config file: Make report from database ' + DB_file + '\n'
            if option == 'quiet':
                quiet = True 
                pause = False
                interactive = False
                msg_buffer = msg_buffer + 'Config file: Be quiet (no pauses)\n'
            if option == 'remove':
                if zip_backup:
                    remove = True
                    msg_buffer = msg_buffer + 'Config file: Remove files after compression\n'
                else:
                    remove = False
                    msg_buffer = msg_buffer + "Config file: No 'zip' option set, remove option will be ignored\n"
            if option == "email":
                if zip_backup:
                    email_backup = True
                    msg_buffer = msg_buffer + 'Config file: Send email\n'
                else:
                    msg_buffer = msg_buffer + "Config file: No 'zip' option set, send email option will be ignored\n"
            if option == 'opencsv':
                open_csv = True
            option = current.split('=')[0]
            option = option.lower()
            try:
                arg = current.split('=')[1]
                #arg = arg[:-1]
                if option == "ifile":
                    if not sql_file:
                        msg_buffer = msg_buffer + "Config file: WARNING: No 'sqlfile' option, 'ifile' option will be ignored\n"
                    else:
                        input_file = arg
                        load2sql = True
                        msg_buffer = msg_buffer + 'Config file: Load Excel file ' + input_file + ' into database ' + DB_file + '\n'
                if option == "extdb":
                        extDB_filename = arg
                        extDB = True
                if option == 'maxalt':
                    max_altitude = arg
                    msg_buffer = msg_buffer + 'Config file: Define altitude flight: ' + max_altitude 
                    if metric:
                        msg_buffer = msg_buffer + 'm\n'
                    else:
                        msg_buffer = msg_buffer + 'ft\n'
                if option == 'maxdist':
                    max_distance = arg
                    msg_buffer = msg_buffer + 'Config file: Define distance between flight and station: ' + max_distance 
                    if metric:
                        msg_buffer = msg_buffer + 'm\n'
                    else:
                        msg_buffer = msg_buffer + 'ft\n'
                if option == "ireport":
                    ireport_filename = arg
                    ireport = True
            except:
                pass
        if ireport and not extDB:
            msg_buffer = msg_buffer + "Config file: WARNING: No '--extDB' option, --ireport option will be ignored\n"
            ireport = False
            ireport_filename = 'N/A'
        if report and ireport:
            msg_buffer = msg_buffer + "Config file: WARNING: '--report' option set, --ireport option will be ignored\n"
            ireport = False
            ireport_filename = 'N/A'
        if report or load2sql:
            option_msg = ''
            if report:
                option_msg = "'report' "
            if report and load2sql:
                option_msg = option_msg + "& "
            if load2sql:
                option_msg = option_msg + "'ifile' "
            msg_buffer = msg_buffer + "Config file: WARNING: " + option_msg + "option(s), other option will be ignored\n"
            write2sql = False
            write2file = False
            interactive = False
        if extDB:
            if not report and not ireport:
                msg_buffer = msg_buffer + "Config file: WARNING: No '--report' or '--ireport' option, --extDB option will be ignored\n"
                extDB = False
                extDB_filename = 'N/A'
            if report: 
                msg_buffer = msg_buffer + "Config file: Add flights infos from database '" + extDB_filename + "' to new report file\n"
            if ireport:
                msg_buffer = msg_buffer + "Config file: Add flights infos from database '" + extDB_filename + "' to report file '" + ireport_filename + "'\n"
        if open_csv:
            if report or ireport:
                msg_buffer = msg_buffer + 'Config file: Open report file in Excel\n'
            else:
                open_csv = False
        net_HTTP = 'http://' + server_IP + ':' + server_PORT + '/data.json'
        msg_buffer = msg_buffer + 'Config file: Define JSON source: ' + net_HTTP + '\n'
        if log:
            content_log = str(content)
            content_log = content_log.replace("\\\\","\\")
            content_log = content_log.replace("['","")
            content_log = content_log.replace("']","")
            content_log = content_log.replace("\\n', '","")
            log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'readConfigFile(): loading config file ' + config_file + '\n'
            log_buffer = log_buffer + str(content_log) + '\n'
            log_buffer = log_buffer + msg_buffer

def checkDB():
    global quiet
    global exe_name
    global log
    global log_file
    global DB_file
    input_DB = os.path.normpath(DB_file)
    # Look for db in local path
    if not os.path.isfile(input_DB):
        if not quiet:
            print(exe_name + ": FATAL ERROR: Database '" + input_DB + "' not found")
            print("Can't create report")
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkDB(): FATAL ERROR: database ' + input_DB + ' not found\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "checkDB(): Can't create report\n")
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkDB(): sys.exit(1): E-o-P\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkDB(): END LOG FILE\n')
        checkPause()
        sys.exit(1)
    if not quiet:
        print("Database: Using file '" + input_DB + "'")
    if log:
        with open(log_file,'a') as target_file:
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkDB(): Using database ' + input_DB + '\n')

def checkExtDB():
    global quiet
    global exe_name
    global log
    global log_file
    global extDB_filename
    input_DB = os.path.normpath(extDB_filename)
    # Look for db in local path
    if not os.path.isfile(input_DB):
        if not quiet:
            print(exe_name + ": FATAL ERROR: External database '" + input_DB + "' not found")
            print("Can't create report with flights from external database")
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkExtDB(): FATAL ERROR: external database ' + input_DB + ' not found\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "checkExtDB(): Can't create report with flights from external database\n")
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkExtDB(): sys.exit(1): E-o-P\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkExtDB(): END LOG FILE\n')
        checkPause()
        sys.exit(1)
    if not quiet:
        print("Database: Using external file '" + input_DB + "'")
    if log:
        with open(log_file,'a') as target_file:
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkExtDB(): Using external database ' + input_DB + '\n')
            
# check xls file allready loaded
def checkDBXLS():
    global quiet
    global log
    global log_file
    global cursor
    global exe_name
    global DB_file
    global report
    global conn
    result = -1
    try:
        cursor = conn.cursor()
        # Look for old record in database
        sql = "SELECT count(*) FROM infos_01DB"
        cursor.execute(sql)
        result = cursor.fetchone()[0]
        if log:
            with open(log_file,'a') as target_file: 
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkDBXLS(): select count(*) from  infos_01DB. result: ' + str(result) + '\n')
    except sqlite3.Error as e:
        report = False
        # Problem during insertion into database
        print('')
        print(exe_name + ": FATAL ERROR: Trying to execute count(*) from database '" + DB_file + "'")
        print(str(e))
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkDBXLS(): FATAL ERROR: trying to execute count(*) from database ' + DB_file + ', table infos_01DB\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkDBXLS(): ' + sql + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkDBXLS(): sqlite3.Error: ' + str(e) + '\n')
        checkPause()
        pass
    if result == 0:
        report = False
        if not quiet:
            print('')
            print(exe_name + ': FATAL ERROR: Excel file was not allready loaded in database')
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkDBXLS(): FATAL ERROR: Excel file was not allready loaded in database\n')
    elif result > 0:
        if not quiet:
            print('Database: Found loaded Excel file in database')
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'checkDBXLS(): Found loaded Excel file in database\n')

# Make report (old)
def oldReport():
    global cursor
    global exe_name
    global log
    global log_file
    global DB_file
    global max_altitude
    global local_path
    global output_report
    max_altitude_msg = ''
    if max_altitude > 0:
       max_altitude_msg = 'AND aircrafts.altitude < ' + str(max_altitude)
    try:
        sql= "CREATE TEMP VIEW report AS SELECT  events_01DB.datetime dateStart, aircrafts.datetime dateFlight, DATETIME(events_01DB.datetime, duration) dateStop, events_01DB.leq, events_01DB.lmax, aircrafts.latitude latFlight, aircrafts.longitude lonFlight, aircrafts.altitude altFlight, infos_01DB.latitude lat01DB, infos_01DB.longitude lon01DB, infos_01DB.altitude alt01DB, MIN ( ((aircrafts.latitude-infos_01DB.latitude) * (aircrafts.latitude-infos_01DB.latitude)) + ((aircrafts.longitude - infos_01DB.longitude) * (aircrafts.longitude - infos_01DB.longitude)) + ((aircrafts.altitude - (infos_01DB.altitude / 0.3048)) * (aircrafts.altitude - (infos_01DB.altitude / 0.3048))) ) distance, icao, flight, speed, heading, location, source, idx_aircrafts, idx_events_01DB, events_01DB.idx_infos_01DB FROM events_01DB, aircrafts INNER JOIN infos_01DB ON infos_01DB.idx_infos_01DB = events_01DB.idx_infos_01DB WHERE (aircrafts.datetime BETWEEN events_01DB.datetime AND DATETIME(events_01DB.datetime, duration)) AND aircrafts.latitude != '0.0' AND aircrafts.longitude != '0.0' AND aircrafts.altitude != '0' AND icao != '' AND infos_01DB.latitude != '0.0' AND infos_01DB.latitude != '0.0' AND infos_01DB.altitude != '0' " + max_altitude_msg + " GROUP BY idx_events_01DB;"
        cursor.execute(sql)
        if not quiet:
           print("Database: Create view")
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'oldReport(): create view\n')
    except sqlite3.Error as e:
        # Problem during insertion into database
        print('')
        print(exe_name + ": FATAL ERROR: Trying to execute main sql request on database '" + DB_file + "'")
        print(str(e))
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'oldReport(): FATAL ERROR: cannot execute main sql request on database ' + DB_file + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'oldReport(): ' + sql + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'oldReport(): sqlite3.Error: ' + str(e) + '\n')
        checkPause()
        sys.exit(0)
    try:
        now = time.strftime("%Y%m%d%H%M", time.localtime())
        #Write header in csv file
        output_report = DB_file.split('.')[0] + '_' + now + '.csv'
        output_report = os.path.normpath(output_report)
        with open(output_report, 'a', newline='') as csvfile:
            spamwriter = csv.writer(csvfile, delimiter=';', quoting=csv.QUOTE_MINIMAL)
            header = ('Event start','Flight date','Event stop', 'Leq', 'Lmax', 'Flight latitude', 'Flight longitude','Flight altitude (ft)','Flight altitude (m)','Station latitude','Station longitude','Station altitude (ft)','Station altitude (m)','Distance (ft)','Distance (m)','Code HEX','Flight','Speed (NM)','Speed (km/h)','Heading','Station location', 'Source','FR24 link','FlightStats link','FlightAware link')
            spamwriter.writerow(header)
        sql= "select * from report ORDER BY dateFlight, distance"
        if not quiet:
            start_time = time.strftime("%H:%M:%S", time.localtime())
            print('Database: Start selection ' + start_time)
            print("Database: Create report file '" + output_report + "'")
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'oldReport(): start selection\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "oldReport(): create report file " + output_report + "\n")
        cursor.execute(sql)
        while True:
            row = cursor.fetchone()
            if row == None:
                break
            row_msg = 'Start:' + row[0] + ',Flight:' + row[1] + ',Stop:' + row[2] + ',leq:' + str(row[3]) + ',lmax:' + str(row[4]) + ',Dist:' + str(row[11]) + '\nICAO:' + row[12] + ',idxAir:' + str(row[18]) + ',idxEvt:' + str(row[19]) + ',idxInf01:' + str(row[20]) + ''
            lst = list(row)
            # Delete last 3 elements (index)
            lst.pop()
            lst.pop()
            lst.pop()
            # calculating distance in ft
            distance = row[11]**(0.5)
            # convert in meters
            distance_meters = distance * 0.3048
            distance = int(distance)
            distance_meters = int(distance_meters)
            lst[11] = distance
            # insert new element in list
            lst.insert(12, distance_meters)
            # Flight altitude in meters
            altSta_meters = row[7] * 0.3048
            altSta_meters = int(altSta_meters)
            # Station altitude in ft
            altitude_feets = row[10] / 0.3048
            altitude_feets = int(altitude_feets)
            lst.insert(8, altSta_meters)
            lst.insert(11, altitude_feets)
            # Speed in meters
            speed_meters = lst[17] * 1.852
            speed_meters = int(speed_meters)
            lst.insert(18, speed_meters)
            link1 = ''
            link2 = ''
            link3 = ''
            if lst[16] !='':
                link1 = 'http://fr24.com/' + lst[16]
                link2 = 'http://www.flightstats.com/go/FlightStatus/flightStatusByFlight.do?flightNumber=' + lst[16]
                link3 = 'http://flightaware.com/live/flight/' + lst[16]
                link1 = link1.rstrip()
                link2 = link2.rstrip()
                link3 = link3.rstrip()
            lst.insert(22, link1)
            lst.insert(23, link2)
            lst.insert(24, link3)
            # Convert point to comma
            for i in range(3, 15):
                lst[i] = str(lst[i]).replace(".", ",")
            # Save to "row"
            row = tuple(lst)
            with open(output_report, 'a', newline='') as csvfile:
                spamwriter = csv.writer(csvfile, delimiter=';', quoting=csv.QUOTE_MINIMAL)
                spamwriter.writerow(row)
            if log:
                with open(log_file,'a') as target_file:
                    target_file.write(row_msg + '\n')
        if not quiet: 
            stop_time = time.strftime("%H:%M:%S", time.localtime())
            print('Database: Stop selection  ' + stop_time)
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'oldReport(): stop selection\n')
    except sqlite3.Error as e:
        print('')
        print(exe_name + ": FATAL ERROR: Trying to execute select(*) from database '" + DB_file + "'")
        print(str(e))
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'oldReport(): FATAL ERROR: cannot execute sql request from database ' + DB_file + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'oldReport(): ' + sql + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'oldReport(): sqlite3.Error: ' + str(e) + '\n')
        checkPause()
        sys.exit(1)

# Make report (new)
def makeReport():
    global cursor
    global exe_name
    global log
    global log_file
    global DB_file
    global max_altitude
    global local_path
    global output_report
    cursor2 = conn.cursor()
    max_altitude_msg = ''
    if max_altitude > 0:
       max_altitude_msg = "AND altitude < '" + str(max_altitude) +"'"
    try:
        sql = "CREATE TEMP VIEW report AS SELECT idx_aircrafts, datetime dateFlight, latitude latFlight, longitude lonFlight, altitude altFlight, icao, flight, speed, heading FROM aircrafts WHERE latitude != '0.0' AND longitude != '0.0' AND altitude != '0' AND icao != '' " + max_altitude_msg + ";"
        if not quiet:
            print("Database: Creating view")
            print("Database: Checking data integrity (flights)")
            if max_altitude > 0:
                print('Database: Filtering on maximum altitude : ' + str(max_altitude) + 'ft')
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'makeReport(): Creating view. Checking data integrity (flights).\n')
                if max_altitude > 0:
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'makeReport(): maximum altitude:' + str(max_altitude) + 'ft')
        cursor.execute(sql)
    except sqlite3.Error as e:
        # Problem during creation of the report view
        print('')
        print(exe_name + ": FATAL ERROR:  Creating view. Checking data integrity (flights).")
        print(sql)
        print(str(e))
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'makeReport(): FATAL ERROR: cannot create view report into database ' + DB_file + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'makeReport(): ' + sql + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'makeReport(): sqlite3.Error: ' + str(e) + '\n')
        checkPause()
        sys.exit(0)
    conn.commit()
    
    # read all 01DB station information and put in memory table
    infos01DB = []
    try:
        sql= "SELECT * FROM infos_01DB ORDER BY idx_infos_01DB ASC"
        cursor.execute(sql)
        if not quiet:
           print("Database: Reading 01DB station informations")
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'makeReport(): Reading 01DB station informations\n')
        rows = cursor.fetchall()
        infos01DB = []
        vide = ('')
        for row in rows:
            while (len(infos01DB) < row[0]):
                infos01DB.append(vide)
            infos01DB.append(row)
    except sqlite3.Error as e: 
        print('')
        print(exe_name + ": FATAL ERROR:  Can't read 01DB station information")
        print(str(e))
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "makeReport(): FATAL ERROR: Can't read infos_01DB table from database " + DB_file + "\n")
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'makeReport(): ' + sql + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'makeReport(): sqlite3.Error: ' + str(e) + '\n')
        checkPause()
        sys.exit(0)
    
    # Open txt report file
    now = time.strftime("%Y%m%d%H%M", time.localtime())
    #Write header in csv file
    output_report = DB_file.split('.')[0] + '_' + now + '.csv'
    output_report = os.path.normpath(output_report)
    with open(output_report, 'a', newline='') as csvfile:
        spamwriter = csv.writer(csvfile, delimiter=';', quoting=csv.QUOTE_MINIMAL)
        header = ('Event start','Flight date','Event stop', 'Leq', 'Lmax', 'Flight latitude', 'Flight longitude','Flight altitude (ft)','Flight altitude (m)','Station latitude','Station longitude','Station altitude (ft)','Station altitude (m)','Distance (ft)','Distance (m)','Code HEX','Flight','Speed (NM)','Speed (km/h)','Heading','Station location', 'Source','FR24 link','FlightStats link','FlightAware link')
        spamwriter.writerow(header)
        if not quiet:
            start_time = time.strftime("%H:%M:%S", time.localtime())
            print('Database: Start selection ' + start_time)
            print("Database: Create report file '" + output_report + "'")
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'oldReport(): start selection\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "oldReport(): create report file " + output_report + "\n")
    # Read all sound alarms
    try:
        sql= "SELECT datetime, DATETIME(datetime, duration), leq, lmax, idx_infos_01DB, count(*) FROM events_01DB"
        cursor.execute(sql)
        if not quiet:
           print("Database: Reading sound alarm events")
           print("Database: Found " + str(row[5]) + " event(s)")
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'makeReport(): Reading sound alarm events\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'makeReport(): Found' + str(row[5]) + 'event(s)\n')
        sql= "SELECT datetime, DATETIME(datetime, duration), leq, lmax, idx_infos_01DB FROM events_01DB"
        cursor.execute(sql)
        if not quiet:
           print("Database: Start " + time.strftime("%d/%m/%Y - %H:%M:%S ", time.localtime()))
        while True:
            row = cursor.fetchone()
            if row == None:
                break
            # Look for min distance for each alarm
            try:
                if not quiet:
                    print('.',end="",flush=True)
                mindist = 0
                minrow = (0,'',0.0,0.0,0,'','',0,0)
                sql2 = "SELECT idx_aircrafts, dateFlight, latFlight, lonFlight, altFlight, icao, flight, speed, heading  FROM report WHERE dateFlight BETWEEN '" + row[0] + "' AND '" + row[1] + "'";
                cursor2.execute(sql2)
                rows = cursor2.fetchall()
                for row2 in rows:
                    distance = ((row2[2] - infos01DB[row[4]][3])**2 + (row2[3] - infos01DB[row[4]][4])**2 + (row2[4] - (infos01DB[row[4]][5] / 0.3048))**2 )
                    distance = int(distance ** 0.5)
                    if (distance < mindist) or (mindist == 0):
                        mindist = distance
                        minrow = row2
            except sqlite3.Error as e:
                pass
            if (minrow[0] != 0):
                # Make new record with all datas
                lst = []
                lst.append(row[0])      #Event start
                lst.append(minrow[1])   #Flight date
                lst.append(row[1])      #Evebt stop
                lst.append(row[2])      #Leq
                lst.append(row[3])      #Lmax
                lst.append(minrow[2])   #Flight lat            
                lst.append(minrow[3])   #Flight lon
                lst.append(minrow[4])   #Flight alt (ft)
                altFlight_meters = minrow[4] * 0.3048
                altFlight_meters = int(altFlight_meters)
                lst.append(altFlight_meters)       #Fligth alt (m)
                lst.append(infos01DB[row[4]][3])   #Station lat            
                lst.append(infos01DB[row[4]][4])   #Station lon
                altStation_feets = infos01DB[row[4]][5] / 0.3048
                altStation_feets = int(altStation_feets)
                lst.append(altStation_feets)      #Station alt(ft)
                lst.append(infos01DB[row[4]][5])   #Station alt (m)
                mindist_ft = int(mindist)
                lst.append(mindist_ft)             #Distance (ft)
                distance_meters = mindist * 0.3048
                distance_meters = int(distance_meters)
                lst.append(distance_meters)        #Distance (m)
                lst.append(minrow[5])   #'Code HEX'
                lst.append(minrow[6])   #'Flight'
                lst.append(minrow[7])   #'Speed (NM)'
                speed_kmph = minrow[7] * 1.852
                speed_kmph = int(speed_kmph)
                lst.append(speed_kmph)  #'Speed (km/h)'
                lst.append(minrow[8])   #'Heading'
                lst.append(infos01DB[row[4]][2])      #'Station location'
                lst.append(infos01DB[row[4]][6])      #'Source'
                link1 = ''
                link2 = ''
                link3 = ''
                if (minrow[6] != ''):
                    link1 = 'http://fr24.com/' + minrow[6]
                    link2 = 'http://www.flightstats.com/go/FlightStatus/flightStatusByFlight.do?flightNumber=' + minrow[6]
                    link3 = 'http://flightaware.com/live/flight/' + minrow[6]
                    link1 = link1.rstrip()
                    link2 = link2.rstrip()
                    link3 = link3.rstrip()
                lst.append(link1)           #'FR24 link'
                lst.append(link2)           #'FlightStats link'
                lst.append(link3)           #'FlightAware link'
                for i in range(3, 15):
                    lst[i] = str(lst[i]).replace(".", ",")
                row = tuple(lst)
                # Write datas in file
                with open(output_report, 'a', newline='') as csvfile:
                    spamwriter = csv.writer(csvfile, delimiter=';', quoting=csv.QUOTE_MINIMAL)
                    spamwriter.writerow(row)
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(str(row) + '\n')
        if not quiet:
            print()
    except sqlite3.Error as e:
        print('')
        print(exe_name + ": FATAL ERROR:  Can't read flights")
        print(str(e))
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "makeReport(): FATAL ERROR: Can't read report table from database " + DB_file + "\n")
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'makeReport(): ' + sql + '\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'makeReport(): sqlite3.Error: ' + str(e) + '\n')
        checkPause()
        sys.exit(0)
    if not quiet: 
        stop_time = time.strftime("%H:%M:%S", time.localtime())
        print('Database: Stop selection  ' + stop_time)
    if log:
        with open(log_file,'a') as target_file:
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'makeReport(): stop selection\n')
    
def findFlight():
    global output_report
    global cursor
    global exe_name
    global quiet
    global log_file
    global log
    global conn
    tmpfile = False
    cursor = conn.cursor()
    report_filename = os.path.normpath(output_report)
    if not os.path.isfile(report_filename):
        if not quiet: 
            print('')
            print(exe_name + ": FATAL ERROR: No report file '" + report_filename + "'")
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "findFlight(): FATAL ERROR: No report file '" + report_filename + "'\n")
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'findFlight(): sys.exit(2): E-o-P\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'findFlight(): END LOG FILE\n')
        conn.close()
        checkPause()
        sys.exit(2)
    else:
        i=0;
        first = True
        if not quiet:   
            print('External database: Reading report file')
        
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "findFlight(): external db: reading report file\n")
        with open(report_filename) as f:
            content = f.readlines()
        for current in content:
            if first:
                first = False
                timestamp = time.strftime("%H%M%S", time.localtime())
                if not quiet:
                    start_time = time.strftime("%H:%M:%S", time.localtime())
                    
                    print('External database: Start selection ' + start_time)
                    print("External database: Adding flights infos to '" + output_report + "'")
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'findFlight(): start selection\n')
                        target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "findFlight(): adding flights infos to " + output_report + "\n")
                newfilename =  output_report[:-4]
                output_report = newfilename + 'F.csv'
                try:
                    os.remove(output_report)
                    if log:
                        with open(log_file,'a') as target_file:
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'findFlight(): deleting old  report file ' + output_report + '\n')
                except OSError as e:
                    output_report = newfilename + 'F' + timestamp + '.csv'
                    if log:
                        with open(log_file,'a') as target_file:
                            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "findFlight(): WARNING: Can't delete " + output_report + " file\n")
                    
                time.sleep(.5)
                with open(output_report, 'a', newline='') as csvfile:
                    spamwriter = csv.writer(csvfile, delimiter=';', quoting=csv.QUOTE_MINIMAL)
                    header = ('Event start','Flight date','Event stop', 'Leq', 'Lmax', 'Flight latitude','Flight longitude','Flight altitude (ft)','Flight altitude (m)','Station latitude','Station longitude','Station altitude (ft)','Station altitude (ft)','Distance (ft)','Distance (m)','Code HEX','Flight','Speed (NM)','Speed (km/h)','Heading','Station location', 'Source', 'FR24 link','FlightStats link','FlightAware link', 'AircraftID', 'ModeS', 'Callsign', 'ModeSCountry', 'Registration', 'ICAOTypeCode', 'SerialNo', 'OperatorFlagCode', 'Manufacturer', 'Type', 'Country', 'RegisteredOwners' )
                    spamwriter.writerow(header)
                tmpfile = True
                if log:
                    with open(log_file,'a') as target_file:
                        target_file.write(str(header) + '\n')
            else:
                #Get code HEX/ICAO
                option = current.split(';')[15]
                if option != '':
                    # Look for flight in external database
                    try:
                        now = time.strftime("%Y%m%d%H%M", time.localtime())
                        sql= "SELECT Aircraft.AircraftID, ModeS, ModeSCountry, Registration, ICAOTypeCode, SerialNo, OperatorFlagCode, Manufacturer, Type, Country, RegisteredOwners, Callsign, count(*) FROM Aircraft LEFT JOIN Flights ON Aircraft.AircraftID = Flights.AircraftID WHERE ModeS = '" + option + "' LIMIT 1"
                        cursor.execute(sql)
                        while True:
                            row = cursor.fetchone()
                            if row == None:
                                break
                            lst = list(row) 
                            lrow = current + ';' + str(lst[0]) + ';' + str(lst[1]) + ';' + str(lst[11]) + ";" + str(lst[2]) + ';' + str(lst[3]) + ';' + str(lst[4]) + ';' + str(lst[5]) + ';' + str(lst[6]) + ';' + str(lst[7]) + ';' + str(lst[8]) + ';' + str(lst[9]) + ';' + str(lst[10])
                            lrow = lrow.replace('None','')
                            lrow = lrow.replace('\n','')
                            if lst[12] > 0:
                                with open(output_report, 'a') as target_file:
                                    target_file.write(lrow + '\n')
                                if log:
                                    with open(log_file,'a') as target_file:
                                        target_file.write(lrow + '\n')
                            else:
                                lrow = current
                                lrow = lrow.replace('\n','')
                                with open(output_report, 'a') as target_file:
                                    target_file.write(lrow + '\n')
                                if log:
                                    with open(log_file,'a') as target_file:
                                        target_file.write(lrow + '\n')
                    except sqlite3.Error as e:
                        print('')
                        print(exe_name + ": FATAL ERROR: Trying to execute select(*) from database '" + DB_file + "'")
                        print(str(e))
                        if log:
                            with open(log_file,'a') as target_file:
                                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'findFlight(): FATAL ERROR: cannot execute sql request from database ' + DB_file + '\n')
                                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'findFlight(): ' + sql + '\n')
                                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'findFlight(): sqlite3.Error: ' + str(e) + '\n')
                        checkPause()
                        sys.exit(1)
                else:
                    # Do not have ICAO/hex code (never?)
                    lrow = current
                    lrow = lrow.replace('\n','')
                    with open(output_report, 'a') as target_file:
                        target_file.write(lrow + '\n')
                    if log:
                        with open(log_file,'a') as target_file:
                            target_file.write(lrow + '\n')
        if not quiet: 
            stop_time = time.strftime("%H:%M:%S", time.localtime())
            print('External database: Stop selection  ' + stop_time)
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'findFlight(): stop selection\n')
                        
# Check unit and convert max altitude, max distance if needed
def setUnit():
    global max_altitude
    global metric
    global max_distance
    global change_to_metric
    global quiet
    global log
    global logfile
    global log_buffer
    if max_altitude == 0:
        max_altitude_msg = 'N/A'
    else:
        if not metric:
            max_altitude_msg = str(max_altitude) + 'ft/'
            max_altitude_meters = int(max_altitude * 0.3048)
            max_altitude_msg = max_altitude_msg + str(max_altitude_meters) + 'm'
        else:   
            max_altitude_meters = max_altitude
            max_altitude = int(max_altitude / 0.3048)
            max_altitude_msg = str(max_altitude) + 'ft/' + str(max_altitude_meters) + 'm'
    if max_distance == 0:
        max_distance_msg = 'N/A'
    else:
        if not metric:
            max_distance_msg = str(max_distance) + 'ft/'
            max_distance_meters = int(max_distance * 0.3048)
            max_distance_msg = max_distance_msg + str(max_distance_meters) + 'm'
        else:   
            max_distance_meters = max_distance
            max_distance = int(max_distance / 0.3048)
            max_distance_msg = str(max_distance) + 'ft/' + str(max_distance_meters) + 'm'
    metric_msg = ''
    if ((max_distance > 0) or (max_altitude > 0)) and change_to_metric:
        metric_msg = 'WARNING: unit was changed in command line'
    if not quiet:
        print('')
        print('Creating report...')
        if change_to_metric:
            print(metric_msg)
        print('Maximum altitude flight:', max_altitude_msg)
        print('Maximum distance between flight and station:', max_distance_msg)
    if log:
        with open(log_file,'a') as target_file:
            target_file.write(log_buffer)
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'setUnit(): creating report \n')
            if change_to_metric:
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'setUnit():' + metric_msg + '\n')
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'setUnit(): Maximum altitude:' + max_altitude_msg + '\n')
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'setUnit(): Maximum distance:' + max_distance_msg + '\n')
    checkPause()

def startLog():
    global log_buffer
    global exe_name
    global version
    global prog_name
    global prog_name_report
    global is_dump1090sql
    log_buffer = time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'main(): START LOG FILE\n'
    log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'main(): script ' + exe_name + ' v' + str(version) + '\n'
    log_buffer = log_buffer + time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'main(): python ' + str(sys.version_info) + '\n'  
    if is_dump1090sql:
        msg_buffer = prog_name + ' - version ' + version + '\n\n'  
    else:   
        msg_buffer = prog_name_report + ' - version ' + version + '\n\n'  

def openCSV():
    global quiet
    global log
    global log_file
    global output_report
    if not quiet:
        print("Database: Open report file in Excel")
    if log:
        with open(log_file,'a') as target_file:
            target_file.write(log_buffer)
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "openCSV(): open report file '" + output_report + "' in Excel\n")
    try:            
        os.system('start excel.exe "%s"' % (output_report, ))
    except:
        if not quiet:
            print('')
            print(exe_name + ": WARNING: Can't open csv file '" + output_report + "' in Excel")
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(log_buffer)
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "openCSV(): WARNING: can't open csv file '" + output_report + "' in Excel\n")
                            
def addFlight():
    global quiet
    global log
    global log_file
    global extDB_filename
    global output_report
    if not quiet:
        print('')
        print("External database: Opening external database '" + extDB_filename + "'")
    if log:
        with open(log_file,'a') as target_file: 
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'addFlight(): external db: opening external database ' + extDB_filename + '\n')
    openExtDB()
    if not quiet:
        print("External database: Looking for flights in database")
    if log:
        with open(log_file,'a') as target_file: 
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'addFlight(): external db: looking for flights in database\n')
    findFlight()
    if not quiet:
        print("External database: Closing external database")
    if log:
        with open(log_file,'a') as target_file: 
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'addFlight(): external db: closing external database\n')
    conn.close()

def closeReport():
    global quiet
    global log
    global log_file
    global DB_file
    global output_report
    global ireport
    if not quiet:
        if not ireport:
            print("Database: Closing file '" + DB_file + "'")
            print('')
        print("New report file: '" + output_report + "'")
        print("Job done!")
    if log:
        with open(log_file,'a') as target_file: 
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'closeReport(): closing file ' + DB_file + '\n')
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "closeReport(): new report file: '" + output_report + "'\n")
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'closeReport(): job done!\n')
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'closeReport(): os._exit(0): Normal E-o-P\n')
            target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'closeReport(): END LOG FILE\n')
    checkPause()
    # End of program
    sys.exit(0)    
#    
# Main function, everythings start here
#
if __name__ == "__main__":
    if (exe_name == 'dump1090sql.py') or (exe_name == 'dump1090sql.exe') or (exe_name == 'dump1090sql'):
        is_dump1090sql = True
    # Prepare log file with header, show program version
    startLog()    
    # Read config file
    readConfigFile()
    # Read Cmd line a parse it
    parseCmdLine(sys.argv[1:])
    # Show previous messages
    if (msg_buffer != '') and not quiet:
        print(msg_buffer)
    msg_buffer = ''
    # Nothing to do?
    if (not write2file) and (not write2sql) and (not load2sql) and (not interactive) and (not report) and (not ireport):
        if not quiet:
            print('')
            print(exe_name + ': WARNING: Nothing to do, check options!')
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(log_buffer)
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseCmdLine(): WARNING: Nothing to do, check options!\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseCmdLine(): sys.exit(1): E-o-P\n')
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'parseCmdLine(): END LOG FILE\n')
        checkPause()
        sys.exit(1)   
    max_altitude = int(float(max_altitude))
    max_distance = int(float(max_distance))
    # Need to access server? test connexion
    if write2file or write2sql:
        # --sql or --txt, create all folders and sub folders (2015\201506...)
        createFolders()
   # Write log if something in log_buffer
    if log and (log_buffer != ''): 
        with open(log_file,'a') as target_file:
            target_file.write(log_buffer)
    log_buffer = ''
    if not quiet:
        # Show all options on screen
        showOptions()
    if write2file or write2sql or interactive:
        pingServer()
    opener = FancyURLopener()
    if extDB:
        checkExtDB()
    if write2sql or load2sql or report:
        # Check if database exist
        checkDB()
        # Open (or create new) database
        openBase()
        # Create all tables (if needed)
        createTables()
        # Create index on aircrafts table (if needed)
        createIndex()
        # Save std informations (date/time/GPS coord...)
        insertInfos()
        # Get new record index (from std infos)
        selectMaxInfos()
    # Load Excel file (if needed)
    if load2sql:
        # Test if file exist in local dir or in sub dir (2015\2015...)
        checkXLSfile()
        # Open XLS file
        openXLSfile()
        # Read all sheets and load into database
        readSheets()
        if (not report) and (not ireport):
            # close connexion
            conn.close()
            if not quiet:
                print("Database: Closing file '" + DB_file + "'")
                print('')
                print("Job done!")
            if log:
                with open(log_file,'a') as target_file: 
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'main(): closing file ' + DB_file + '\n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'main(): job done!\n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'main(): os._exit(0): Normal E-o-P\n')
                    target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + 'main(): END LOG FILE\n')
            checkPause()
            # End of program
            sys.exit(0) 
        checkPause()
    if report:
        # Check if unit is ft or meters
        setUnit()
        if not quiet:
            print("Database: File '" + DB_file + "' open with success")
        if log:
            with open(log_file,'a') as target_file:
                target_file.write(log_buffer)
                target_file.write(time.strftime("%Y%m%d.%H%M%S ", time.localtime()) + "main(): file '" + DB_file + "' open with success\n")
        # check xls file allready loaded
        checkDBXLS()
        if report:
            # Make report
            if onesql_report:
                oldReport()
            else:
                makeReport()
            conn.close()
            if extDB:
                addFlight()
            if open_csv:
                openCSV()
        else:
            conn.close()
        closeReport()
    if ireport:
        output_report = ireport_filename
        checkCSV()
        addFlight()
        if open_csv:
            openCSV() 
        closeReport()    
    showWhatToDo()
    # Read HTTP JSON stream and update database or file
    try:
        getNewJSON()
    except KeyboardInterrupt:
        # CTRL+C detected ?
        exitPgm()
          