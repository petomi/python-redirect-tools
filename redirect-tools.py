# -*- coding: utf-8 -*-
import sys
import os
import re
import pandas #https://pypi.python.org/pypi/pandas/0.20.1
import requests #https://pypi.python.org/pypi/requests/2.14.2
import configparser #https://pypi.python.org/pypi/configparser
from xlutils.copy import copy   #https://pypi.python.org/pypi/xlutils/2.0.0
from xlrd import open_workbook #https://pypi.python.org/pypi/xlrd
from xlwt import easyxf #https://pypi.python.org/pypi/xlwt/1.2.0

# read from config file
config= configparser.ConfigParser()
config.read('settings.cfg')

input_file, URL_column, redirect_column, keep_original_comments, save_every_row, check_for_sharepoint_404, verify_SSL, new_root_domain, complex_regex, strip_query, old_root_domains = str(config['EXCEL FILE']['input_file']), str(config['EXCEL FILE']['URL_column']), str(config['EXCEL FILE']['redirect_column']), (config.getboolean('EXCEL FILE','keep_original_comments')), (config.getboolean('EXCEL FILE','save_every_row')), (config.getboolean('TESTING','check_for_sharepoint_404')), (config.getboolean('TESTING','verify_SSL')), str(config['RULE CREATION']['new_root_domain']), (config.getboolean('RULE CREATION','complex_regex')), (config.getboolean('RULE CREATION','strip_query')) dict(config.items('OLD ROOT DOMAINS'))

def __sanitize_URLs__(rule, isMap, isHtaccess, isRegex):

    current_redirect = str(rule[0]).lower().strip(" ><").rstrip('/')
    future_redirect = str(rule[1]).lower().strip(" ><").rstrip('/')

    domain_list = old_root_domains.values()
    for domain in domain_list:
        current_redirect = current_redirect.replace(domain, "")
    if(isRegex):
        current_redirect = current_redirect.replace(".", "\.").lstrip('\/').rstrip('\/').replace("/", "\/")
    elif(isMap):
        current_redirect = current_redirect.replace(u"\u2018", "'").replace(u"\u2019", "'").replace(u"\u201c",'').replace(u"\u201d", '').replace('"', '').replace(u"\u0026", "&amp;")

        if(strip_query):
            current_redirect = re.sub(r'(\?).*', r'', current_redirect).rstrip('/') #strip query from any incoming url

        future_redirect = future_redirect.replace(u"\u2018", "'").replace(u"\u2019", "'").replace(u"\u201c",'').replace(u"\u201d", '').replace('"', '').replace(u"\u0026", "&amp;")
    elif(isHtaccess):
        current_redirect = current_redirect.replace("%23", "#")
        future_redirect = future_redirect.replace("%23", "#")
    return current_redirect, future_redirect

def __write_rules_to_file__(isRegexRule, isMap, isHtaccess, output_file, format_syntax):
    cols_to_parse=str(URL_column + ", " + redirect_column)
    rf = pandas.read_excel(input_file, index_col=None, parse_cols=cols_to_parse)
    list_of_redirects = list(rf.values)
    previous_URLS = []
    file_to_write = open(output_file, "w+")
    to_prepend = format_syntax[0]
    to_insert = format_syntax[1]
    to_append = format_syntax [2]
    line_endings = "\r\n"
    if(isHtaccess):
        line_endings = "\n"

    if(isMap):
        file_to_write.write('<rewriteMaps>\r\n\t<rewriteMap name="Redirects">\r\n')
    for rule in list_of_redirects:
        current_redirect, future_redirect = __sanitize_URLs__(rule, isMap, isHtaccess, isRegexRule)
        if(isRegexRule):
            to_insert = str('" stopProcessing="true">\r\n\t<match url="(' + current_redirect + ')(.*)" ignoreCase="true"/>\r\n\t<action type ="Redirect" url="' + new_root_domain)
            if(complex_regex):
                to_append = str('{R:2}" redirectType = "Permanent" />\r\n</rule>')
            else:
                to_append = str( '" redirectType = "Permanent" />\r\n</rule>')
        print (current_redirect + " = " + future_redirect)
        if(current_redirect not in previous_URLS):
            file_to_write.write(str(to_prepend + current_redirect + to_insert + future_redirect + to_append + line_endings))
        else:
            print("DUPLICATE ENTRY")
        previous_URLS.append(current_redirect)
    if(isMap):
        file_to_write.write('\t</rewriteMap>\r\n</rewriteMaps>')
    file_to_write.close()

def __test_redirects__():
    df = pandas.read_excel(input_file, index_col=None, parse_cols=URL_column)
    print(df) #prints summary of column with URLs to test
    list_of_links = list(df.values.flatten())
    requests.packages.urllib3.disable_warnings() #get rid of "SSL cert check disabled" warnings
    workbook = open_workbook(input_file, formatting_info=True) #only works with .xls documents currently, as per xlrd documentation
    write_copy = copy(workbook)
    sheet = write_copy.get_sheet(0)
    sheet.cell_overwrite_ok = True
    row_number = 1

    for link in list_of_links:
        current_link = str(link).strip() #update link from spreadsheet
        try:
            print("Requested URL: " + current_link)
            r = requests.get(current_link, verify=verify_SSL) #attempt to get URL + ignore testing SSL cretificate
            print("Redirect URL: " + str(r.url))
            status_code = str(r.status_code)
            if(check_for_sharepoint_404): #if enabled, check for standard sharepoint 404 page
                if(status_code == "200"):
                    try:
                        title = (str(r.text)).split('<meta name="description" content="')[1].split('" />')[0] #parse html
                        page_title = (title[3:-3])
                        if(page_title=="404 Error"):
                            print("URL is generic Sitepoint 404 page!")
                            status_code = "404"
                            if(not keep_original_comments):
                                sheet.write(row_number, 3, "Sharepoint 404, page does not exist.")
                    except:
                        pass
            print("Status code: " + status_code)
            sheet.write(row_number, 1, r.url) #write URL redirect
            sheet.write(row_number, 2, status_code) #write status code
        except requests.exceptions.RequestException as e: #catch failed HTTP request
            status_code = "404" #return a 404; this should be the only error triggering an exception
            print("ERROR: " + str(e))
            print("Redirect URL: N/A")
            print("Status code: " + status_code)
            sheet.write(row_number, 3, str(e)) #write error to comments
            sheet.write(row_number, 2, status_code) #write response code
            sheet.write(row_number, 1, "") #write blank redirect url
        if(save_every_row):
            write_copy.save(os.path.splitext(input_file)[0] + '-OUT' + os.path.splitext(input_file)[-1]) #saves excel doc after every row if enabled
        row_number+=1 #increment row number so next loop cycle writes to next row in excel sheet
        print("===========")
    write_copy.save(os.path.splitest(input_file)[0] + '-OUT' + os.path.splitext(input_file)[-1]) #save file after all changes are made
    print("===================LIST COMPLETE=========================")

def __create_redirect_map__():
    print("=====REDIRECT MAP======")
    isRegexRule = False
    isMap = True
    isHtaccess = False
    output_file = "rewritemaps.config"
    format_syntax = ['\t\t<add key="', '" value="', '" />\r\n']
    __write_rules_to_file__(isRegexRule, isMap, isHtaccess, output_file, format_syntax)
    print("===================REDIRECT MAP CREATED======================")

def __create_htaccess__():
    print("=====HTACCESS CREATOR======")
    isRegexRule = False
    isMap = False
    isHtaccess = True
    output_file = ".htaccess"
    format_syntax = ['Redirect 301 ', " ", ""]
    __write_rules_to_file__(isRegexRule, isMap, isHtaccess, output_file, format_syntax)
    print("===================HTACCESS CREATED======================")

def __create_redirect_rules__():
    print("=====REDIRECT RULES======")
    isRegexRule = True
    isMap = False
    isHtaccess = False
    output_file = "rewriterules.txt"
    format_syntax = [ '<rule name="', '', '']
    __write_rules_to_file__(isRegexRule, isMap, isHtaccess, output_file, format_syntax)
    print("===================REDIRECT RULES CREATED======================")

def __help__():
    print("This script takes at least one parameter, depending on the desired function.")
    print("Available functions:")
    print("test                     -           tests redirects and logs results to new file, according to options in settings.cfg")
    print("create_map               -           creates IIS redirect rule map according to links in spreadsheet and options in settings.cfg")
    print("create_htaccess          -           creates htaccess redirect rules according to spreadsheet")
    print("create_rules             -           creates regex IIS rewrite rules from spreadsheet")

# CALL HANDLING #
if (len(sys.argv) < 2):
    __help__()
elif((sys.argv[1]).lower() == "test"):
    __test_redirects__()
elif((sys.argv[1]).lower() == "create_map"):
    __create_redirect_map__()
elif((sys.argv[1]).lower() == "create_htaccess"):
    __create_htaccess__()
elif((sys.argv[1]).lower() == "create_rules"):
    __create_redirect_rules__()
else:
    __help__()
