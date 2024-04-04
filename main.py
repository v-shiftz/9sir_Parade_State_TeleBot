import logging
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from dateutil.relativedelta import relativedelta
from datetime import datetime
from collections import Counter
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters
import re
# Enable logging
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    level=logging.INFO)

logger = logging.getLogger(__name__)
PORT = int(os.environ.get('PORT', '8443'))
TOKEN = "YOUR-TOKEN-HERE"
scope = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.file'
    ]
file_name = 'client_key.json'
creds = ServiceAccountCredentials.from_json_keyfile_name(file_name, scope)
client = gspread.authorize(creds)


def start(update, context):
    if not getid(update.message.chat.id):
        update.message.reply_text("you are not verified")
        return
    update.message.reply_text("you are verified")


def getid(getid):
    id = str(getid)
    first_coy_id = {"telegram_id": "name"}
    # this can be expanded if need be.
    bn_hq_id = {""}
    if id in first_coy_id or id in second_coy_id or id in third_coy_id or id in bn_hq_id:
        return True
    else:
        return False


def gettext(update, context):
    if not getid(update.message.chat.id):
        update.message.reply_text("you are not verified")
        return
    raw_text = update.message.text
    split_text = raw_text.split('\n')
    length = len(split_text)
    date = split_text[0][26:28]+"/"+split_text[0][24:26]+"/"+split_text[0][28:30]
    checking = split_text[0][0:23]
    coy = split_text[0][0:3]
    if length < 13:
        return
    if checking != "1st Coy UFD/RSO/MA list" and checking != "2nd Coy UFD/RSO/MA list" and checking != "3rd Coy UFD/RSO/MA list":
        print("fail")
        return
    else:
        print("pass")
        update.message.reply_text("Inserting Data")
        i = 0
        while True:
            if split_text[i][0:3] == "UFD":
                if split_text[i][-2:] == "00":
                    i += 2
                    continue
                else:
                    if coy == "1st":
                        sheet = client.open('9SIR MC Tracking').worksheet("1st Coy MC Data Raw")
                    elif coy == "2nd":
                        sheet = client.open('9SIR MC Tracking').worksheet("2nd Coy MC Data Raw")
                    elif coy == "3rd":
                        sheet = client.open('9SIR MC Tracking').worksheet("3rd Coy MC Data Raw")
                    else:
                        update.message.reply_text("Company not specified.")
                        return
                    while True:
                        if split_text[i+1][0:7] == "RSO/RSI":
                            break
                        else:
                            """if split_text[i] == "": if split text is empty string do something"""
                            if split_text[i+1] != "":
                                eachline = split_text[i+1]
                                if len(eachline) < 10:
                                    camp = eachline.split(':')[0].upper()
                                    i += 1
                                    print(camp)
                                    if camp.isspace():
                                        continue
                                    else:
                                        if coy == "1st" and (camp == "DTTB" or camp == "PEH" or camp == "CLEMENTI" or camp == "RRRC" or camp == "MAJU" or camp == "KC" or camp == "KHC" or camp == "MOWBRAY" or camp == "5AD1" or camp == "5AD2" or camp == "SGC" or camp == "HQ"):
                                            continue
                                        elif coy == "2nd" and (camp == "CCK" or camp == "LCK" or camp == "SC" or camp == "MHC" or camp == "MWC" or camp == "MC2" or camp == "JC1" or camp == "JC2" or camp == "SAFTI" or camp == "PLAD" or camp == "PLC" or camp == "TNB" or camp == "HQ"):
                                            continue
                                        elif coy == "3rd" and (camp == "JGL" or camp == "ROT2" or camp == "ROT3" or camp == "HQ"):
                                            continue
                                        else:
                                            update.message.reply_text("Your camps and coy does not match up. Please check your camps again")
                                            return

                                else:
                                    name = eachline.split('(')[0].strip().upper()
                                    vocation = eachline[eachline.find("(")+1:eachline.find(")")]
                                    check_name = eachline.split('(')[0][3:].strip().upper()
                                    check_date_start = eachline.split(':')[1][1:7]
                                    check_date_end = eachline.split(':')[1][11:17]
                                    start_date = eachline.split(':')[1][3:5]+"/"+eachline.split(':')[1][1:3]+"/"+eachline.split(':')[1][5:7]
                                    end_date = eachline.split(':')[1][13:15]+"/"+eachline.split(':')[1][11:13]+"/"+eachline.split(':')[1][15:17]
                                    mc_reason = eachline.lower().split("for")[1].strip()
                                    start = datetime.strptime(start_date.replace("/", ""), "%m%d%y")
                                    end = datetime.strptime(end_date.replace("/", ""), "%m%d%y")
                                    days = (end - start).days + 1
                                    list_all = sheet.get_all_values()
                                    for x in range(len(list_all)):
                                        if list_all[x][1][4:] == check_name and (list_all[x][3] == check_date_start or list_all[x][4] == check_date_end):
                                            break
                                        else:
                                            continue
                                    else:
                                        row = [date, name, mc_reason, str(start_date), str(end_date), camp, vocation, check_name, days, 0, "=$I2-$J2", 0]
                                        index = 2
                                        sheet.insert_row(row, index, value_input_option='USER_ENTERED')
                                    i += 1
                            else:
                                i += 1

            elif split_text[i][0:7] == "RSO/RSI":
                if split_text[i][-2:] == "00":
                    i += 2
                    continue
                else:
                    if coy == "1st":
                        sheet = client.open('9SIR MC Tracking').worksheet("1st Coy RSO/RSI Data")
                    elif coy == "2nd":
                        sheet = client.open('9SIR MC Tracking').worksheet("2nd Coy RSO/RSI Data")
                    elif coy == "3rd":
                        sheet = client.open('9SIR MC Tracking').worksheet("3rd Coy RSO/RSI Data")
                    while True:
                        if split_text[i+1][0:2] == "MA":
                            break
                        else:
                            if split_text[i+1] != "":
                                eachline = split_text[i+1]
                                if len(eachline) < 10:
                                    camp = eachline.split(':')[0]
                                    i += 1
                                else:
                                    name = eachline.split('(')[0].strip().upper()
                                    vocation = eachline[eachline.find("(")+1:eachline.find(")")]
                                    rso_rsi_type = eachline.split(":")[1][1:4]
                                    rso_rsi_reason = eachline.split("for")[1].strip()
                                    if rso_rsi_type == "RSO":
                                        method = eachline[eachline.find("[")+1:eachline.find("]")].upper()
                                    else:
                                        method = ""
                                    row = [date, name, rso_rsi_type, rso_rsi_reason, camp, vocation, method]
                                    index = 2
                                    sheet.insert_row(row, index, value_input_option='USER_ENTERED')
                                    i += 1
                            else:
                                i += 1

            elif split_text[i][0:2] == "MA":
                if split_text[i][-2:] == "00":
                    i += 2
                    continue
                else:
                    if coy == "1st":
                        sheet = client.open('9SIR MC Tracking').worksheet("1st Coy MA Data")
                    elif coy == "2nd":
                        sheet = client.open('9SIR MC Tracking').worksheet("2nd Coy MA Data")
                    elif coy == "3rd":
                        sheet = client.open('9SIR MC Tracking').worksheet("3rd Coy MA Data")
                    while True:
                        if split_text[i+1][0:8] == "AG+ / C+":
                            break
                        else:
                            if split_text[i+1] != "":
                                eachline = split_text[i+1]
                                if len(eachline) < 10:
                                    camp = eachline.split(':')[0]
                                    i += 1
                                else:
                                    name = eachline.split('(')[0].strip().upper()
                                    vocation = eachline[eachline.find("(")+1:eachline.find(")")]
                                    ma_timing = eachline.split(":")[1][1:6]
                                    ma_location = re.search('@ (.*) for', eachline)
                                    ma_reason = eachline.split("for")[1].strip()
                                    row = [date, name, ma_timing, ma_location.group(1), ma_reason, camp, vocation]
                                    index = 2
                                    sheet.insert_row(row, index, value_input_option='USER_ENTERED')
                                    i += 1
                            else:
                                i += 1

            elif split_text[i][0:8] == "AG+ / C+":
                if split_text[i][-2:] == "00":
                    i += 2
                    continue
                else:
                    if coy == "1st":
                        sheet = client.open('9SIR MC Tracking').worksheet("1st Coy AG+/C+")
                    elif coy == "2nd":
                        sheet = client.open('9SIR MC Tracking').worksheet("2nd Coy AG+/C+")
                    elif coy == "3rd":
                        sheet = client.open('9SIR MC Tracking').worksheet("3rd Coy AG+/C+")
                    while True:
                        if split_text[i+1][0:6] == "Others":
                            break
                        else:
                            if split_text[i+1] != "":
                                eachline = split_text[i+1]
                                if len(eachline) < 10:
                                    camp = eachline.split(':')[0]
                                    i += 1
                                else:
                                    name = eachline.split('(')[0].strip().upper()
                                    vocation = eachline[eachline.find("(")+1:eachline.find(")")]
                                    check_name = eachline.split('(')[0][3:].strip().upper()
                                    check_date = eachline.split('to')[1].strip()
                                    start_date = eachline.split(':')[1][3:5]+"/"+eachline.split(':')[1][1:3]+"/"+eachline.split(':')[1][5:7]
                                    end_date = eachline.split('to')[1][3:5]+"/"+eachline.split('to')[1][1:3]+"/"+eachline.split('to')[1][5:7]
                                    list_all = sheet.get_all_values()
                                    for x in range(len(list_all)):
                                        if list_all[x][1][4:] == check_name and list_all[x][3] == check_date:
                                            break
                                        else:
                                            continue
                                    else:
                                        row = [date, name, start_date, end_date, camp, vocation]
                                        index = 2
                                        sheet.insert_row(row, index, value_input_option='USER_ENTERED')
                                    i += 1
                            else:
                                i += 1

            elif split_text[i][0:6] == "Others":
                if split_text[i][-2:] == "00":
                    i += 2
                    continue
                else:
                    if coy == "1st":
                        sheet = client.open('9SIR MC Tracking').worksheet("1st Coy Others")
                    elif coy == "2nd":
                        sheet = client.open('9SIR MC Tracking').worksheet("2nd Coy Others")
                    elif coy == "3rd":
                        sheet = client.open('9SIR MC Tracking').worksheet("3rd Coy Others")
                    while True:
                        if split_text[i+1][0:8] == "Overseas":
                            break
                        else:
                            if split_text[i+1] != "":
                                eachline = split_text[i+1]
                                if len(eachline) < 10:
                                    camp = eachline.split(':')[0]
                                    i += 1
                                else:
                                    name = eachline.split('(')[0].strip().upper()
                                    vocation = eachline[eachline.find("(")+1:eachline.find(")")]
                                    other_reason = eachline.split(':')[1].strip()
                                    row = [date, name, other_reason, camp, vocation]
                                    index = 2
                                    sheet.insert_row(row, index, value_input_option='USER_ENTERED')
                                    i += 1
                            else:
                                i += 1

            elif split_text[i][0:8] == "Overseas":
                if split_text[i][-2:] == "00":
                    break
                else:
                    if coy == "1st":
                        sheet = client.open('9SIR MC Tracking').worksheet("1st Coy Overseas")
                    elif coy == "2nd":
                        sheet = client.open('9SIR MC Tracking').worksheet("2nd Coy Overseas")
                    elif coy == "3rd":
                        sheet = client.open('9SIR MC Tracking').worksheet("3rd Coy Overseas")
                    while True:
                        if i+1 == len(split_text):
                            break
                        else:
                            if split_text[i+1] != "":
                                eachline = split_text[i+1]
                                if len(eachline) < 10:
                                    camp = eachline.split(':')[0]
                                    i += 1
                                else:
                                    name = eachline.split('(')[0].strip().upper()
                                    vocation = eachline[eachline.find("(")+1:eachline.find(")")]
                                    check_name = eachline.split('(')[0][3:].strip().upper()
                                    check_date = eachline.split('to')[1].strip()
                                    start_date = eachline.split(':')[1][3:5]+"/"+eachline.split(':')[1][1:3]+"/"+eachline.split(':')[1][5:7]
                                    end_date = eachline.split('to')[1][3:5]+"/"+eachline.split('to')[1][1:3]+"/"+eachline.split('to')[1][5:7]
                                    list_all = sheet.get_all_values()
                                    for x in range(len(list_all)):
                                        if list_all[x][1][4:] == check_name and list_all[x][3] == check_date:
                                            break
                                        else:
                                            continue
                                    else:
                                        row = [date, name, start_date, end_date, camp, vocation]
                                        index = 2
                                        sheet.insert_row(row, index, value_input_option='USER_ENTERED')
                                    i += 1
                            else:
                                i += 1
            else:
                if i+1 == len(split_text):
                    break
                else:
                    i += 1
                    continue
        update.message.reply_text("Data added")
        return


def retrievedate(update, context):
    if not getid(update.message.chat.id):
        update.message.reply_text("you are not verified")
        return
    ufd_data = {}
    mc_num = 0
    retrieve_text = update.message.text
    coy = retrieve_text[14:17]
    chosen_date = retrieve_text[-6:]
    check = chosen_date[0:3]
    retrieve_all = False
    if coy == "1st":
        sheet = client.open('9SIR MC Tracking').worksheet("1st Coy MC Data Raw")
    elif coy == "2nd":
        sheet = client.open('9SIR MC Tracking').worksheet("2nd Coy MC Data Raw")
    elif coy == "3rd":
        sheet = client.open('9SIR MC Tracking').worksheet("3rd Coy MC Data Raw")
    else:
        if check == coy:
            retrieve_all = True
        else:
            update.message.reply_text("Company not specified.")
            return

    if retrieve_all:
        mc_details = "All Coys UFD/RSO/MA list for {} \n \n".format(chosen_date)
        sheet1 = client.open('9SIR MC Tracking').worksheet("1st Coy MC Data Raw")
        sheet2 = client.open('9SIR MC Tracking').worksheet("2nd Coy MC Data Raw")
        sheet3 = client.open('9SIR MC Tracking').worksheet("3rd Coy MC Data Raw")
        list1 = sheet1.get_all_values()
        list2 = sheet2.get_all_values()
        list3 = sheet3.get_all_values()
        ufd_data1 = {}
        ufd_data2 = {}
        ufd_data3 = {}
        mc1 = 0
        mc2 = 0
        mc3 = 0
        ufd1 = False
        ufd2 = False
        ufd3 = False
        for i in range(len(list1)):
            if list1[i][0] == chosen_date:
                name = list1[i][1]
                mc_reason = list1[i][2]
                start_date = list1[i][3]
                end_date = list1[i][4]
                camp = list1[i][5]
                if camp in ufd_data1:
                    ufd_data1[camp].append([name, mc_reason, start_date, end_date])
                    mc1 += 1
                else:
                    ufd_data1[camp] = [[name, mc_reason, start_date, end_date]]
                    mc1 += 1
            else:
                continue

        if ufd_data1:
            mc_details = mc_details + "1st Coy UFD: {:02d}".format(mc1)
            for c in range(len(ufd_data1)):
                camp_name = list(ufd_data1.keys())[c]
                mc_details = mc_details + "\n" + str(list(ufd_data1.keys())[c]) + ":\n"
                for v in range(len(ufd_data1[camp_name])):
                    rank_name = ufd_data1[list(ufd_data1.keys())[c]][v][0]
                    mcreason = ufd_data1[list(ufd_data1.keys())[c]][v][1]
                    mcstart = ufd_data1[list(ufd_data1.keys())[c]][v][2]
                    mcend = ufd_data1[list(ufd_data1.keys())[c]][v][3]
                    mc_details = mc_details + "{}: {} to {} for {} \n".format(rank_name, mcstart, mcend, mcreason)
        else:
            ufd1 = True
            update.message.reply_text("*NO* New UFD for 1st Coy on this date", parse_mode='MarkdownV2')

        for i in range(len(list2)):
            if list2[i][0] == chosen_date:
                name = list2[i][1]
                mc_reason = list2[i][2]
                start_date = list2[i][3]
                end_date = list2[i][4]
                camp = list2[i][5]
                if camp in ufd_data2:
                    ufd_data2[camp].append([name, mc_reason, start_date, end_date])
                    mc2 += 1
                else:
                    ufd_data2[camp] = [[name, mc_reason, start_date, end_date]]
                    mc2 += 1
            else:
                continue

        if ufd_data2:
            mc_details = mc_details + "\n\n2nd Coy UFD: {:02d}".format(mc2)
            for c in range(len(ufd_data2)):
                camp_name = list(ufd_data2.keys())[c]
                mc_details = mc_details + "\n" + str(list(ufd_data2.keys())[c]) + ":\n"
                for v in range(len(ufd_data2[camp_name])):
                    rank_name = ufd_data2[list(ufd_data2.keys())[c]][v][0]
                    mcreason = ufd_data2[list(ufd_data2.keys())[c]][v][1]
                    mcstart = ufd_data2[list(ufd_data2.keys())[c]][v][2]
                    mcend = ufd_data2[list(ufd_data2.keys())[c]][v][3]
                    mc_details = mc_details + "{}: {} to {} for {} \n".format(rank_name, mcstart, mcend, mcreason)
        else:
            ufd2 = True
            update.message.reply_text("*NO* New UFD for 2nd Coy on this date", parse_mode='MarkdownV2')

        for i in range(len(list3)):
            if list3[i][0] == chosen_date:
                name = list3[i][1]
                mc_reason = list3[i][2]
                start_date = list3[i][3]
                end_date = list3[i][4]
                camp = list3[i][5]
                if camp in ufd_data3:
                    ufd_data3[camp].append([name, mc_reason, start_date, end_date])
                    mc3 += 1
                else:
                    ufd_data3[camp] = [[name, mc_reason, start_date, end_date]]
                    mc3 += 1
            else:
                continue

        if ufd_data3:
            mc_details = mc_details + "\n\n3rd Coy UFD: {:02d}".format(mc3)
            for c in range(len(ufd_data3)):
                camp_name = list(ufd_data3.keys())[c]
                mc_details = mc_details + "\n" + str(list(ufd_data3.keys())[c]) + ":\n"
                for v in range(len(ufd_data3[camp_name])):
                    rank_name = ufd_data3[list(ufd_data3.keys())[c]][v][0]
                    mcreason = ufd_data3[list(ufd_data3.keys())[c]][v][1]
                    mcstart = ufd_data3[list(ufd_data3.keys())[c]][v][2]
                    mcend = ufd_data3[list(ufd_data3.keys())[c]][v][3]
                    mc_details = mc_details + "{}: {} to {} for {} \n".format(rank_name, mcstart, mcend, mcreason)
        else:
            ufd3 = True
            update.message.reply_text("*NO* New UFD for 3rd Coy on this date", parse_mode='MarkdownV2')
        if ufd1 and ufd2 and ufd3:
            pass
        else:
            update.message.reply_text(mc_details)
    else:
        list_all = sheet.get_all_values()
        mc_details = coy + " Coy UFD/RSO/MA list for {} \n \n".format(chosen_date)
        for i in range(len(list_all)):
            if list_all[i][0] == chosen_date:
                name = list_all[i][1]
                mc_reason = list_all[i][2]
                start_date = list_all[i][3]
                end_date = list_all[i][4]
                camp = list_all[i][5]
                if camp in ufd_data:
                    ufd_data[camp].append([name, mc_reason, start_date, end_date])
                    mc_num += 1
                else:
                    ufd_data[camp] = [[name, mc_reason, start_date, end_date]]
                    mc_num += 1
            else:
                continue

        if ufd_data:
            mc_details = mc_details + "UFD: {:02d}".format(mc_num)
            for c in range(len(ufd_data)):
                camp_name = list(ufd_data.keys())[c]
                mc_details = mc_details + "\n" + str(list(ufd_data.keys())[c]) + ":\n"
                for v in range(len(ufd_data[camp_name])):
                    rank_name = ufd_data[list(ufd_data.keys())[c]][v][0]
                    mcreason = ufd_data[list(ufd_data.keys())[c]][v][1]
                    mcstart = ufd_data[list(ufd_data.keys())[c]][v][2]
                    mcend = ufd_data[list(ufd_data.keys())[c]][v][3]
                    mc_details = mc_details + "{}: {} to {} for {} \n".format(rank_name, mcstart, mcend, mcreason)
            update.message.reply_text(mc_details)
        else:
            update.message.reply_text("*NO* UFD for this date", parse_mode='MarkdownV2')


def retrievemonthyr(update, context):
    if not getid(update.message.chat.id):
        update.message.reply_text("you are not verified")
        return
    ufd_data = {}
    retrieve_text = update.message.text
    coy = retrieve_text[17:20]
    chosen_month = retrieve_text[21:23].strip().zfill(2)
    chosen_year = retrieve_text[23:25].strip()
    year = datetime.strptime(chosen_year, "%y").strftime("%Y")
    month = datetime.strptime(chosen_month, "%m").strftime("%b")
    if coy == "1st":
        sheet = client.open('9SIR MC Tracking').worksheet("1st Coy MC Data Raw")
    elif coy == "2nd":
        sheet = client.open('9SIR MC Tracking').worksheet("2nd Coy MC Data Raw")
    elif coy == "3rd":
        sheet = client.open('9SIR MC Tracking').worksheet("3rd Coy MC Data Raw")
    else:
        update.message.reply_text("Company not specified or wrong.")
        return
    list_all = sheet.get_all_values()
    mc_num = 0
    mc_details = coy + " Coy UFD/RSO/MA list for {} {} \n \n".format(month, year)
    for i in range(len(list_all)):
        if list_all[i][0][2:6] == chosen_month+chosen_year:
            name = list_all[i][1]
            mc_reason = list_all[i][2]
            start_date = list_all[i][3]
            end_date = list_all[i][4]
            camp = list_all[i][5]
            if camp in ufd_data:
                ufd_data[camp].append([name, mc_reason, start_date, end_date])
                mc_num += 1
            else:
                ufd_data[camp] = [[name, mc_reason, start_date, end_date]]
                mc_num += 1
        else:
            continue

    if ufd_data:
        mc_details = mc_details + "UFD: {:02d}".format(mc_num)
        for c in range(len(ufd_data)):
            camp_name = list(ufd_data.keys())[c]
            mc_details = mc_details + "\n" + str(list(ufd_data.keys())[c]) + ":\n"
            for v in range(len(ufd_data[camp_name])):
                rank_name = ufd_data[list(ufd_data.keys())[c]][v][0]
                mcreason = ufd_data[list(ufd_data.keys())[c]][v][1]
                mcstart = ufd_data[list(ufd_data.keys())[c]][v][2]
                mcend = ufd_data[list(ufd_data.keys())[c]][v][3]
                mc_details = mc_details + "{}: {} to {} for {} \n".format(rank_name, mcstart, mcend, mcreason)
        update.message.reply_text(mc_details)
    else:
        update.message.reply_text("*NO* UFD for " + coy + " Coy on this month and year", parse_mode='MarkdownV2')


def retrievecamp(update, context):
    if not getid(update.message.chat.id):
        update.message.reply_text("you are not verified")
        return
    mc_details = ""
    retrieve_text = update.message.text
    campchosen = retrieve_text[17:].strip().upper()
    coy = retrieve_text[13:17].strip().lower()
    month_year = retrieve_text[-4:].strip()
    if month_year.isdigit():
        size = len(campchosen)
        chosen_camp = campchosen[:size-4].strip()
    else:
        chosen_camp = campchosen

    if campchosen != "":
        if chosen_camp == "DTTB" or "PEH" or "CLEMENTI" or "RRRC" or "MAJU" or "KC" or "KHC" or "MOWBRAY" or "5AD1" or "5AD2" or "SGC":
            coy = "1st"
            sheet = client.open('9SIR MC Tracking').worksheet("1st Coy MC Data Raw")
        elif chosen_camp == "CCK" or "LCK" or "SC" or "MHC" or "MWC" or "MC2" or "JC1" or "JC2" or "SAFTI MI" or "PLAD" or "PLC" or "TNB":
            coy = "2nd"
            sheet = client.open('9SIR MC Tracking').worksheet("2nd Coy MC Data Raw")
        elif chosen_camp == "JGL" or "ROT2" or "ROT3":
            coy = "3rd"
            sheet = client.open('9SIR MC Tracking').worksheet("3rd Coy MC Data Raw")
        else:
            update.message.reply_text("Chosen camp does not exist in any of the companies")

        if month_year.isnumeric():
            print("month yr")
            month_yr = datetime.strptime(month_year, "%m%y").strftime("%b"+" "+"%Y")
            list_all = sheet.get_all_values()
            mc_num = 0
            name_list = []
            date_dict = {}
            mc_detail = coy + " Coy UFD/RSO/MA list for {} in {} \n \n".format(chosen_camp, month_yr)
            for i in range(len(list_all)):
                if list_all[i][5] == chosen_camp:
                    name = list_all[i][1]
                    mc_reason = list_all[i][2]
                    start_date = list_all[i][3]
                    end_date = list_all[i][4]
                    start = datetime.strptime(start_date, "%d%m%y")
                    end = datetime.strptime(end_date, "%d%m%y")
                    range1 = datetime.strptime(month_year, "%m%y")
                    range2 = datetime.strptime(month_year, "%m%y") + relativedelta(months=1) - relativedelta(days=1)
                    dayss = " "
                    mc_pass = True
                    if range1 <= start <= range2 and end <= range2:
                        dayss = (end-start).days + 1
                    elif range1 <= start <= range2 <= end:
                        dayss = (range2-start).days + 1
                    elif start <= range1 <= end <= range2:
                        dayss = (end-range1).days + 1
                    elif start <= range1 and end >= range2:
                        dayss = (range2-range1).days + 1
                    else:
                        mc_pass = False

                    if mc_pass is True:
                        name_list.append(name)
                        if name in date_dict:
                            if dayss == 0:
                                dayss = 1
                            else:
                                dayss = dayss
                            date_dict[name] = date_dict[name]+dayss
                        else:
                            if dayss == 0:
                                dayss = 1
                            else:
                                dayss = dayss
                            date_dict[name] = dayss
                        mc_details = mc_details + "{}: {} to {} for {} \n\n".format(name, start_date, end_date, mc_reason)
                        mc_num += 1
                else:
                    continue

            if mc_details != "":
                most_num_ufd = Counter(name_list).most_common()
                sorted_dates = sorted(date_dict.items(), key=lambda date_dict: date_dict[1], reverse=True)
                mc_detail += "Top 3 Number of MCs taken is:\n"
                for i in range(len(most_num_ufd)):
                    if i != 3:
                        mc_detail += "\n{} with {:02d} MCs taken".format(most_num_ufd[i][0], most_num_ufd[i][1])
                    else:
                        break
                mc_detail += "\n\nTop 3 Days on MC is:\n"
                for x in range(len(sorted_dates)):
                    if x != 3:
                        mc_detail += "\n{} with {:02d} days on MC".format(sorted_dates[x][0], sorted_dates[x][1])
                    else:
                        break
                mc_detail += "\n\nUFD: {:02d}\n\n".format(mc_num)
                mc_detail += mc_details
                update.message.reply_text(mc_detail)
            else:
                update.message.reply_text("*NO* UFD for {} yet".format(chosen_camp), parse_mode='MarkdownV2')

        else:
            list_all = sheet.get_all_values()
            mc_num = 0
            name_list = []
            date_dict = {}
            mc_detail = coy + " Coy UFD/RSO/MA list for {} \n \n".format(chosen_camp)
            for i in range(len(list_all)):
                if list_all[i][5] == chosen_camp:
                    name = list_all[i][1]
                    name_list.append(name)
                    mc_reason = list_all[i][2]
                    start_date = list_all[i][3]
                    end_date = list_all[i][4]
                    start = datetime.strptime(start_date, "%d%m%y")
                    end = datetime.strptime(end_date, "%d%m%y")
                    days = (end - start).days
                    if name in date_dict:
                        if days == 0:
                            days = 1
                        else:
                            days = days
                        date_dict[name] = date_dict[name]+days
                    else:
                        if days == 0:
                            days = 1
                        else:
                            days = days
                        date_dict[name] = days
                    mc_details = mc_details + "{}: {} to {} for {} \n\n".format(name, start_date, end_date, mc_reason)
                    mc_num += 1
                else:
                    continue

            if mc_details != "":
                most_num_ufd = Counter(name_list).most_common()
                sorted_dates = sorted(date_dict.items(), key=lambda date_dict: date_dict[1], reverse=True)
                mc_detail += "Top 3 Number of MCs taken is:\n"
                for i in range(len(most_num_ufd)):
                    if i != 3:
                        mc_detail += "\n{} with {:02d} MCs taken".format(most_num_ufd[i][0], most_num_ufd[i][1])
                    else:
                        break
                mc_detail += "\n\nTop 3 Days on MC is:\n"
                for x in range(len(sorted_dates)):
                    if x != 3:
                        mc_detail += "\n{} with {:02d} days on MC".format(sorted_dates[x][0], sorted_dates[x][1])
                    else:
                        break
                mc_detail += "\n\nUFD: {:02d}\n\n".format(mc_num)
                mc_detail += mc_details
                update.message.reply_text(mc_detail)
            else:
                update.message.reply_text("*NO* UFD for {} yet".format(chosen_camp), parse_mode='MarkdownV2')

    else:
        sheet = client.open('9SIR MC Tracking').worksheet("Camps")
        list_all = sheet.get_all_values()
        name_list = []
        name_list2 = []
        name_list3 = []
        for i in range(1, len(list_all)):
            name = list_all[i][0]
            name2 = list_all[i][1]
            name3 = list_all[i][2]
            if name in name_list:
                continue
            else:
                name_list.append(name)
            if name2 in name_list2:
                continue
            else:
                name_list2.append(name2)
            if name3 in name_list3:
                continue
            else:
                name_list3.append(name3)

        if name_list:
            mc_details = "Camps under 1st Coy: \n\n"
            for x in range(len(name_list)):
                mc_details += "{}\n\n".format(name_list[x])
            update.message.reply_text(mc_details)
        if name_list2:
            mc_details = "Camps under 2nd Coy: \n\n"
            for x in range(len(name_list2)):
                mc_details += "{}\n\n".format(name_list2[x])
            update.message.reply_text(mc_details)
        if name_list3:
            mc_details = "Camps under 3rd Coy: \n\n"
            for x in range(len(name_list3)):
                mc_details += "{}\n\n".format(name_list3[x])
            update.message.reply_text(mc_details)


def retrievename(update, context):
    if not getid(update.message.chat.id):
        update.message.reply_text("you are not verified")
        return
    retrieve_text = update.message.text
    coy = retrieve_text[14:17]
    chosen_name = retrieve_text[17:].strip().upper()
    getallnames = False
    if coy == "1st":
        sheet = client.open('9SIR MC Tracking').worksheet("1st Coy MC Data Raw")
    elif coy == "2nd":
        sheet = client.open('9SIR MC Tracking').worksheet("2nd Coy MC Data Raw")
    elif coy == "3rd":
        sheet = client.open('9SIR MC Tracking').worksheet("3rd Coy MC Data Raw")
    else:
        if coy == "" and chosen_name == "":
            getallnames = True
        else:
            update.message.reply_text("Company not specified.")
    if chosen_name != "":
        list_all = sheet.get_all_values()
        ufd_list = {}
        days_dict = {}
        mc_details = ""
        mc_num = 0
        for i in range(1, len(list_all)):
            name = list_all[i][7]
            if chosen_name == name:
                rank_name = list_all[i][1]
                ufd_name = list_all[i][1][4:]
                submitted_date = list_all[i][0]
                mc_reason = list_all[i][2]
                start_date = list_all[i][3]
                end_date = list_all[i][4]
                start = datetime.strptime(start_date, "%d%m%y")
                end = datetime.strptime(end_date, "%d%m%y")
                days = (end - start).days + 1
                mc_num += 1
                if ufd_name in ufd_list:
                    ufd_list[ufd_name].append([submitted_date, mc_reason, start_date, end_date])
                else:
                    ufd_list[ufd_name] = [[submitted_date, mc_reason, start_date, end_date]]
                if ufd_name in days_dict:
                    if days == 0:
                        days = 1
                    else:
                        days = days
                    days_dict[ufd_name] = days_dict[ufd_name]+days
                else:
                    if days == 0:
                        days = 1
                    else:
                        days = days
                    days_dict[ufd_name] = days
            else:
                continue

        if ufd_list:
            # how many days on mc and number of mcs
            mc_details += coy + " Coy UFD list for {} \n\n".format(rank_name)
            mc_details += "Number of MCs taken: {} \n\n".format(mc_num)
            mc_details += "Total number of days on MC: {} \n\n".format(days_dict[ufd_name])
            for x in range(len(ufd_list[ufd_name])):
                mc_details += "{}: MC for {} from {} to {} \n".format(ufd_list[ufd_name][x][0], ufd_list[ufd_name][x][1], ufd_list[ufd_name][x][2], ufd_list[ufd_name][x][3])
        else:
            mc_details += "{} does not have MC".format(chosen_name)
        update.message.reply_text(mc_details)

    else:
        if not getallnames:
            list_all = sheet.get_all_values()
            name_list = []
            for i in range(1, len(list_all)):
                name = list_all[i][7]
                if name in name_list:
                    continue
                else:
                    name_list.append(name)

            if name_list:
                mc_details = "UFD Personnel Name List for " + coy + " Coy: \n\n"
                sorted_list = sorted(name_list)
                for x in range(len(sorted_list)):
                    mc_details += "{}\n\n".format(sorted_list[x])
                update.message.reply_text(mc_details)
        else:
            sheet1 = client.open('9SIR MC Tracking').worksheet("1st Coy MC Data Raw")
            sheet2 = client.open('9SIR MC Tracking').worksheet("2nd Coy MC Data Raw")
            sheet3 = client.open('9SIR MC Tracking').worksheet("3rd Coy MC Data Raw")
            name_list1 = []
            name_list2 = []
            name_list3 = []
            list_all1 = sheet1.get_all_values()
            list_all2 = sheet2.get_all_values()
            list_all3 = sheet3.get_all_values()
            for i in range(1, len(list_all1)):
                name = list_all1[i][7]
                if name in name_list1:
                    continue
                else:
                    name_list1.append(name)

            if name_list1:
                mc_details = "UFD Personnel Name List for 1st Coy: \n\n"
                sorted_list1 = sorted(name_list1)
                for x in range(len(sorted_list1)):
                    mc_details += "{}\n\n".format(sorted_list1[x])
                update.message.reply_text(mc_details)

            for i in range(1, len(list_all2)):
                name = list_all2[i][7]
                if name in name_list2:
                    continue
                else:
                    name_list2.append(name)

            if name_list2:
                mc_details = "UFD Personnel Name List for 2nd Coy: \n\n"
                sorted_list2 = sorted(name_list2)
                for x in range(len(sorted_list2)):
                    mc_details += "{}\n\n".format(sorted_list2[x])
                update.message.reply_text(mc_details)

            for i in range(1, len(list_all3)):
                name = list_all3[i][7]
                if name in name_list3:
                    continue
                else:
                    name_list3.append(name)

            if name_list3:
                mc_details = "UFD Personnel Name List for 3rd Coy: \n\n"
                sorted_list3 = sorted(name_list3)
                for x in range(len(sorted_list3)):
                    mc_details += "{}\n\n".format(sorted_list3[x])
                update.message.reply_text(mc_details)


def mostmcdays(update, context):
    if not getid(update.message.chat.id):
        update.message.reply_text("you are not verified")
        return
    get_by_range = False
    get_by_month = False
    retrieve_text = update.message.text
    coy = retrieve_text[12:15]
    chosen_month = retrieve_text[16:18].strip()
    chosen_year = retrieve_text[18:20].strip()
    chosen_othermonth = retrieve_text[24:26].strip()
    chosen_otheryear = retrieve_text[26:28].strip()
    if coy == "1st":
        sheet = client.open('9SIR MC Tracking').worksheet("1st Coy MC Data Raw")
    elif coy == "2nd":
        sheet = client.open('9SIR MC Tracking').worksheet("2nd Coy MC Data Raw")
    elif coy == "3rd":
        sheet = client.open('9SIR MC Tracking').worksheet("3rd Coy MC Data Raw")
    else:
        update.message.reply_text("Company not specified.")
        return

    if chosen_year != "" and chosen_month != "" and chosen_othermonth != "" and chosen_otheryear != "":
        year = datetime.strptime(chosen_year, "%y").strftime("%Y")
        month = datetime.strptime(chosen_month, "%m").strftime("%b")
        other_year = datetime.strptime(chosen_otheryear, "%y").strftime("%Y")
        other_month = datetime.strptime(chosen_othermonth, "%m").strftime("%b")
        get_by_range = True
    elif chosen_year != "" and chosen_month != "":
        year = datetime.strptime(chosen_year, "%y").strftime("%Y")
        month = datetime.strptime(chosen_month, "%m").strftime("%b")
        get_by_month = True
    else:
        get_by_range = False
        get_by_month = False

    mc_details = coy + " Coy UFD list sorted by most number of days"
    list_all = sheet.get_all_values()
    if get_by_month is False and get_by_range is False:
        date_dict = {}
        for i in range(1, len(list_all)):
            name = list_all[i][1]
            start_date = list_all[i][3]
            end_date = list_all[i][4]
            camp = list_all[i][5]
            start = datetime.strptime(start_date, "%d%m%y")
            end = datetime.strptime(end_date, "%d%m%y")
            days = (end - start).days + 1
            if name in date_dict:
                if days == 0:
                    days = 1
                else:
                    days = days
                date_dict[name] = date_dict[name][0]+days, camp
            else:
                if days == 0:
                    days = 1
                else:
                    days = days
                date_dict[name] = days, camp

        if date_dict:
            sorted_dates = sorted(date_dict.items(), key=lambda date_dict: date_dict[1], reverse=True)
            for x in range(len(sorted_dates)):
                mc_details += "\n \n{} with {} days on MC from {}".format(sorted_dates[x][0], sorted_dates[x][1][0], sorted_dates[x][1][1])
                if len(mc_details) > 3900:
                    update.message.reply_text(mc_details)
                    mc_details = ""
        else:
            update.message.reply_text("*NO* one is on UFD", parse_mode='MarkdownV2')
            return
        update.message.reply_text(mc_details)

    elif get_by_month is True:
        mc_details += " for {} {}".format(month, year)
        date_dict = {}
        for i in range(1, len(list_all)):
            name = list_all[i][1]
            start_date = list_all[i][3]
            end_date = list_all[i][4]
            camp = list_all[i][5]
            start = datetime.strptime(start_date, "%d%m%y")
            end = datetime.strptime(end_date, "%d%m%y")
            range1 = datetime.strptime(chosen_month+chosen_year, "%m%y")  # start of month
            range2 = datetime.strptime(chosen_month+chosen_year, "%m%y") + relativedelta(months=1) - relativedelta(days=1)  # end of month
            dayss = " "
            mc_pass = True
            if range1 <= start <= range2 and end <= range2:
                dayss = (end-start).days + 1
            elif range1 <= start <= range2 <= end:
                dayss = (range2-start).days + 1
            elif start <= range1 <= end <= range2:
                dayss = (end-range1).days + 1
            elif start <= range1 and end >= range2:
                dayss = (range2-range1).days + 1
            else:
                mc_pass = False

            if mc_pass is True:
                if name in date_dict:
                    date_dict[name] = date_dict[name][0]+dayss, camp
                else:
                    date_dict[name] = dayss, camp

        if date_dict:
            sorted_dates = sorted(date_dict.items(), key=lambda date_dict: date_dict[1], reverse=True)
            for x in range(len(sorted_dates)):
                mc_details += "\n \n{} with {} days on MC from {}".format(sorted_dates[x][0], sorted_dates[x][1][0], sorted_dates[x][1][1])
                if len(mc_details) > 3900:
                    update.message.reply_text(mc_details)
                    mc_details = ""
        else:
            update.message.reply_text("*NO* one is on UFD for the chosen month", parse_mode='MarkdownV2')
            return
        update.message.reply_text(mc_details)

    elif get_by_range is True:
        mc_details += " for {} {} to {} {}".format(month, year, other_month, other_year)
        date_dict = {}
        for i in range(1, len(list_all)):
            name = list_all[i][1]
            start_date = list_all[i][3]
            end_date = list_all[i][4]
            camp = list_all[i][5]
            start = datetime.strptime(start_date, "%d%m%y")
            end = datetime.strptime(end_date, "%d%m%y")
            range1 = datetime.strptime(chosen_month+chosen_year, "%m%y")  # start of month
            range2 = datetime.strptime(chosen_othermonth+chosen_otheryear, "%m%y") + relativedelta(months=1) - relativedelta(days=1) # end of month
            dayss = " "
            mc_pass = True
            if range1 <= start <= range2 and end <= range2:
                dayss = (end-start).days + 1
            elif range1 <= start <= range2 <= end:
                dayss = (range2-start).days + 1
            elif start <= range1 <= end <= range2:
                dayss = (end-range1).days + 1
            elif start <= range1 and end >= range2:
                dayss = (range2-range1).days + 1
            else:
                mc_pass = False

            if mc_pass is True:
                if name in date_dict:
                    date_dict[name] = date_dict[name][0]+dayss, camp
                else:
                    date_dict[name] = dayss, camp

        if date_dict:
            sorted_dates = sorted(date_dict.items(), key=lambda date_dict: date_dict[1], reverse=True)
            for x in range(len(sorted_dates)):
                mc_details += "\n \n{} with {} days on MC from {}".format(sorted_dates[x][0], sorted_dates[x][1][0], sorted_dates[x][1][1])
                if len(mc_details) > 3900:
                    update.message.reply_text(mc_details)
                    mc_details = ""
        else:
            update.message.reply_text("*NO* one is on UFD for the chosen date range", parse_mode='MarkdownV2')
            return
        update.message.reply_text(mc_details)


def mostmc(update, context):
    if not getid(update.message.chat.id):
        update.message.reply_text("you are not verified")
        return
    mc_details = "1st Coy UFD list sorted by most number of MCs taken"
    retrieve_text = update.message.text
    get_by_range = False
    get_by_month = False
    coy = retrieve_text[8:11]
    chosen_month = retrieve_text[12:14].strip()
    chosen_year = retrieve_text[14:16].strip()
    chosen_othermonth = retrieve_text[20:22].strip()
    chosen_otheryear = retrieve_text[22:24].strip()

    if coy == "1st":
        sheet = client.open('9SIR MC Tracking').worksheet("1st Coy MC Data Raw")
    elif coy == "2nd":
        sheet = client.open('9SIR MC Tracking').worksheet("2nd Coy MC Data Raw")
    elif coy == "3rd":
        sheet = client.open('9SIR MC Tracking').worksheet("3rd Coy MC Data Raw")
    else:
        update.message.reply_text("Company not specified.")
        return

    if chosen_year != "" and chosen_month != "" and chosen_othermonth != "" and chosen_otheryear != "":
        year = datetime.strptime(chosen_year, "%y").strftime("%Y")
        month = datetime.strptime(chosen_month, "%m").strftime("%b")
        other_year = datetime.strptime(chosen_otheryear, "%y").strftime("%Y")
        other_month = datetime.strptime(chosen_othermonth, "%m").strftime("%b")
        get_by_range = True
    elif chosen_year != "" and chosen_month != "":
        year = datetime.strptime(chosen_year, "%y").strftime("%Y")
        month = datetime.strptime(chosen_month, "%m").strftime("%b")
        get_by_month = True
    else:
        get_by_range = False
        get_by_month = False

    list_all = sheet.get_all_values()
    if get_by_month is False and get_by_range is False:
        name_dict = {}
        for i in range(1, len(list_all)):
            name = list_all[i][1][4:]
            camp = list_all[i][5]
            if name in name_dict:
                name_dict[name] = name_dict[name][0] + 1, camp
            else:
                name_dict[name] = 1, camp

        if name_dict:
            most_num_ufd = sorted(name_dict.items(), key=lambda name_dict: name_dict[1], reverse=True)
            for i in range(len(most_num_ufd)):
                mc_details += "\n\n{} with {:02d} MCs taken from {}".format(most_num_ufd[i][0], most_num_ufd[i][1][0], most_num_ufd[i][1][1])
                if len(mc_details) > 3900:
                    update.message.reply_text(mc_details)
                    mc_details = ""
        else:
            update.message.reply_text("*NO* one is on UFD", parse_mode='MarkdownV2')
        update.message.reply_text(mc_details)

    elif get_by_month is True:
        name_dict = {}
        for i in range(1, len(list_all)):
            name = list_all[i][1][4:]
            start_date = list_all[i][3]
            end_date = list_all[i][4]
            camp = list_all[i][5]
            between = datetime.strptime(chosen_month+chosen_year, "%m%y")
            start = datetime.strptime(start_date, "%d%m%y")
            end = datetime.strptime(end_date, "%d%m%y")
            if start < between < end or start_date[2:6] == chosen_month+chosen_year or end_date[2:6] == chosen_month+chosen_year:
                if name in name_dict:
                    name_dict[name] = name_dict[name][0] + 1, camp
                else:
                    name_dict[name] = 1, camp
            else:
                continue

        if name_dict:
            most_num_ufd = sorted(name_dict.items(), key=lambda name_dict: name_dict[1], reverse=True)
            mc_details += " for {} {}".format(month, year)
            for i in range(len(most_num_ufd)):
                mc_details += "\n\n{} with {:02d} MCs taken from {}".format(most_num_ufd[i][0], most_num_ufd[i][1][0], most_num_ufd[i][1][1])
                if len(mc_details) > 3900:
                    update.message.reply_text(mc_details)
                    mc_details = ""
        else:
            update.message.reply_text("*NO* one is on UFD for the chosen month", parse_mode='MarkdownV2')
            return
        update.message.reply_text(mc_details)

    elif get_by_range is True:
        name_dict = {}
        for i in range(1, len(list_all)):
            name = list_all[i][1][4:]
            start_date = list_all[i][3]
            end_date = list_all[i][4]
            camp = list_all[i][5]
            start = datetime.strptime(start_date, "%d%m%y")
            end = datetime.strptime(end_date, "%d%m%y")
            range1 = datetime.strptime(chosen_month+chosen_year, "%m%y")
            range2 = datetime.strptime(chosen_othermonth+chosen_otheryear, "%m%y") + relativedelta(months=1) - relativedelta(days=1)
            if start <= range1 <= end or start <= range2 <= end or range1 <= start <= range2 or range1 <= end <= range2:
                if name in name_dict:
                    name_dict[name] = name_dict[name][0] + 1, camp
                else:
                    name_dict[name] = 1, camp
            else:
                continue

        if name_dict:
            most_num_ufd = sorted(name_dict.items(), key=lambda name_dict: name_dict[1], reverse=True)
            mc_details += " for {} {} to {} {}".format(month, year, other_month, other_year)
            for i in range(len(most_num_ufd)):
                mc_details += "\n\n{} with {:02d} MCs taken from {}".format(most_num_ufd[i][0], most_num_ufd[i][1][0], most_num_ufd[i][1][1])
                if len(mc_details) > 3900:
                    update.message.reply_text(mc_details)
                    mc_details = ""
        else:
            update.message.reply_text("*NO* one is on UFD for the chosen date range", parse_mode='MarkdownV2')
            return
        update.message.reply_text(mc_details)


def activemc(update, context):
    if not getid(update.message.chat.id):
        update.message.reply_text("you are not verified")
        return
    retrieve_text = update.message.text
    coy = retrieve_text[10:].strip()
    all_mc = False
    if coy == "1st":
        sheet = client.open('9SIR MC Tracking').worksheet("1st Coy MC Data Raw")
    elif coy == "2nd":
        sheet = client.open('9SIR MC Tracking').worksheet("2nd Coy MC Data Raw")
    elif coy == "3rd":
        sheet = client.open('9SIR MC Tracking').worksheet("3rd Coy MC Data Raw")
    else:
        if coy == "":
            all_mc = True
        else:
            update.message.reply_text("Company not specified.")
    if all_mc:
        sheet1 = client.open('9SIR MC Tracking').worksheet("1st Coy MC Data Raw")
        sheet2 = client.open('9SIR MC Tracking').worksheet("2nd Coy MC Data Raw")
        sheet3 = client.open('9SIR MC Tracking').worksheet("3rd Coy MC Data Raw")
        list_all1 = sheet1.get_all_values()
        list_all2 = sheet2.get_all_values()
        list_all3 = sheet3.get_all_values()
        active_list1 = {}
        active_list2 = {}
        active_list3 = {}
        mc_num1 = 0
        mc_num2 = 0
        mc_num3 = 0
        now = datetime.today().date()
        for i in range(1, len(list_all1)):
            mc_end = datetime.strptime(list_all1[i][4], "%d%m%y").date()
            if mc_end >= now:
                camp = list_all1[i][5]
                start_date = list_all1[i][3]
                end_date = list_all1[i][4]
                reason = list_all1[i][2]
                name = list_all1[i][1]
                if camp in active_list1:
                    active_list1[camp].append([name, reason, start_date, end_date])
                    mc_num1 += 1
                else:
                    active_list1[camp] = [[name, reason, start_date, end_date]]
                    mc_num1 += 1
            else:
                continue

        for i in range(1, len(list_all2)):
            mc_end = datetime.strptime(list_all2[i][4], "%d%m%y").date()
            if mc_end >= now:
                camp = list_all2[i][5]
                start_date = list_all2[i][3]
                end_date = list_all2[i][4]
                reason = list_all2[i][2]
                name = list_all2[i][1]
                if camp in active_list2:
                    active_list2[camp].append([name, reason, start_date, end_date])
                    mc_num2 += 1
                else:
                    active_list2[camp] = [[name, reason, start_date, end_date]]
                    mc_num2 += 1
            else:
                continue

        for i in range(1, len(list_all3)):
            mc_end = datetime.strptime(list_all3[i][4], "%d%m%y").date()
            if mc_end >= now:
                camp = list_all3[i][5]
                start_date = list_all3[i][3]
                end_date = list_all3[i][4]
                reason = list_all3[i][2]
                name = list_all3[i][1]
                if camp in active_list3:
                    active_list3[camp].append([name, reason, start_date, end_date])
                    mc_num3 += 1
                else:
                    active_list3[camp] = [[name, reason, start_date, end_date]]
                    mc_num3 += 1
            else:
                continue

        if active_list1:
            mc_details = "1st Coy active UFD List"
            mc_details += "\n\nUFD: {:02d}\n".format(mc_num1)
            for c in range(len(active_list1)):
                camp_name = list(active_list1.keys())[c]
                mc_details = mc_details + "\n" + str(list(active_list1.keys())[c]) + ":"
                for v in range(len(active_list1[camp_name])):
                    rank_name = active_list1[list(active_list1.keys())[c]][v][0]
                    mcreason = active_list1[list(active_list1.keys())[c]][v][1]
                    mcstart = active_list1[list(active_list1.keys())[c]][v][2]
                    mcend = active_list1[list(active_list1.keys())[c]][v][3]
                    mc_details = mc_details + "\n{}: {} to {} for {} \n".format(rank_name, mcstart, mcend, mcreason)
            update.message.reply_text(mc_details)
        else:
            update.message.reply_text("*NO* active UFD currently for 1st Coy", parse_mode='MarkdownV2')

        if active_list2:
            mc_details = "2nd Coy active UFD List"
            mc_details += "\n\nUFD: {:02d}\n".format(mc_num2)
            for c in range(len(active_list2)):
                camp_name = list(active_list2.keys())[c]
                mc_details = mc_details + "\n" + str(list(active_list2.keys())[c]) + ":"
                for v in range(len(active_list2[camp_name])):
                    rank_name = active_list2[list(active_list2.keys())[c]][v][0]
                    mcreason = active_list2[list(active_list2.keys())[c]][v][1]
                    mcstart = active_list2[list(active_list2.keys())[c]][v][2]
                    mcend = active_list2[list(active_list2.keys())[c]][v][3]
                    mc_details = mc_details + "\n{}: {} to {} for {} \n".format(rank_name, mcstart, mcend, mcreason)
            update.message.reply_text(mc_details)
        else:
            update.message.reply_text("*NO* active UFD currently for 2nd Coy", parse_mode='MarkdownV2')

        if active_list3:
            mc_details = "3rd Coy active UFD List"
            mc_details += "\n\nUFD: {:02d}\n".format(mc_num3)
            for c in range(len(active_list3)):
                camp_name = list(active_list3.keys())[c]
                mc_details = mc_details + "\n" + str(list(active_list3.keys())[c]) + ":"
                for v in range(len(active_list3[camp_name])):
                    rank_name = active_list3[list(active_list3.keys())[c]][v][0]
                    mcreason = active_list3[list(active_list3.keys())[c]][v][1]
                    mcstart = active_list3[list(active_list3.keys())[c]][v][2]
                    mcend = active_list3[list(active_list3.keys())[c]][v][3]
                    mc_details = mc_details + "\n{}: {} to {} for {} \n".format(rank_name, mcstart, mcend, mcreason)
            update.message.reply_text(mc_details)
        else:
            update.message.reply_text("*NO* active UFD currently for 3rd Coy", parse_mode='MarkdownV2')

    else:
        mc_details = coy + " Coy active UFD list"
        list_all = sheet.get_all_values()
        now = datetime.today().date()
        active_list = {}
        mc_num = 0
        for i in range(1, len(list_all)):
            mc_end = datetime.strptime(list_all[i][4], "%d%m%y").date()
            if mc_end >= now:
                camp = list_all[i][5]
                start_date = list_all[i][3]
                end_date = list_all[i][4]
                reason = list_all[i][2]
                name = list_all[i][1]
                if camp in active_list:
                    active_list[camp].append([name, reason, start_date, end_date])
                    mc_num += 1
                else:
                    active_list[camp] = [[name, reason, start_date, end_date]]
                    mc_num += 1
            else:
                continue

        if active_list:
            mc_details += "\n\nUFD: {:02d}\n".format(mc_num)
            for c in range(len(active_list)):
                camp_name = list(active_list.keys())[c]
                mc_details = mc_details + "\n" + str(list(active_list.keys())[c]) + ":"
                for v in range(len(active_list[camp_name])):
                    rank_name = active_list[list(active_list.keys())[c]][v][0]
                    mcreason = active_list[list(active_list.keys())[c]][v][1]
                    mcstart = active_list[list(active_list.keys())[c]][v][2]
                    mcend = active_list[list(active_list.keys())[c]][v][3]
                    mc_details = mc_details + "\n{}: {} to {} for {} \n".format(rank_name, mcstart, mcend, mcreason)
            update.message.reply_text(mc_details)
        else:
            update.message.reply_text("*NO* active UFD currently", parse_mode='MarkdownV2')


def help(update, context):
    if not getid(update.message.chat.id):
        update.message.reply_text("you are not verified")
        return
    update.message.reply_text("Please click the following link for the user manual guide.")
    update.message.reply_text("https://docs.google.com/document/d/1zg9gyZZyrKFaIhj7Zla7_IR-94b-kqE5pMy0BBVem_8/edit?usp=sharing")


def error(update, context):
    """Log Errors caused by Updates."""
    update.message.reply_text("An error may have happened while inserting data, please check for formatting error")
    logger.warning('Update "%s" caused error "%s"', update, context.error)


def main():
    """Start the bot."""
    # Create the Updater and pass it your bot's token.
    # Make sure to set use_context=True to use the new context based callbacks
    # Post version 12 this will no longer be necessary
    updater = Updater(TOKEN, use_context=True)
    # Get the dispatcher to register handlers
    dp = updater.dispatcher
    # on different commands - answer in Telegram
    dp.add_handler(CommandHandler("start", start))
    dp.add_handler(CommandHandler("retrievedate", retrievedate))
    dp.add_handler(CommandHandler("retrievename", retrievename))
    dp.add_handler(CommandHandler("retrievemonthyr", retrievemonthyr))
    dp.add_handler(CommandHandler("mostmcdays", mostmcdays))
    dp.add_handler(CommandHandler("retrievecamp", retrievecamp))
    dp.add_handler(CommandHandler("mostmc", mostmc))
    dp.add_handler(CommandHandler("activemc", activemc))
    dp.add_handler(CommandHandler("help", help))
    dp.add_handler(MessageHandler(Filters.text, gettext))
    # log all errors
    dp.add_error_handler(error)
    updater.start_polling()
    updater.idle()


if __name__ == '__main__':
    main()
