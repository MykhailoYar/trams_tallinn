import re
import time

import xlsxwriter
from selenium import webdriver
from selenium.webdriver.common.by import By

# basic url for each route
route_urls = [f"https://transport.tallinn.ee/#tram/{route}/a-b" for route in range(1, 6)]

# file for saving tram data
workbook = xlsxwriter.Workbook('/Users/mykhailoyaroshenko/Desktop/trams_tallinn/test.xlsx')

# constants
format_num = workbook.add_format({'num_format': '0.000'})
format_time = workbook.add_format({'num_format': 'hh:mm'})

all_stations = dict()
for ur in route_urls:

    # open basic url
    driver = webdriver.Chrome()
    driver.get(ur)
    driver.minimize_window()

    # loading the page
    time.sleep(4)

    # name of the direction from left and right side
    element2 = driver.find_element("id", "divScheduleHeader")
    direction = element2.text
    directions = re.split('\n', direction)

    # all station from left
    element2 = driver.find_element("id", "divScheduleLeft")
    stops_left = element2.text
    stops_left_list = re.split('\n', stops_left)

    # all station from right
    element2 = driver.find_element("id", "divScheduleRight")
    stops_right = element2.text
    stops_right_list = re.split('\n', stops_right)

    stops_all = stops_left_list + stops_right_list
    list_set, unique_list = set(stops_all), (list(stops_all))
    dir_left, dir_right = {'Direction': directions[0]}, {'Direction': directions[1]}

    for station in unique_list:
        continue_link2 = driver.find_elements(By.LINK_TEXT, station)
        for cl in continue_link2:
            link = str(cl.get_attribute('href'))
            if link:
                if 'a-b' in link:
                    dir_left[station] = link
                else:
                    dir_right[station] = link

    driver.close()

    for i in unique_list:
            l = []
            if i in dir_left.keys():
                l.append(dir_left[i])
            else:
                l.append('')
            if i in dir_right.keys():
                l.append(dir_right[i])
            else:
                l.append('')

            if i in all_stations.keys():
                addlist = all_stations[i]
                addlist[directions[0]] = l[0]
                addlist[directions[1]] = l[1]
                all_stations[i] = addlist
            else:
                all_stations[i] = {directions[0]: l[0], directions[1]: l[1]}

all_links = 0
for i in all_stations.values():
    all_links += 1

N = 0

for stop in all_stations.keys():
    rows_st, rows_fn = 5, 29
    print(round(N/all_links*100, 1))

    worksheet, dirlink = workbook.add_worksheet(stop), all_stations[stop]

    for direct in dirlink.keys():

        rows = [rows_st, rows_fn]
        worksheet.write(0, 0, 'Station:')
        worksheet.write(0, 1, stop)
        worksheet.write(rows_st-2, 0, 'Direction:')
        worksheet.write(rows_st-2, 1, direct)
        worksheet.write(rows_st - 1, 1, 'Monday-Friday')
        worksheet.write(rows_st - 1, 13, 'Saturday')
        worksheet.write(rows_st - 1, 24, 'Sunday and public holiday')

        r0, h_tab = 1, []
        for r in range(rows_st, rows_fn):
            worksheet.write(r, 0, r0)
            h_tab.append(r0)
            r0 += 1

        worksheet.write(r, 0, 0)
        h_tab.append(0)

        link = dirlink[direct]

        if len(link) > 0:
            worksheet.write(rows_st-2, 5, 'Link')
            worksheet.write(rows_st-2, 6, link)

            numb = link.split('/')
            numb = list(numb[4])

            driver = webdriver.Chrome()
            driver.get(link)
            driver.minimize_window()

            # loading the page
            time.sleep(2)

            # full timetable
            element2 = driver.find_element("id", "divScheduleContentInner")
            timetable = element2.text
            timetable_list = re.split('\n', timetable)
            driver.close()

            time_int, time_all, h0 = [], [], 1
            int0 = [timetable_list.index('Tööpäev'), timetable_list.index('Laupäev,'),
                    timetable_list.index('Pühapäev ja riiklik püha')]
            for i2 in timetable_list:
                timetable_dig = i2.split()

                if i2 not in int0 and len(timetable_dig) == 2:

                    x = list(timetable_dig[1])
                    for i3 in range(0, len(x), 2):
                        x2 = x[i3] + x[i3+1]
                        time_int.append(timetable_dig[0]+':'+x2)
                else:
                    time_all.append(time_int)
                    time_int = []

            time_all.append(time_int)
            lendat, time_sap, len_all = [], [], []
            for i in time_all:
                if len(i) > 0:
                    lendat.append(len(i))
                    h0, time_h, time_h_temp, k = 5, [], [], 1
                    for i2 in i:
                        h = int(i2.split(':')[0])
                        if h != h0:
                            time_h.append(time_h_temp)
                            time_h_temp = []
                            h0 = h

                        time_h_temp.append(i2)
                    time_h.append(time_h_temp)
                    time_sap.append(time_h)
                    len_all.append(lendat)

            head = ['Monday-Friday', 'Saturday', 'Holidays']
            k, col = 0, [1, 13, 24]

            for i in time_sap:

                for i2 in i:
                    cols = col[k]
                    for i3 in i2:
                        h0 = int(i3.split(':')[0])
                        x = h_tab.index(h0)
                        if h0 != 0:
                            worksheet.write(rows_st+x, cols, i3, format_time)
                        else:
                            worksheet.write(rows_st + x-1, cols, i3, format_time)
                        cols += 1
                k += 1

            for r in range(rows_st, rows_fn):
                worksheet.write(r, 35, stop)
                worksheet.write(r, 36, direct)
            rows_st += 31
            rows_fn += 31
    N += 1

workbook.close()