# -*- Coding: utf-8 -*-

from bs4 import BeautifulSoup
import requests
import pandas as pd


CASEID = [i for i in range(18, 340)]
list_of_dicts_of_all_cases = []

for caseID in CASEID:
    url = "http://www.dsd.at/EESEstimation/Results.php?CaseID={}&Names=0".format(caseID)

    soup = BeautifulSoup(requests.get(url).text, 'lxml')

    if len(soup.find_all("p")) == 1:
        pass
    else:
        allRows = soup.find_all('tr')
        veh_num = 0
        Manus = [i for i in allRows[2].find_all('td')]
        for i in range(1, 4):
            if Manus[i].text != u'':
                veh_num += 1
        print veh_num

        blank_row_num = 0
        for i in range(len(allRows)):
            if allRows[i].text != u'&nbsp':
                blank_row_num += 1
            else:
                break

        list_of_vehicles_in_1_case = []

        for i in range(veh_num):
            case_dict = {
                'Case ID': 0,
                'Vehicle': 0,
                'Manufacture': '',
                'Model': '',
                'Year': '',
                'Curb weight': 0,
                'Average': 0,
                'std deviation': 0
            }
            case_dict['Case ID'] = int(caseID)
            case_dict['Vehicle'] = i + 1
            case_dict['Manufacture'] = Manus[i + 1].text
            case_dict['Model'] = [x for x in allRows[3].find_all('td')][i + 1].text
            case_dict['Year'] = [x for x in allRows[4].find_all('td')][i + 1].text
            # try:
            case_dict['Curb weight'] = int([x for x in allRows[5].find_all('td')][i + 1].text)
            try:
                case_dict['Average'] = float([x for x in allRows[blank_row_num + 3].find_all('td')][2 * i + 1].text)
                case_dict['std deviation'] = float([x for x in allRows[blank_row_num + 5].find_all('td')][2 * i + 1].text)
            except IndexError:
                print(u'Scheisse datei ID: ' + str(caseID))
            except ValueError:
                print("bad value at ID: " + str(caseID))

            list_of_vehicles_in_1_case.append(case_dict)

        for dict in list_of_vehicles_in_1_case:
            list_of_dicts_of_all_cases.append(dict)

dataframe = pd.DataFrame(list_of_dicts_of_all_cases)

writer = pd.ExcelWriter('output.xlsx')
dataframe.to_excel(writer)
writer.save()
