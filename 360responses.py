# Python 3

import requests
import zipfile
import json
import io
import send2trash, csv, os, openpyxl, win32com.client, shutil

#csv files removed to make things easier

try:
    send2trash.send2trash('somecsvfile')
    send2trash.send2trash('somecsvfile')
except:
    pass

surveys = {
    '360 Physician Survey: DOE, JOHN': 'QUALTRICS_SURVEY_ID',
    '360 Physician Survey: Other Physician': 'QUALTRICS_SURVEY_ID'
}

# Below is identifier for "Other physician" survey
physicianNumberDict = {
    'DOE, JOHN': '1'
}

def surveysD(surveys):
    for link in surveys.values():
        # Setting user Parameters
        apiToken = "YOUR API TOKEN"
        surveyId = link
        fileFormat = "csv"
        dataCenter = 'YOUR DATA CENTER'

        # Setting static parameters
        requestCheckProgress = 0
        progressStatus = "in progress"
        baseUrl = "https://{0}.qualtrics.com/API/v3/responseexports/".format(dataCenter)
        headers = {
            "content-type": "application/json",
            "x-api-token": apiToken,
            }

        # Step 1: Creating Data Export
        downloadRequestUrl = baseUrl
        downloadRequestPayload = '{"format":"' + fileFormat + '","surveyId":"' + surveyId + '"}'
        downloadRequestResponse = requests.request("POST", downloadRequestUrl, data=downloadRequestPayload, headers=headers)
        progressId = downloadRequestResponse.json()["result"]["id"]
        print(downloadRequestResponse.text)

        # Step 2: Checking on Data Export Progress and waiting until export is ready
        while requestCheckProgress < 100 and progressStatus is not "complete":
            requestCheckUrl = baseUrl + progressId
            requestCheckResponse = requests.request("GET", requestCheckUrl, headers=headers)
            requestCheckProgress = requestCheckResponse.json()["result"]["percentComplete"]
            print("Download is " + str(requestCheckProgress) + " complete")

        # Step 3: Downloading file
        requestDownloadUrl = baseUrl + progressId + '/file'
        requestDownload = requests.request("GET", requestDownloadUrl, headers=headers, stream=True)

        # Step 4: Unzipping the file
        zipfile.ZipFile(io.BytesIO(requestDownload.content)).extractall("MyQualtricsDownload")
        print('Complete')

surveysD(surveys)

#code redacted here that shows file movement and directory navigation

rows = []
attname = []

for files in directory:
    with open(files, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        for row in reader:
            if reader.line_num == 1 or reader.line_num == 2 or reader.line_num == 3:
                continue
            rows.append(row)
            attname.append(str(os.path.splitext(files)[0])[24:] + ',')

def csvtoexcel(csvfile):
    c = csvfile
    wb = openpyxl.Workbook()
    sheet = wb.create_sheet(index=0, title="Att")
    att = open(c, encoding="utf8")
    attcsv = csv.reader(att)
    for row in attcsv:
        sheet.append(row)
    wb.save(str(c)[:-4] + ".xlsx")
    wb.close()

def excelRefresh(source, f):
    SourcePathName = source
    FileName = f 
    Application = win32com.client.Dispatch("Excel.Application")
    Application.Visible = 1
    Workbook = Application.Workbooks.open(SourcePathName + '/' + FileName)
    Workbook.RefreshAll()
    Workbook.Save()
    Application.Quit()

#code redacted here that shows file movement and directory navigation

try:
    send2trash.send2trash('somecsvfile')
    send2trash.send2trash('somecsvfile')
except:
    pass

excelRefresh('H:/Personal/misc/360survey/360QualAuto/OtherAtt', 'AttId.xlsx')

wb1 = openpyxl.load_workbook('excel file', data_only=True)
sheet = wb1.get_active_sheet()
for i in range(1, 600):
    rowvalue = sheet.cell(row=i, column=1).value
    if rowvalue == 0:
        continue
    else:
        with open('names.csv', 'a', newline='') as f:
            f.write(str(rowvalue) + "\n")

othernames = []
ids = []

with open('somecsv.csv', 'r', encoding='utf-8') as f:
    reader = csv.reader(f)
    for row in reader:
        ids.append(str(row))
    for k, v in physicianNumberDict.items():
        if str("['" + v + "']") in ids:
            othernames.append(k)

excelRefresh('source', 'excelfile')

#code redacted here that shows file movement and directory navigation

with open('csvfile', 'a', newline='') as f:
    f.write('ResponseID,ResponseSet,IPAddress,StartDate,EndDate,RecipientLastName,RecipientFirstName,RecipientEmail,ExternalDataReference,Finished,Status,Physician Feedback Survey,text,clinical role,This physician treats patients in a professional and courteous manner,This physician treats staff in a professional and courteous manner,This physician communicates well and is responsive to nursesâ€™ concerns,This physician helps build a Team concept on each shift,This physician ensures patients are seen and dispositioned in timely fashion,This physician is able to prioritize and organize,This physician is able to remain calm under stress,This physician provides the patient/family with a thorough explanation of the problem and treatment,This physician explains decision-making to staff when questions arise,This physician effectively explains decision-making to PAs and is receptive to PA input and concerns,This physician adequately collaborates with the PAs on his/her team,')
    f.write('This physician adequately supports the PAs on his/her team,I look forward to coming to work when I see this physicians name on the schedule,Additional positive comments for the physician:,Additional constructive comments for the physician:,LocationLatitude,LocationLongitude,LocationAccuracy\n')
    writer = csv.writer(f)
    for row in rows:
        writer.writerow(row)
    wb = openpyxl.load_workbook('excelfile', data_only=True)
    sh = wb.get_active_sheet()
    for r in sh.rows:
        writer.writerow([cell.value for cell in r])

with open('csvfile', 'a', newline='') as f:
    f.write('Attending\n')
    writer = csv.writer(f)
    for i in attname:
        f.write(i + "\n")
    for i in othernames:
        writer.writerow([i])

files = ['csvfiles']

for i in files:
    csvtoexcel(i)

wb2 = openpyxl.load_workbook('excelfile', data_only=True)
wb2.save('excelfile_final.xlsx')
wb2.close()
