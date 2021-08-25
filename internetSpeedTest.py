### Date: 25/08/2021
### Author: Dishant Vyas
### Desc: This Python script tests the speed of your internet
######### connection and notes the value down in an Excel sheet.
######### It may be automated using Windows Task Manager in
######### order to get test results at specified time intervals.



import speedtest, openpyxl
from datetime import datetime, date, time

itWorked = 0

def speedTest():
    test = speedtest.Speedtest()

    ########## To check download speed.
    print("\nPerforming download speed test...")
    numberOfTests = 3

    downloadSpeedTestResList = []
    for i in range(numberOfTests):
        downloadSpeedTestResList.append(test.download())
    downloadSpeedTestRes = sum(downloadSpeedTestResList)/len(downloadSpeedTestResList)

    print(f"Download Speed = {downloadSpeedTestRes:.2f} bps = {downloadSpeedTestRes/pow(1024,2):.2f} Mbps")
    downloadSpeedTestRes = round(downloadSpeedTestRes/pow(1024, 2), 2)

    ########## To check upload speed.
    # print("\nPerforming upload speed test...")
    # uploadSpeedTestRes = test.upload()
    # print(f"Upload Speed = {uploadSpeedTestRes:.2f} bps = {uploadSpeedTestRes/pow(1024,2):.2f} Mbps")

    # print("\nPerforming ping test...")
    # pingRes = test.results.ping
    # print(f"Ping (the response time of your connection) = {pingRes:.2f} ms\n")

    return downloadSpeedTestRes


try:
    downloadSpeedTestRes = speedTest()
    itWorked = 1

except:
    time.sleep(60)
    try:
        downloadSpeedTestRes = speedTest()
        itWorked = 1
    except:
        print("SPEED TEST FAILED: Computer not connected to internet!")


path = "InternetSpeedTestResults.xlsx"
wb = openpyxl.load_workbook(path.strip())
sheet = wb.active
editIndicator = sheet.cell(1,27)

row = int(editIndicator.value)
col = int(datetime.now().strftime("%H"))+2
editCell = sheet.cell(row,col)
dateCell = sheet.cell(row, 1)

dateCell.value = date.today().strftime("%d/%m/%Y")

if(itWorked):
    editCell.value = downloadSpeedTestRes
else:
    editCell.value = None

if(col==25):
    editIndicator.value = editIndicator.value + 1

wb.save(path)
print("\n###################### Internet speed noted succesfully. ##############################\n")