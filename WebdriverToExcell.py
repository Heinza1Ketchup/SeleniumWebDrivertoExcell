from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from openpyxl.workbook import Workbook
import time
import openpyxl

alphabet = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
Statsheadercells = ["Team", "SRS", "SOS", "Goals For", "Goals Against", "Total shots", "Shot %", "Save %"]
headercells = ["Away", "Away Team Pts", "Home", "Home Team Pts", "Margin", "AwaySRS", "HomeSRS"]
def GetSRS(x,y, count):
    for x in range(x, y):
        TeamcellNum = "A" + str(count)
        elemTeam = driver.find_element_by_xpath(
            "/html/body/div[2]/div[5]/div[2]/div/div/div[2]/table/tbody/tr[" + str(x) + "]/th")
        TeamCell = wb.active[TeamcellNum]
        TeamCell.value = elemTeam.text

        TeamcellNumVal = "C" + str(count)
        elemSRS = driver.find_element_by_xpath(
            "/html/body/div[2]/div[5]/div[2]/div/div/div[2]/table/tbody/tr[" + str(x) + "]/td[9]")
        TeamcellNumVal = wb.active[TeamcellNumVal]
        TeamcellNumVal.value = float(elemSRS.text)
        count = count + 1
    return count

def WritetoExcell(x, teamindex, cellLetter, cellCount, floater):
    elemTeam = driver.find_element_by_xpath(
        "/html/body/div[5]/div[1]/div/div[1]/div[2]/div[4]/div[2]/div[1]/div/div/div[" + str(x) + "]/div[" + str(
            teamindex) + "]")
    TeamCellnum = cellLetter + str(cellCount)
    Teamcell = wb.active[TeamCellnum]
    if floater == False:
        Teamcell.value = elemTeam.text
    else:
        Teamcell.value = float(elemTeam.text)
    return Teamcell

### Write to Excell
wb = Workbook()
ws1 = wb.active
ws1.title = "Stats"

### Open up webdriver
driver = webdriver.Chrome()
driver.get("https://www.hockey-reference.com/leagues/")
assert "Hockey-Reference.com" in driver.title
### Navigation
elem = driver.find_element_by_xpath("/html/body/div[2]/div[4]/div[1]/div[2]/table/tbody/tr[1]/th")
elem.click()
time.sleep(1)
### Fill in data
count = 1
#count = GetSRS(2,10, count)
#count = GetSRS(11,19, count)
#count = GetSRS(20,27, count)
#count = GetSRS(28,36, count)

#/html/body/div[2]/div[5]/div[3]/div[4]/table/tbody/tr[1]/td[1] team
#/html/body/div[2]/div[5]/div[3]/div[4]/table/tbody/tr[1]/td[13] srs
#/html/body/div[2]/div[5]/div[3]/div[4]/table/tbody/tr[1]/td[14] sos
#/html/body/div[2]/div[5]/div[3]/div[4]/table/tbody/tr[1]/td[15] goals for
#/html/body/div[2]/div[5]/div[3]/div[4]/table/tbody/tr[1]/td[16] goals against
#/html/body/div[2]/div[5]/div[3]/div[4]/table/tbody/tr[1]/td[27] total shots
#/html/body/div[2]/div[5]/div[3]/div[4]/table/tbody/tr[1]/td[28] shot %
#/html/body/div[2]/div[5]/div[3]/div[4]/table/tbody/tr[1]/td[30] save %

### Headers
for x in range(1,9):
    headercellnum = alphabet[x - 1] + str(1)
    headercell = wb.active[headercellnum]
    headercell.value = Statsheadercells[x - 1]
### Fill in excell
for x in range(1, 32):
    TeamcellNum = "A" + str(x + 1)
    elemTeam = driver.find_element_by_xpath("/html/body/div[2]/div[5]/div[3]/div[4]/table/tbody/tr[" + str(x) + "]/td[1]")
    TeamCell = wb.active[TeamcellNum]
    TeamCell.value = elemTeam.text

    cellNum = "B" + str(x + 1)
    elemSRS = driver.find_element_by_xpath("/html/body/div[2]/div[5]/div[3]/div[4]/table/tbody/tr[" + str(x) + "]/td[13]")
    Cell = wb.active[cellNum]
    Cell.value = float(elemSRS.text)

    cellNum = "C" + str(x + 1)
    elem1 =  driver.find_element_by_xpath("/html/body/div[2]/div[5]/div[3]/div[4]/table/tbody/tr[" + str(x) + "]/td[14]")
    Cell = wb.active[cellNum]
    Cell.value = float(elem1.text)

    cellNum = "D" + str(x + 1)
    elem1 = driver.find_element_by_xpath("/html/body/div[2]/div[5]/div[3]/div[4]/table/tbody/tr[" + str(x) + "]/td[15]")
    Cell = wb.active[cellNum]
    Cell.value = float(elem1.text)

    cellNum = "E" + str(x + 1)
    elem1 = driver.find_element_by_xpath("/html/body/div[2]/div[5]/div[3]/div[4]/table/tbody/tr[" + str(x) + "]/td[16]")
    Cell = wb.active[cellNum]
    Cell.value = float(elem1.text)

    cellNum = "F" + str(x + 1)
    elem1 = driver.find_element_by_xpath("/html/body/div[2]/div[5]/div[3]/div[4]/table/tbody/tr[" + str(x) + "]/td[27]")
    Cell = wb.active[cellNum]
    Cell.value = float(elem1.text)

    cellNum = "G" + str(x + 1)
    elem1 = driver.find_element_by_xpath("/html/body/div[2]/div[5]/div[3]/div[4]/table/tbody/tr[" + str(x) + "]/td[28]")
    Cell = wb.active[cellNum]
    Cell.value = float(elem1.text)

    cellNum = "H" + str(x + 1)
    elem1 = driver.find_element_by_xpath("/html/body/div[2]/div[5]/div[3]/div[4]/table/tbody/tr[" + str(x) + "]/td[30]")
    Cell = wb.active[cellNum]
    Cell.value = float(elem1.text)

### Close webpage
driver.close()

### Open new page
driver = webdriver.Chrome()
driver.get("https://www.flashscore.com/nhl/results/")
assert "NHL Results" in driver.title
### Navigation
time.sleep(1)

#/html/body/div[6]/div[1]/div/div[1]/div[2]/div[4]/div[2]/div[1]/div/div/div[2]/div[2]
#/html/body/div[6]/div[1]/div/div[1]/div[2]/div[4]/div[2]/div[1]/div/div/div[116]/div[1]
#/html/body/div[6]/div[1]/div/div[1]/div[2]/div[4]/div[2]/div[1]/div/div/div[203]/div[1]
#/html/body/div[6]/div[1]/div/div[1]/div[2]/div[4]/div[2]/div[1]/div/div/div[303]/div[1]
#elem = driver.find_element_by_xpath("/html/body/div[6]/div[1]/div/div[1]/div[2]/div[4]/div[2]/div[1]/div/div/a")
#elem.get_attribute("href").click()
### Creat new excell sheet
ws2 = wb.create_sheet("Schedule")
wb.active = ws2
### Headers
for x in range(1,8):
    headercellnum = alphabet[x-1] + str(1)
    headercell = wb.active[headercellnum]
    headercell.value = headercells[x-1]



### Fill table
teamindex = 2
cellCount = 2
for x in range(2, 1273):
    try:
        elem = driver.find_element_by_xpath(
            "/html/body/div[5]/div[1]/div/div[1]/div[2]/div[4]/div[2]/div[1]/div/div/div[" + str(x) + "]/div[1]")
    except NoSuchElementException:
        break

    if (elem.get_attribute("class") == "event__check"):
        teamindex = 4
        scoreindex = 6
        floater = False
        Teamcell1 = WritetoExcell(x, teamindex, "A", cellCount, floater)
        teamindex = teamindex - 1
        Floater = False
        Teamcell2 = WritetoExcell(x, teamindex, "C", cellCount, floater)
        floater = True
        TeamcellScore1 = WritetoExcell(x, scoreindex, "B", cellCount, floater)
        scoreindex = scoreindex - 1
        floater = True
        TeamcellScore2 = WritetoExcell(x, scoreindex, "D", cellCount, floater)
    if (elem.get_attribute("class") == "event__time"):
        teamindex = 3
        scoreindex = 5
        floater = False
        Teamcell1 = WritetoExcell(x, teamindex, "A", cellCount, floater)
        teamindex = teamindex - 1
        floater = False
        Teamcell2 = WritetoExcell(x, teamindex, "C", cellCount, floater)
        floater = True
        TeamcellScore1 = WritetoExcell(x, scoreindex, "B", cellCount, floater)
        scoreindex = scoreindex - 1
        floater = True
        TeamcellScore2 = WritetoExcell(x, scoreindex, "D", cellCount, floater)
    floater = False
    marginCellnum = "E" + str(x)
    marginCell = wb.active[marginCellnum]
    marginCell.value = float(TeamcellScore1.value) - float(TeamcellScore2.value)
    ### get SRS from SRS tab
    awaySRScellNum = "F" + str(x)
    awaySRScell = wb.active[awaySRScellNum]
    awayTeam = Teamcell1.value
    wb.active = ws1
    for y in range(1, 33):
        srsCell = "A" + str(y + 1)
        srsCellRate = "B" + str(y + 1)
        SRSTeam = wb.active[srsCell]
        SRSRate = wb.active[srsCellRate]
        if SRSTeam.value == awayTeam:
            awaySRScell.value = float(SRSRate.value)
    wb.active = ws2
    homeSRScellNum = "G" + str(x)
    homeSRScell = wb.active[homeSRScellNum]
    homeTeam = Teamcell2.value
    wb.active = ws1
    for y in range(1, 32):
        srsCell = "A" + str(y)
        srsCellRate = "B" + str(y)
        SRSTeam = wb.active[srsCell]
        SRSRate = wb.active[srsCellRate]
        if SRSTeam.value == homeTeam:
            homeSRScell.value = float(SRSRate.value)
    wb.active = ws2


    cellCount = cellCount + 1
### Close webpage
driver.close()

### Save to excell
wb.save("D:\\Python\\WebdriverPractice\\basicmodel.xlsx")