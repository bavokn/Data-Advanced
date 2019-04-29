import random

from openpyxl import load_workbook
from openpyxl.styles import Font


class Spelertje:
    def __init__(self, name, position, goals, birthDate, effort, weight, length):
        self.name = name
        self.position= position
        self.goals = goals
        self.birthDate = birthDate
        self.effort = effort
        self.weight = weight
        self.length = length

    def returnArray(self):
        return [self.name, self.position, self.goals, self.birthDate, self.effort, self.weight, self.length]

    def __str__(self):
        return str(self.name) + " - " + str(self.position) + " - " + str(self.goals) + " - " + str(self.birthDate) + " - " + str(self.effort) + " - " + str(self.weight) + " - " + str(self.length)

class Spelertjes:
    def __init__(self):
        self.spelertjes = []
        self.effort = {1:"zeer goed", 2 : "goed", 3:"goed",4:"matig"}
        self.spelertjesValues = []

    def addSpeler(self, spelertje):
        self.spelertjes.append(spelertje)

    def readFile(self, fileName, sheetName):
        wb = load_workbook(fileName)
        ws = wb[sheetName]

        for row in ws.rows:
            args = [cell.value for cell in row]
            birthDate = random.randrange(1,5)
            spelertje = Spelertje(args[0],args[2],args[3],birthDate,self.effort[birthDate],args[6],args[7])
            self.spelertjesValues.append(spelertje.returnArray())
            self.addSpeler(spelertje)

        self.spelertjes.remove(self.spelertjes[0])
        self.spelertjesValues.remove(self.spelertjesValues[0])

    def writeFile(self, fileName, sheetName, saveFileName):
        wb = load_workbook(fileName)
        ws = wb[sheetName]

        font = Font(name='Calibri',size=12,bold=True)
        header = ["naam", "positie", "aantal gemaakte goalen", "geboortedatum", "inzet", "gewicht", "lengte"]

        #fill the headers with the right values
        for i in range(len(header)):
            cellref = ws.cell(1,i+1)
            cellref.value = header[i]
            cellref.font = font
            #clear last column
            cellref = ws.cell(1,i+2)
            cellref.value = None

        #fill gegevens with new generated values
        for i in range(len(self.spelertjesValues)):
            for j in range(len(self.spelertjesValues[i])):
                cellref = ws.cell(row=i+2, column=j+1)
                cellref.value = self.spelertjesValues[i][j]
                #clear last column
                cellref = ws.cell(row=i + 2, column=j + 2)
                cellref.value = None

        wb.save(saveFileName)
