
from openpyxl import load_workbook
from Graphs import visual as visual
from Spelertjes import Spelertjes


def main():
    fixAndFillFile()
    fileName = "correct.xlsx"
    sheetName = "grafiek"
    graphs.drawScatterChart(fileName, sheetName, fileName)
    graphs.drawBarChart(fileName,sheetName,fileName)
    graphs.averageAndModus(fileName,sheetName,fileName)


def fixAndFillFile():
    fileName = 'voetbal.xlsx'
    sheetName = 'gegevens'
    saveFileName = "correct.xlsx"

    spelertjes.readFile(fileName, sheetName)
    spelertjes.writeFile(fileName, sheetName, saveFileName)


if __name__ == '__main__':
    spelertjes = Spelertjes()
    graphs = visual()

    main()





























