import openpyxl
filename="C:\\Users\prana\\OneDrive\\Documents\\Gant Chart Data\\DD.CrewList.CR.Jan 15_2022_as of Jan 27_2022.xlsx"
crewsheet="Dino Daycare crew"
frollcell="G13"
book=openpyxl.load_workbook(filename)
sheet=book[crewsheet]

