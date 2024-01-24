#Importing libraries
import pandas as pd
import numpy as np

#Function cleaning and merging the data from the database
def CleanAndMerge(name, sheetName, count):
    #Using the excel spreadsheets give to collect specified rows
    df = pd.read_excel(name, sheet_name=sheetName)
    rowsToDrop = []
    for i in range(0, len(df)):
        if(i <= 42 or i >= 55):
            rowsToDrop.append(i)
    df = df.drop(rowsToDrop)
    columnsToDrop = []
    for i in range(0, 27):
        if(i != 1 and i != 11):
            columnsToDrop.append(i)
    df = df.drop(df.columns[columnsToDrop], axis=1)

    #If this is the first time it is cleaning, create a new doccument, otherwise combine the collected data with the created file
    if (count == 0):
        df.to_excel("Collected Robberies.xlsx")
    else:
        collectiveDf = pd.read_excel("Collected Robberies.xlsx")
        combined_file = pd.concat([collectiveDf, df])
        combined_file.to_excel("Collected Robberies.xlsx")

#Calling the cleaning and merging function
count = 0
CleanAndMerge('policeforceareatablesyearendingmarch2018v2.xlsx', 'Table P1', count)
count = 1
CleanAndMerge('policeforceareadatatablesyearendingjune2018corrected.xlsx', 'Table P1', count)
CleanAndMerge('policeforceareatablesyearendingsep18corrected.xlsx', 'Table P1', count)
CleanAndMerge('policeforceareatablesyearendingdecember2018.xlsx', 'Table P1', count)
CleanAndMerge('policeforceareatablesyeendingmarch2019.xlsx', 'Table P1 ', count)
CleanAndMerge('policeforceareatablesyearendingjune2019.xlsx', 'Table P1', count)
CleanAndMerge('policeforceareatablesyearendingseptember2019.xlsx', 'Table P1', count)
CleanAndMerge('pfatablesdec19.xlsx', 'Table P1', count)

#Cleaning the collected robberies and adding a date to each collected statistic
df = pd.read_excel("Collected Robberies.xlsx")
columnsToDrop = []
for i in range(0, 8):
    columnsToDrop.append(i)
df = df.drop(df.columns[columnsToDrop], axis=1)
df = df.rename({'Unnamed: 1':'Area Name'}, axis=1)
df = df.rename({'Unnamed: 11':'Robbery Number'}, axis=1)
date = []
for i in range(96):
    if(i <= 11):
        date.append("2018-03")
    elif(i <= 23):
        date.append("2018-06")
    elif (i <= 35):
        date.append("2018-09")
    elif (i <= 47):
        date.append("2018-12")
    elif (i <= 59):
        date.append("2019-03")
    elif (i <= 71):
        date.append("2019-06")
    elif (i <= 83):
        date.append("2019-09")
    elif (i <= 95):
        date.append("2019-12")

#Creating a list of region and earning statistics
region = []
regionList = []
earnings = []
earningsList = []
for i in range (12):
    if (i <= 5):
        region.append(0)
        if(i == 0):
            earnings.append(2)
        else:
            earnings.append(1)
    else:
        region.append(1)
        if (i == 6):
            earnings.append(2)
        else:
            earnings.append(0)

#Creating a list of robbery factor
robberies = []
robColumns = [8,9]
mid = 1959 / 2
df2 = pd.read_excel("Collected Robberies.xlsx", usecols=robColumns)
for i in range(96):
    if(i == 0 or i % 6 == 0):
        robberies.append(2)
    else:
        if(df2.iat[i, 1] >= mid):
            robberies.append(1)
        else:
            robberies.append(0)

#Creating list of population
population = []
populationList = []
populationFactor = []
populationFactorList = []
popColumns = [1,2]
df3 = pd.read_excel("pfatablesdec19.xlsx", sheet_name="Table P3", usecols=popColumns)
mid = 1200000
rowsToDrop = []
for i in range(0, len(df3)):
    if(i <= 42 or i >= 55):
        rowsToDrop.append(i)
df3 = df3.drop(rowsToDrop)
for i in range(12):
    population.append(df3.iat[i, 1])
    if (i == 0 or i % 6 == 0):
        populationFactor.append(2)
    else:
        if (df3.iat[i, 1] >= mid):
            populationFactor.append(1)
        else:
            populationFactor.append(0)

for i in range(8):
    regionList.extend(region)
    earningsList.extend(earnings)
    populationList.extend(population)
    populationFactorList.extend(populationFactor)

#Adding the newly created columns
df.insert(0, "Date", date)
df.insert(1, "Region (SE = 0, SW = 1)", regionList)
df.insert(4, "Robbery Factor", robberies)
df.insert(5, "Earnings Factor", earningsList)
df.insert(6, "Population",populationList)
df.insert(7, "Population Factor",populationFactorList)
df.to_excel("Clean Collected Robberies.xlsx", index=False)

#Creating earnings csv file
earningColumns = []
for i in range(0,22):
    earningColumns.append(i + 3)
southEastEarning = pd.read_excel("regionalgrossdisposablehouseholdincomelocalauthorityukjsoutheast.xls", sheet_name="Table 2", usecols=earningColumns)
southWestEarning = pd.read_excel("regionalgrossdisposablehouseholdincomelocalauthorityukksouthwest.xls", sheet_name="Table 2", usecols=earningColumns)
southEastEarningAvg = []
southWestEarningAvg = []
years = []
southEast = np.empty(shape=(22,67), dtype=int)
southWest = np.empty(shape=(22,30), dtype=int)
for c in range(0, 22):
    for r in range(0, 67):
        if(r <= 29):
            southWest[c][r] = southWestEarning.iat[r + 1, c]

        southEast[c][r] = southEastEarning.iat[r + 1, c]

    southEastEarningAvg.append(np.mean(southEast[c]))
    southWestEarningAvg.append(np.mean(southWest[c]))
    years.append(1997 + c)
earningFile = pd.DataFrame(list(zip(southEastEarningAvg, southWestEarningAvg)), index=[years], columns=['South East', 'South West'])
earningFile.to_csv("Earnings Average South East and South West.csv")
