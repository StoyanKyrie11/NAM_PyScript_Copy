import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import numpy as np

purchaseOrderHeader = []
purchaseOrderItems = []
purchaseOrderItemDescription = []
materialNumber = []
materialDescription = []
manufacturerPartNumber = []


# Function reading PO_DeliveryData.xlsx spreadsheet data
def read_PODeliveryData():
    # Read data/insert data into data frames - PODelivery_Data spreadsheet file
    df = pd.read_excel('PODelivery_Data.xlsx', sheet_name=0, index_col=False)
    purchaseOrderHeader = df['Purch.Doc.']
    purchaseOrderItems = df['Item']
    purchaseOrderItemDescription = df['Short Text']
    return purchaseOrderHeader, purchaseOrderItems, purchaseOrderItemDescription


# Function reading Materials_Master.xlsx spreadsheet data
def read_MatMasterData():
    # Read data/insert data into data frames - Materials_Master spreаdsheet file
    df = pd.read_excel('Materials_Master.xlsx', sheet_name=0, index_col=False)
    manufacturerPartNumber = df['Manufacturer Part No']
    materialNumber = df['Material']
    materialDescription = df['Material Description']
    return manufacturerPartNumber, materialNumber, materialDescription


# Call both functions with the returned data frame variable values
purchaseOrderHeader, purchaseOrderItems, purchaseOrderItemDescription = read_PODeliveryData()
manufacturerPartNumber, materialNumber, materialDescription = read_MatMasterData()

# Reformat data into numpy arrays for the process extraction
array_POData = np.array(purchaseOrderItemDescription)
array_MaterialsData = np.array(materialDescription)
array_MatNumber = np.array(materialNumber)
array_ManPartNumber = np.array(manufacturerPartNumber)

# Removing Nan values from data frame from Manufacturer Part Number column values
array_ManPartNumber = array_ManPartNumber[~pd.isnull(array_ManPartNumber)]

data = [process.extract(x1, array_POData, scorer=fuzz.token_sort_ratio, limit=3) for x1 in array_ManPartNumber]
data = [
    [x, y, results[0][0], results[0][1], results[1][0], results[1][1], results[2][0], results[2][1]]
    for (x, y, results)
    in zip(array_MatNumber, array_ManPartNumber, data)]
df = pd.DataFrame(data,
                  columns=['Material number', 'Material description', "Material first match", "First best match %",
                           "Material second match", "Second best match %", "Material third match",
                           "Third best match %"])
df.to_excel("BestMatch_Mat_Number_MatchPercentages.xlsx", index=False)

# Condition - Output a file whenever the First best match % is of a desired percentage value - requires input from user
percentageInput = str(input("Please, input a percentage threshold value: "))
data = pd.read_excel("BestMatch_Mat_Number_MatchPercentages.xlsx", index_col="Material number")
dataFiltered = data[data["First best match %"] >= int(percentageInput)]
dataFiltered.to_excel("Filtered_BestMatch_Mat_Number_MatchPercentages.xlsx", index=True)
