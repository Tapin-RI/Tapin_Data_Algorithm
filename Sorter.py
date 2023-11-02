"""
This csv sorting algorithm is designed for the Rode Island Food Bank donation spreadsheets granted to Tapin at the end of each month.

Developed by: Landon Montecalvo
"""
import glob
import csv

tapinCategoriesList = {
    "01": "General Non-Food",
    "02": "General Food",
    "03": "General Food",
    "04": "General Food",
    "05": "General Food",
    "06": "General Food",
    "07": "General Food",
    "08": "General Food",
    "09": "General Food",
    "10": "General Food",
    "11": "General Food",
    "12": "Toiletries",
    "13": "Cleaning Supplies",
    "14": "General Food",
    "15": "General Food",
    "16": "General Food",
    "17": "Toiletries",
    "18": "General Food",
    "19": "General Non-Food",
    "20": "General Non-Food",
    "21": "General Food",
    "22": "General Produce",
    "23": "General Food",
    "24": "General Food",
    "25": "General Food",
    "26": "General Food",
    "27": "General Food",
    "28": "General Produce",
    "29": "General Food",
    "30": "TEFAP Food",
    "31": "TEFAP Food",
    "32": "TEFAP Food",
    "33": "TEFAP Produce"
}

# Temporary Storage Lists
dataRows = []
tapinCategories = []
productRefs = []
productTypes = []
productCategories = []
productWeights = []
productServiceFees = []
productPurchaseCosts = []
productValues = []

totalWeight = 0
totalServiceFee = 0
totalPurchaseCost = 0
totalValue = 0

files = glob.glob('*.csv') # Return all files with the .csv extension.

def RemoveCharacters(number): # Function to remove the comma and the $ from a number.
    number = number.replace(",", "")
    number = number.replace("$", "")
    return number

with open(files[0], 'r') as file: # Read data in the .csv file.
    reader = csv.reader(file)

    for index, value in enumerate(reader):
        if index > 2: # Removes header rows
            if value[0] != "Total Weight" and value[0] != "Service Fee" and value[0] != "Tap-In 113700" and value[0] != "Total \nWeight" and value[0] != " Grand Totals: " and value[0] != "Page -1 of 1": # Get rid of junk rows.
                if (not RemoveCharacters(value[0]).isdigit()): # Remove number rows.
                    dataRows.append(value) # Add item row to dataRows list.

    for index, value in enumerate(dataRows): # Lookup tapin category for each item and append it to tapinCategories.
        tempSplitString = value[2].split(" ")
        tapinCategories.append(tapinCategoriesList.get(tempSplitString[0]))

    for item in dataRows:
        print(item)

    for index, value in enumerate(dataRows):
        # Add corresponding data from csv to one of the lists above.
        productRefs.append(value[0])
        productTypes.append(value[2]) 
        productCategories.append(value[3])
        productWeights.append(value[4])
        productServiceFees.append(value[5])
        productPurchaseCosts.append(value[6])
        productValues.append(value[7])

    for item in productWeights:
        print(item)

    # Generate subtotal for number columns.
    for weight in productWeights:
        totalWeight += int(weight)
    for serviceFee in productServiceFees:
        totalServiceFee += float(RemoveCharacters(serviceFee))
    for purchaseCost in productPurchaseCosts:
        totalPurchaseCost += float(RemoveCharacters(purchaseCost))
    for value in productValues:
        totalValue += float(RemoveCharacters(value))
