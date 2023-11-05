"""
This csv sorting algorithm is designed for the Rode Island Food Bank donation spreadsheets granted to Tapin at the end of each month.

Developed by: Landon Montecalvo
"""
import os
import csv

directory = "D:\Tapin\Tapin_Data_Algorithm"

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

months = {
    1: "Jan",
    2: "Feb",
    3: "Mar",
    4: "Apr",
    5: "May",
    6: "Jun",
    7: "Jul",
    8: "Aug",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dec"
}

# Temporary Storage Lists
dataRows = []
tapinCategories = []
productRefs = []
productNames = []
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

def GetCSV(): # Return all files with the .csv extension.
    files = []
    for file in os.listdir(directory):
        if file.endswith(".csv"):
            files.append(os.path.join(directory, file))

    return files

def RemoveCharacters(number): # Function to remove the comma and the $ from a number.
    number = number.replace(",", "")
    number = number.replace("$", "")
    return number

def AddMoneySign(number): # Function to add the $ to a number.
    number = "$" + str(number)
    return number

print("REMEMBER TO REMOVE ALL FORMATTED SHEET FILES FROM THIS FOLDER WHEN DONE!!!\n")

while True:
    try:
        month = int(input("MONTH NUMBER: "))
        
        if month > 1 and month < 13:
            break
        else:
            print("\nINVALID.")
            print("PLEASE ENTER A VALID MONTH NUMBER.\n")
    except ValueError:
        print("\nINVALID.")
        print("PLEASE ENTER A VALID MONTH NUMBER.\n")

while True:
    try:
        year = input("YEAR: ")
        yearNum = int(year)
        break
    except ValueError:
        print("\nINVALID.")
        print("PLEASE ENTER A VALID YEAR.\n")

with open(GetCSV()[0], 'r') as file: # Read data in the .csv file.
    reader = csv.reader(file)

    for index, value in enumerate(reader):
        if index > 2: # Removes header rows
            if value[0] != "Total Weight" and value[0] != "Service Fee" and value[0] != "Tap-In 113700" and value[0] != "Total \nWeight" and value[0] != " Grand Totals: " and value[0] != "Page -1 of 1": # Get rid of junk rows.
                if (not RemoveCharacters(value[0]).isdigit()): # Remove number rows.
                    dataRows.append(value) # Add item row to dataRows list.

    for index, value in enumerate(dataRows): # Lookup tapin category for each item and append it to tapinCategories.
        tempSplitStringType = value[2].split(" ")
        tempSplitStringCategory = value[3].split(" ")

        if float(tempSplitStringCategory[0]) < 30 or float(tempSplitStringCategory[0]) > 33:
            tapinCategories.append(tapinCategoriesList.get(tempSplitStringType[0]))
        elif float(tempSplitStringCategory[0]) > 29 and float(tempSplitStringCategory[0]) < 34:
            tapinCategories.append(tapinCategoriesList.get(tempSplitStringCategory[0]))

    for index, value in enumerate(dataRows):
        # Add corresponding data from csv to one of the lists above.
        productRefs.append(value[0])
        productNames.append(value[1])
        productTypes.append(value[2]) 
        productCategories.append(value[3])
        productWeights.append(value[4])
        productServiceFees.append(value[5])
        productPurchaseCosts.append(value[6])
        productValues.append(value[7])

    # Generate grand total for number columns.
    for weight in productWeights:
        totalWeight += int(RemoveCharacters((weight)))
    for serviceFee in productServiceFees:
        totalServiceFee += float(RemoveCharacters(serviceFee))
    for purchaseCost in productPurchaseCosts:
        totalPurchaseCost += float(RemoveCharacters(purchaseCost))
    for value in productValues:
        totalValue += float(RemoveCharacters(value))

    # Round all subtotals to 2 decimal places.
    round(totalWeight, 2)
    round(totalServiceFee, 2)
    round(totalPurchaseCost, 2)
    round(totalValue, 2)

    # Add money sign to the money subtotals.
    AddMoneySign(totalServiceFee)
    AddMoneySign(totalPurchaseCost)
    AddMoneySign(totalValue)

    fileName = "Food Bank Report - " + months[month] + " " + year + ".csv"

    with open(directory + "\\" + fileName, "w", newline="") as file:
        writer = csv.writer(file)

        formattedRows = []

        subtotals = {
            "General Food": [0, 0, 0, 0],
            "General Non-Food": [0, 0, 0, 0],
            "General Produce": [0, 0, 0, 0],
            "Cleaning Supplies": [0, 0, 0, 0],
            "Toiletries": [0, 0, 0, 0],
            "TEFAP Food": [0, 0, 0, 0],
            "TEFAP Produce": [0, 0, 0, 0]
        }

        writer.writerow(["Tap-In Category", "Product Ref", "ProductName", "Product Type", "Product Category", "Weight", "Service Fee", "Purchase Cost", "Food Bank Cost/Value"])

        for i in range(0, len(dataRows)):
            formattedRows.append([])
            
            formattedRows[i].append(tapinCategories[i])
            formattedRows[i].append(productRefs[i])
            formattedRows[i].append(productNames[i])
            formattedRows[i].append(productTypes[i])
            formattedRows[i].append(productCategories[i])
            formattedRows[i].append(productWeights[i])
            formattedRows[i].append(productServiceFees[i])
            formattedRows[i].append(productPurchaseCosts[i])
            formattedRows[i].append(productValues[i])

        formattedRows.sort()

        for row in formattedRows:
            tempCategory = row[0]
            subtotals[tempCategory][0] += float(RemoveCharacters(row[5]))
            subtotals[tempCategory][1] += float(RemoveCharacters(row[6]))
            subtotals[tempCategory][2] += float(RemoveCharacters(row[7]))
            subtotals[tempCategory][3] += float(RemoveCharacters(row[8]))

        for subtotal in subtotals:
            column = subtotals.get(subtotal)

            fee = "{:.2f}".format(column[1])
            cost = "{:.2f}".format(column[2])
            value = "{:.2f}".format(column[3])

            weight = str(column[0])
            fee = "$" + str(fee)
            cost = "$" + str(cost)
            value = "$" + str(value)

            column.clear()
            column.append(weight)
            column.append(fee)
            column.append(cost)
            column.append(value)

        for row in formattedRows:
            writer.writerow(row)
        
        writer.writerow([])

        if subtotals["General Food"] != ["0", "$0.00", "$0.00", "$0.00"]:
            contents = ["GENERAL FOOD", "", "", "", ""] + subtotals["General Food"]
            writer.writerow(contents)
        if subtotals["General Non-Food"] != ["0", "$0.00", "$0.00", "$0.00"]:
            contents = ["GENERAL NON-FOOD", "", "", "", ""] + subtotals["General Non-Food"]
            writer.writerow(contents)
        if subtotals["General Produce"] != ["0", "$0.00", "$0.00", "$0.00"]:
            contents = ["GENERAL PRODUCE", "", "", "", ""] + subtotals["General Produce"]
            writer.writerow(contents)
        if subtotals["Cleaning Supplies"] != ["0", "$0.00", "$0.00", "$0.00"]:
            contents = ["CLEANING SUPPLIES", "", "", "", ""] + subtotals["Cleaning Supplies"]
            writer.writerow(contents)
        if subtotals["Toiletries"] != ["0", "$0.00", "$0.00", "$0.00"]:
            contents = ["TOILETRIES", "", "", "", ""] + subtotals["Toiletries"]
            writer.writerow(contents)
        if subtotals["TEFAP Food"] != ["0", "$0.00", "$0.00", "$0.00"]:
            contents = ["TEFAP FOOD", "", "", "", ""] + subtotals['TEFAP Food']
            writer.writerow(contents)
        if subtotals["TEFAP Produce"] != ["0", "$0.00", "$0.00", "$0.00"]:
            contents = ["TEFAP PRODUCE", "", "", "", ""] + subtotals["TEFAP Produce"]
            writer.writerow(contents)

        grandTotalContents = ["GRAND TOTAL", "", "", "", ""] + [totalWeight, "$" + str(totalServiceFee), "$" + str(totalPurchaseCost), "$" + str(totalValue)]
        writer.writerow(grandTotalContents)
