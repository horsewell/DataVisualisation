# Modules
import os
import csv

# Set path for file
csvpath = os.path.join(".", "Resources", "budget_data.csv")

data = []

total_months        = 0
total_amount        = 0
average_change      = 0
greatest_inc_date   = ""
greatest_inc_amount = 0
greatest_dec_date   = ""
greatest_dec_amount = 0

# Open the CSV using the UTF-8 encoding
with open(csvpath, encoding='UTF-8') as csvfile:
    csv_reader = csv.reader(csvfile, delimiter=",")

    # store the header
    header_row = next(csv_reader)   

    for row in csv_reader:
        data = row

# analysis

# The total number of months included in the dataset
# The net total amount of "Profit/Losses" over the entire period
# The changes in "Profit/Losses" over the entire period, and then the average of those changes
# The greatest increase in profits (date and amount) over the entire period
# The greatest decrease in profits (date and amount) over the entire period

total_months        = 86
total_amount        = 22564198
average_change      = -8311.11
greatest_inc_date   = "Aug-16"
greatest_inc_amount = 1862002
greatest_dec_date   = "Feb-14"
greatest_dec_amount = -1825558

# output

print(f"Financial Analysis")
print(f"----------------------------")
print(f"Total Months: {total_months}")
print(f"Total: ${total_amount}")
print(f"Average Change: ${average_change}")
print(f"Greatest Increase in Profits: {greatest_inc_date} (${greatest_inc_amount})")
print(f"Greatest Decrease in Profits: {greatest_dec_date} (${greatest_dec_amount})")