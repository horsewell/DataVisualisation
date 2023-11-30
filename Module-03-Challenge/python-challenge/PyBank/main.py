# Modules
import os
import csv
import copy

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

r_amount = 0
p_amount = 0
r_date = ""
r_date = ""



# Open the CSV using the UTF-8 encoding
with open(csvpath, encoding='UTF-8') as csvfile:
    csv_reader = csv.reader(csvfile, delimiter=",")

    # store the header
    header_row = next(csv_reader)   

    for row in csv_reader:
        r_date   = copy.deepcopy(row[0])
        r_amount = int(copy.deepcopy(row[1]))

        total_months += 1
        total_amount += int(row[1])

        if greatest_inc_amount < (r_amount - p_amount):
            greatest_inc_date   = r_date
            greatest_inc_amount = (r_amount - p_amount)

        if greatest_dec_amount > (r_amount - p_amount):
            greatest_dec_date   = r_date
            greatest_dec_amount = (r_amount - p_amount)

        p_amount = r_amount
        p_date   = r_date

        

# analysis

# The total number of months included in the dataset
# The net total amount of "Profit/Losses" over the entire period
# The changes in "Profit/Losses" over the entire period, and then the average of those changes
# The greatest increase in profits (date and amount) over the entire period
# The greatest decrease in profits (date and amount) over the entire period

#total_months        = 86
#total_amount        = 22564198
average_change      = -8311.11
#greatest_inc_date   = "Aug-16"
#greatest_inc_amount = 1862002
#greatest_dec_date   = "Feb-14"
#greatest_dec_amount = -1825558

# output

print(f"Financial Analysis")
print(f"----------------------------")
print(f"Total Months: {total_months}")
print(f"Total: ${total_amount}")
print(f"Average Change: ${average_change}")
print(f"Greatest Increase in Profits: {greatest_inc_date} (${greatest_inc_amount})")
print(f"Greatest Decrease in Profits: {greatest_dec_date} (${greatest_dec_amount})")