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

r_amount = 0
p_amount = 0
r_date = ""
r_date = ""

change = 0
sum_change = 0


# Open the CSV using the UTF-8 encoding
with open(csvpath, encoding='UTF-8') as csvfile:
    csv_reader = csv.reader(csvfile, delimiter=",")

    # store the header
    header_row = next(csv_reader)   
    first = True

    # cycle through the rows
    for row in csv_reader:
        # collect the data
        r_date   = row[0]
        r_amount = int(row[1])

        # get the change in profit or loss
        change  = (r_amount - p_amount)

        # sum the change
        if not first:
            sum_change += change

        # get the total months and amount
        total_months += 1
        total_amount += r_amount

        if greatest_inc_amount < change:
            greatest_inc_date   = r_date
            greatest_inc_amount = change

        if greatest_dec_amount > change:
            greatest_dec_date   = r_date
            greatest_dec_amount = change

        p_amount = r_amount
        p_date   = r_date
        first = False

average_change = round(sum_change / (total_months -1), 2)

# output

print(f"Financial Analysis")
print(f"----------------------------")
print(f"Total Months: {total_months}")
print(f"Total: ${total_amount}")
print(f"Average Change: ${average_change}")
print(f"Greatest Increase in Profits: {greatest_inc_date} (${greatest_inc_amount})")
print(f"Greatest Decrease in Profits: {greatest_dec_date} (${greatest_dec_amount})")