# load the data

# store the header

# analysis

total_months        = 86
total_amount        = 22564198
average_change      = -8311.11
greatest_inc_date   = "Aug-16"
greatest_inc_amount = 1862002
greatest_dec_date   = "Feb-14"
greatest_dec_amount = -1825558

# The total number of months included in the dataset
# The net total amount of "Profit/Losses" over the entire period
# The changes in "Profit/Losses" over the entire period, and then the average of those changes
# The greatest increase in profits (date and amount) over the entire period
# The greatest decrease in profits (date and amount) over the entire period

# output

print(f"Financial Analysis")
print(f"----------------------------")
print(f"Total Months: {total_months}")
print(f"Total: ${total_amount}")
print(f"Average Change: ${average_change}")
print(f"Greatest Increase in Profits: {greatest_inc_date} (${greatest_inc_amount})")
print(f"Greatest Decrease in Profits: {greatest_dec_date} (${greatest_dec_amount})")