# Modules
import os
import csv

# Set path for file
csvpath = os.path.join(".", "Resources", "election_data.csv")

data = []

# Open the CSV using the UTF-8 encoding
with open(csvpath, encoding='UTF-8') as csvfile:
    csv_reader = csv.reader(csvfile, delimiter=",")

    # store the header
    header_row = next(csv_reader)   

    for row in csv_reader:
        data = row

# analysis

# The total number of votes cast
# A complete list of candidates who received votes
# The percentage of votes each candidate won
# The total number of votes each candidate won
# The winner of the election based on popular vote

total_votes = 369711
candidates  = ["Charles Casper Stockham", "Diana DeGette", "Raymon Anthony Doane"]
count_votes = [85213, 272892, 11606]
count_max   = 0
winner      = ""

# output

print(f"Election Results")
print(f"-------------------------")
print(f"Total Votes: {total_votes}")
print(f"-------------------------")
for candidate in candidates:
    candidate_votes = count_votes[candidates.index(candidate)]
    print(f"{candidate}: {candidate_votes/total_votes:.3%} ({candidate_votes})")
    if candidate_votes > count_max:
        count_max = candidate_votes
winner = candidates[count_votes.index(count_max)]
print(f"-------------------------")
print(f"Winner: {winner}")
print(f"-------------------------")