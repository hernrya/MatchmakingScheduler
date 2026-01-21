#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
NOTES:
This code was developed by Ryan Hernandez. Any questions? Reach out directly: ryan.hernandez@ucsf.edu.
(c) 2026
"""

import pandas as pd
import os
import re
from collections import defaultdict
import random

# Verbosity and Debug options
VERBOSE = False  # Set this variable to True to print out verbose messages; primarily for debugging
DEBUG = False    # Set this variable to True to print out lines to debug; mostly not useful!
DEBUGPAUSE = False  # Set this variable to True to pause at debug messages

# Set base directories for input (DIR) and output (DIROUT)
DIR = "../"
DIROUT = "../OUT/"

if not os.path.exists(DIROUT):
    os.makedirs(DIROUT)

#this is a simple function to remove non-printable control characters.
def clean_string(text):
    return ''.join(c for c in text if c.isprintable())

def randomize_order(items):
    randomized_items = list(items)
    random.shuffle(randomized_items)
    return randomized_items

# Set the start/stop times of each interview; this also sets the maximum number of interviews
TIMES = [
    "9:45-9:55P/11:45-11:55C",
    "10:00-10:10P/12:00-12:10C",
    "10:15-10:25P/12:15-12:45C",
    "10:30-10:40P/12:30-12:40C",
    "11:15-11:25P/1:15-1:25C",
    "11:30-11:40P/1:30-1:40C",
    "11:45-11:55P/1:45-1:55C",
    "12:00-12:10P/2:00-2:10C"
]
if VERBOSE:
    print("Interview Times:")
    for time_slot in TIMES:
        print(time_slot)

# Initialize dictionaries to store various information
Faculty = {}  # Faculty who registered; the keys of this dictionary are email addresses, and other info is stored in a list
##  {email} ->
##  [0]=name;
##  [1]=meeting link;
##  [2]=University;
##  [3]=total number of research interests;
##  [4-k]=each of their research interests;
nFaculty = 0
facultyNAMEtoEMAIL = {}  # Map faculty names to their email
facultyEMAILtoNAME = {}  # Map faculty emails to their name
FacAvail = {}  # Record whether faculty are available; default is yes
cat2fac = {}  # Map of research categories to faculty
facCatPop = {}  # Popularity of categories
nfacCatPop = 0  # Total number of research categories selected across faculty
int2fac = {}  # Map of research interests to faculty
facIntPop = {}  # Popularity of interests
nfacIntPop = 0  # Total number of research interests selected across faculty
fac2uni = {} # map faculty to their university
uni2fac = {} # map university to their faculty
nUniFac = {} # track number of faculty at each university

facCancel = {}  # Keep track of faculty that cancel to remove from schedule
nfacCancel = 0  # Total number of faculty that have canceled
scholCancel = {}  # Keep track of scholars that cancel to remove from schedule
nscholCancel = 0  # Total number of scholars that have canceled

# File path for the faculty cancellations
faculty_cancel_file = os.path.join(DIR, "Data/Faculty_Cancel.xlsx")

# Check if the file path is not empty
if faculty_cancel_file:
    try:
        # Read the Excel file into a pandas DataFrame
        df = pd.read_excel(faculty_cancel_file, skiprows=0, header=None)

        # Initialize line count
        nlines = 0

        # Iterate over the DataFrame rows
        for index, row in df.iterrows():
            data = row.tolist()
            nlines += 1
            if nlines == 1:  # Skip header line
                continue

            # Concatenate the first and second columns to form the full name
            first_name = str(data[0]).strip().lower()
            last_name = str(data[1]).strip().lower()
            full_name = first_name + last_name
            email = clean_string(str(data[2]).strip().lower())

            # Remove all control characters and non-printable characters
            full_name = clean_string(full_name)

            # Store the name in the facCancel dictionary if it doesn't exist
            if email not in facCancel:
                nfacCancel += 1
                facCancel[email] = 1

        print(f"there are {nfacCancel} faculty cancellations")
        # Print the contents of the facCancel dictionary
        if VERBOSE & nfacCancel>0:
            print("\nFaculty Cancellations:")
            for name in facCancel:
                print(f"{name}")

    except Exception as e:
        print(f"Cannot read {faculty_cancel_file}\nError: {e}")

# File path for the scholar cancellations
scholar_cancel_file = os.path.join(DIR, "Data/Scholar_Cancel.xlsx")

# Check if the file path is not empty
if scholar_cancel_file:
    try:
        # Read the Excel file into a pandas DataFrame
        df = pd.read_excel(scholar_cancel_file, skiprows=0, header=None)

        # Initialize line count
        nlines = 0

        # Iterate over the DataFrame rows
        for index, row in df.iterrows():
            data = row.tolist()
            nlines += 1
            if nlines == 1:  # Skip header line
                continue

            # store emails
            email = clean_string(str(data[2]).strip())

            # Store the name in the scholCancel dictionary if it doesn't exist
            if email not in scholCancel:
                nscholCancel += 1
                scholCancel[email] = 1

        print(f"there are {nscholCancel} scholar cancellations")
        # Print the contents of the facCancel dictionary
        if VERBOSE & nscholCancel>0:
            print("\nScholar cancellations:")
            for name, count in scholCancel.items():
                print(f"{name}")

    except Exception as e:
        print(f"Cannot read {scholar_cancel_file}\nError: {e}")

# Indicates if some/all faculty meeting links are in a separate file
meetinglinkFile = 0

# Dictionary to store all the faculty meeting links, with email addresses as keys
FacultyMeetingType = {}
FacultyMeetingLinks = {}
nFacultyMeetingLinks = 0  # Total number of meeting links

if meetinglinkFile == 1:
    # Now let's get faculty meeting links
    facMeetingFile = os.path.join(DIR, "Data/MissingFacultyMeetingLinks.xlsx")
    if VERBOSE:
        print(f"Looking for faculty meeting links in file {facMeetingFile}")
    
    try:
        # Read the Excel file into a pandas DataFrame
        df = pd.read_excel(facMeetingFile)

        # Initialize line count
        nlines = 0

        # Iterate over the DataFrame rows
        for index, row in df.iterrows():
            nlines += 1
            if nlines == 1:  # Skip header line
                continue

            # Get the email and link from the appropriate columns
            email = str(row[2]).strip().lower()
            link = str(row[3]).strip()

            # Remove all control characters and non-printable characters
            email = clean_string(email)

            # Store the email and link in the FacultyMeetingLinks dictionary if it doesn't exist
            if email not in FacultyMeetingLinks:
                FacultyMeetingLinks[email] = link
                nFacultyMeetingLinks += 1
            else:
                print(f"error, {email} duplicated in faculty meeting file?")
                exit(1)

        # Print the contents of the FacultyMeetingLinks dictionary
        if VERBOSE:
            print(f"\n{nFacultyMeetingLinks} Faculty Meeting links read:")
            for email, link in FacultyMeetingLinks.items():
                print(f"{email}: {link}")

    except Exception as e:
        print(f"Cannot read {facMeetingFile}\nError: {e}")


# File path for the faculty registration data
facRegFile = os.path.join(DIR, "Data/Faculty_Registration.xlsx")

# Read the Excel file into a pandas DataFrame
df = pd.read_excel(facRegFile, skiprows=1, header=None) #in qualtrics, first line is not helpful, and want to define my own headers

# Dictionary to map relevant headers to their column indices
facRegHeaderColumns = {
    "FirstName": 2,
    "LastName": 3,
    "email": 4,
    "University": 5,
    "ResearchCat": 8,
    "ResearchInt1": 10,
    "ResearchInt2": 12,
    "ResearchInt3": 14,
    "website": 17,
    "MeetingType": 18,
    "MeetingLink": 19
}

# Find the maximum column index
maxColumnEntry = max(facRegHeaderColumns.values())
nlines = 0
# Iterate over the DataFrame rows
for index, row in df.iterrows():
    nlines += 1
    
    if DEBUG:
        print(f"{nlines}; line={row}")

    if nlines == 1:  # From Qualtrics, header info is on line 2 and data begins on line 4... is yours the same?
        facRegHeaders = row.tolist()

        nfacRegHeaders = len(facRegHeaders)
        if VERBOSE:
            print("Check that the column names match the desired columns:")
            for col in sorted(facRegHeaderColumns, key=facRegHeaderColumns.get):
                print(f"column {col}: {facRegHeaderColumns[col]}")
                print(f"\t{facRegHeaderColumns[col]}: {col} =?= {facRegHeaders[facRegHeaderColumns[col]]}")
            print("\n")
        continue #check which line data will start on!
    if nlines < 3:
        continue

    # Convert row to a list of values
    sp = row.tolist()

    # Skip rows with only one value
    if len(sp) == 1:
        continue

    if DEBUG:
        print(f"line {nlines}: {sp}")
        for col in sorted(facRegHeaderColumns, key=facRegHeaderColumns.get):
            print(f"{col}: {sp[facRegHeaderColumns[col]]}")

    # Concatenate the first and last names to form the full name
    first_name = clean_string(str(sp[facRegHeaderColumns["FirstName"]])).strip()
    last_name = clean_string(str(sp[facRegHeaderColumns["LastName"]])).strip()
    CapName = first_name + " " + last_name
    name = first_name.lower().replace(" ","") + last_name.lower().replace(" ","")
    email = clean_string(str(sp[facRegHeaderColumns["email"]])).strip().lower()
    uni = clean_string(str(sp[facRegHeaderColumns["University"]])).strip()

    if VERBOSE:
        print(f"\nReading faculty {CapName} at {uni} on line {nlines}: {email}")

    # Check if the faculty member has canceled
    if email in facCancel:
        if VERBOSE:
            print(f"\n\n\n\nFaculty {CapName} cancelled. Skipping. *******")
        continue


    # Check for duplicate emails and issue a warning if found
    if VERBOSE:
        if email in Faculty:
            print(f"\n\n\n\nWarning: {email} repeated in registration file line {nlines}; overwriting with new info *****")
            exit(1)

    # If meetinglink is not included in a separate document, collect it from registration data
    if email not in FacultyMeetingLinks and facRegHeaderColumns["MeetingLink"] != "":
        FacultyMeetingType[email] = clean_string(str(sp[facRegHeaderColumns["MeetingType"]])).strip()
        FacultyMeetingLinks[email] = clean_string(str(sp[facRegHeaderColumns["MeetingLink"]])).strip()
        nFacultyMeetingLinks += 1
    elif email not in FacultyMeetingLinks:
        print(f"\n\n\n\nWarning: Did not find meeting link for {email} in MeetingLink file ******")
        exit(1)

    if VERBOSE:
        print(f"\tMeetingLink: {FacultyMeetingLinks[email]}")

    # Capture faculty research categories
    numCat = 0
    resCat = clean_string(str(sp[facRegHeaderColumns["ResearchCat"]])).strip().lower()
    cat = resCat.split(',')

    for c in cat:
        cat2fac.setdefault(c, {})[email] = numCat
        facCatPop[c] = facCatPop.get(c, 0) + 1
        nfacCatPop += 1
        numCat += 1

    if VERBOSE:
        print(f"\tResearchCat: {resCat}")

    # Capture faculty research interests
    numInt = 0
    interests = []
    for i in range(1, 4):
        interest = clean_string(str(sp[facRegHeaderColumns[f"ResearchInt{i}"]])).strip().lower()
        if "other" in interest:
            interest = clean_string(str(sp[facRegHeaderColumns[f"ResearchInt{i}"] + 1])).strip().lower()
        if interest and interest != "other" and interest != "nopreference":
            facIntPop[interest] = facIntPop.get(interest, 0) + 1
            nfacIntPop += 1
            int2fac.setdefault(interest, {})[email] = numInt
            interests.append(interest)
            numInt += 1
    if VERBOSE:
        print(f"\tResearchInt: {interests}")

    # Add faculty information to the Faculty dictionary
    if email not in Faculty:
        nFaculty += 1
    Faculty[email] = [name, FacultyMeetingLinks.get(email), uni, numInt] + interests
    uni2fac.setdefault(uni, {})[email] = 1
    if VERBOSE:
        print(f"\tUni: {uni}")

    # Check if the faculty member has a meeting link
    if Faculty[email][1] is None:
        raise ValueError(f"Faculty {email} does not have a meeting link")

    facultyNAMEtoEMAIL[name] = email
    facultyEMAILtoNAME[email] = CapName

    # Set default faculty availability; updated below in choices
    FacAvail[email] = [1] * 8

    if VERBOSE:
        print(f"Read faculty {CapName} {email} on line {nlines}:")
        for i, info in enumerate(Faculty[email]):
            print(f"\t[{i}]={info}")


print(f"Read {nFaculty} unique faculty across {len(uni2fac)} Universities:")
for uni in sorted(uni2fac):
    print(f"\t{uni}: {len(uni2fac[uni])}")

# Initialize dictionaries to store various scholar information
Scholars = {}  # Scholars who registered; the keys of this dictionary are email addresses, and other info is stored in a list
##  {email} ->
##  [0]=name
##  [1]=num interests
##  [2-k]=interest k;
ScholarUni = {} # map scholar to their universities of interest
## {email} -> 
##[0]=number of uni interests
##[1-k]=uni1; uni2; ...
nScholars = 0
scholN2E = {}  # Map scholar name to email
scholE2N = {}  # Map email to scholar name
cat2schol = {}  # Map of research categories to scholars
scholCatPop = {}  # Scholar research category popularity
nscholCatPop = 0  # Total number of research categories selected across scholars
int2schol = {}  # Map of research interests to scholars
scholIntPop = {}  # Scholar research interest popularity
nscholIntPop = 0  # Total number of research interests selected across scholars
uni2schol = {} # map university to the scholars interested
# File path for the scholar registration data
scholRegFile = os.path.join(DIR, "Data/Scholar_Registration.xlsx")

# Read the Excel file into a pandas DataFrame
scholReg_df = pd.read_excel(scholRegFile, skiprows=0, header=None)  # Assuming the first line is not useful and defining own headers

# Dictionary to map relevant headers to their column indices
scholRegHeaderColumns = {
    "FirstName": 0,
    "LastName": 1,
    "email": 2,
    "SchoolChoice": 10,
    "ResearchCat": 11,
    "ResearchInt1": 13,
    "ResearchInt2": 15,
    "ResearchInt3": 17
}

# Find the maximum column index
maxColumnEntry = max(scholRegHeaderColumns.values())

nlines = 0
for index, row in scholReg_df.iterrows():
    nlines += 1
    if DEBUG:
        print(f"{nlines}; line={row}")

    if nlines == 1:  # Parse header line
        scholRegHeaders = row.tolist()
        nscholRegHeaders = len(scholRegHeaders)
        if VERBOSE:
            print("\n\nNow check that the scholar column names match the desired columns:")
            for col in sorted(scholRegHeaderColumns, key=scholRegHeaderColumns.get):
                print(f"\t{col}[{scholRegHeaderColumns[col]}] =?= {scholRegHeaders[scholRegHeaderColumns[col]]}")
            print("\n")
        continue
    
    # Convert row to a list of values
    sp = row.tolist()

    # Skip rows with only one value
    if len(sp) < nscholRegHeaders:
        raise ValueError(f"Check for newlines on line {nlines}: {sp}")

    if DEBUG:
        print(f"line {nlines}: {sp}")
        for col in sorted(scholRegHeaderColumns, key=scholRegHeaderColumns.get):
            print(f"{col}: {sp[scholRegHeaderColumns[col]]}")
    
    # Concatenate the first and last names to form the full name
    first_name = clean_string(str(sp[scholRegHeaderColumns["FirstName"]])).strip()
    last_name = clean_string(str(sp[scholRegHeaderColumns["LastName"]])).strip()
    CapName = first_name + " " + last_name
    name = first_name.lower().replace(" ","") + last_name.lower().replace(" ","")

    # Get the email and clean it
    email = clean_string(str(sp[scholRegHeaderColumns["email"]])).strip().lower()

    if VERBOSE:
        print(f"\nReading scholar {CapName} on line {nlines}: {email}")

    # Check if the scholar has canceled
    if email in scholCancel:
        if VERBOSE:
            print(f"Scholar {CapName} cancelled, skip reading data")
            quit()
        continue


    # Capture scholar research categories
    numCat = 0
    resCat = str(sp[scholRegHeaderColumns["ResearchCat"]]).replace("\n", ",")
    cat = resCat.split(',')

    for c in cat:
        c = c.lower()
        scholCatPop[c] = scholCatPop.get(c, 0) + 1
        nscholCatPop += 1
        cat2schol.setdefault(c, {})[email] = numCat
        numCat += 1
        if DEBUG:
            if c not in facCatPop:
                print(f"\t\t***\t\tschol cat \"{c}\" NOT in facCatPop")
                if DEBUGPAUSE: #change to True to pause on this error
                    user_entry = input("Continue? (y/n)")
                    if user_entry != "y" and user_entry != "":
                        exit(1)

    numInt = 0
    interests = []
    for i in range(1, 4):
        research_int = clean_string(str(sp[scholRegHeaderColumns[f"ResearchInt{i}"]])).strip().lower()
        if "other" in research_int:
            research_int = clean_string(str(sp[scholRegHeaderColumns[f"ResearchInt{i}"] + 1])).strip().lower()
        if research_int and research_int != "other" and research_int != "nopreference":
            scholIntPop[research_int] = scholIntPop.get(research_int, 0) + 1
            nscholIntPop += 1
            int2schol.setdefault(research_int, {})[email] = numInt
            numInt += 1
            interests.append(research_int)

    # Capture scholar university choices
    unis = str(sp[scholRegHeaderColumns["SchoolChoice"]]).split("\n")
    numUni = len(unis)
    ScholarUni[email] = [numUni] + unis  
    uniCount = 0
    for u in unis:
        u = u.strip()
        uniCount += 1
        uni2schol.setdefault(u, {})[email] = uniCount

    if VERBOSE:
        print(f"\tInterested in {ScholarUni[email][0]} schools: {ScholarUni[email][1:]}")

    # Add scholar information to the Scholars dictionary
    if email not in Scholars:
        nScholars += 1
    Scholars[email] = [name, numInt] + interests
    scholN2E[name] = email
    scholE2N[email] = CapName

    if VERBOSE:
        print(f"Read scholar {CapName} {email} on line {nlines}:")
        for i, info in enumerate(Scholars[email]):
            print(f"\t[{i}]={info}")

print(f"Read {nScholars} unique scholars interested in {len(uni2schol)} Universities:")
for uni in sorted(uni2schol):
    print(f"\t{uni}: {len(uni2schol[uni])}")


# Print the popularity of faculty vs scholar research categories
print("\npopularity of faculty vs scholar research categories:\n")

# Sort the categories by popularity in descending order
sorted_facCatPop = sorted(facCatPop.items(), key=lambda x: x[1], reverse=True)

for cat, fac_count in sorted_facCatPop:
    if fac_count < 1:
        continue
    if cat not in scholCatPop:
        scholCatPop[cat] = 0
    
    print(f"\"{cat}\"\t{fac_count:2d} ({fac_count/nfacCatPop:.4f})\t{scholCatPop[cat]:2d} ({scholCatPop[cat]/nscholCatPop:.4f})")
    #print(f"\"{cat:>42}\"\t{fac_count:2d} ({fac_count/nfacCatPop:.4f})\t{scholCatPop[cat]:2d} ({scholCatPop[cat]/nscholCatPop:.4f})")

# Print the popularity of faculty vs scholar research interests
print("\npopularity of faculty vs scholar research interests:\n")

# Sort the interests by popularity in descending order
sorted_facIntPop = sorted(facIntPop.items(), key=lambda x: x[1], reverse=True)

for interest, fac_count in sorted_facIntPop:
    if fac_count < 10:
        continue
    if interest not in scholIntPop:
        scholIntPop[interest] = 0
    
    print(f"{interest:>42}\t{fac_count:2d} ({fac_count/nfacIntPop:.4f})\t{scholIntPop[interest]:2d} ({scholIntPop[interest]/nscholIntPop:.4f})")

# now build a dictionary to map faculty and scholar overlapping interests, weighted by ranking 
FSintMap = defaultdict(dict) #restricted to faculty at universities the scholar is interested in
SFintMap = defaultdict(dict) #restricted to faculty at universities the scholar is interested in
FSintMapBK = defaultdict(dict) #restricted to faculty at universities that the scholar is not interested in
SFintMapBK = defaultdict(dict) #restricted to faculty at universities that the scholar is not interested in
for fac in Faculty.keys():
    if fac in facCancel:
        continue
    for schol in Scholars.keys():
        FSintMap[fac][schol] = 0
        SFintMap[schol][fac] = 0
        FSintMapBK[fac][schol] = 0
        SFintMapBK[schol][fac] = 0

for interest in facIntPop:
    for fac in int2fac[interest]:
        if interest not in int2schol:
            continue
        for schol in int2schol[interest]:
            #check if faculty is at a university the scholar is interested in
            uniMatch = False
            for u in ScholarUni[schol][1:]:
                if u == Faculty[fac][2]:
                    uniMatch = True
                    break
            if uniMatch:
                FSintMap[fac][schol] += (4-int2fac[interest][fac])*(4-int2schol[interest][schol]);
                SFintMap[schol][fac] += (4-int2fac[interest][fac])*(4-int2schol[interest][schol]);
            else:
                FSintMapBK[fac][schol] += (4-int2fac[interest][fac])*(4-int2schol[interest][schol]);
                SFintMapBK[schol][fac] += (4-int2fac[interest][fac])*(4-int2schol[interest][schol]);

if False: #change this to True to print out the Faculty-Scholar interest map
    for fac in FSintMap.keys():
        for schol in sorted(FSintMap[fac].keys(), key=lambda x: FSintMap[fac][x], reverse=True):
            if FSintMap[fac][schol] <= 10:
                break
            print(f"{fac}:{schol}={FSintMap[fac][schol]}")

# Initialize dictionaries to store faculty choices and availability
FacChoices = {}
nFacChoices = {}
nFacWithChoices = 0
scholarCNT = {}  # Number of times a scholar is selected
nscholarCNT = 0  # Number of unique scholars selected
totScholarCNT = 0  # Total number of scholar selections
nNotAvail = 0  # Number of slots faculty are not available
nTotAvail = 0  # Total number of slots faculty are available for

# File path for the faculty choices data
facChoiceFile = os.path.join(DIR, "Data/Faculty_Choices.xlsx")

# Read the Excel file into a pandas DataFrame
df_fac_choices = pd.read_excel(facChoiceFile, skiprows=3, header=None)  # Skip first 2 line2, header in 3rd line, data starts on 4th line

nlines = 0
for index, row in df_fac_choices.iterrows():
    nlines += 1
    sp = row.tolist()

    email = clean_string(str(sp[3])).strip().lower()
    if DEBUG:
        print(f"working on {email}")

    if email not in Faculty:
        if VERBOSE:
            print(f"did not find Faculty[{email}]...")
        continue

    if email not in nFacChoices:
        nFacWithChoices += 1
    else:
        if VERBOSE:
            print(f"faculty {email} entered choices twice... overwriting")

    nFacChoices[email] = 0
    ## check which timezone was entered
    timezone = clean_string(str(sp[11])).strip().lower()

    SlotAvail = range(12, 19) #default is eastern time, first slot on the datafile... double check columns!
    if timezone == "eastern time":
        if VERBOSE:
            print(f"Timezone for {email}: {timezone}")
    elif timezone == "central time":
        if VERBOSE:
            print(f"Timezone for {email}: {timezone}")
        SlotAvail = range(20, 27)
    elif timezone == "mountain time":
        if VERBOSE:
            print(f"Timezone for {email}: {timezone}")
        SlotAvail = range(28, 35)
    elif timezone == "pacific time":
        if VERBOSE:
            print(f"Timezone for {email}: {timezone}")
        SlotAvail = range(36, 44)
    else:
        print(f"did not find timezone for {email}")
        exit(1)
    
    for i in range(6, 11):
        if len(sp) > i:
            choice = clean_string(str(sp[i])).strip().lower().replace(" ","")
            if choice == "":
                continue
            if DEBUG:
                print(f"\tchoice {i-5}: {choice}")

            if choice in scholN2E:
                FacChoices.setdefault(email, []).append(scholN2E[choice])
                nFacChoices[email] += 1
                if choice not in scholarCNT:
                    nscholarCNT += 1
                scholarCNT[choice] = scholarCNT.get(choice, 0) + 1
                totScholarCNT += 1
            else:
                if DEBUG:
                    print(f"did not find :{choice}: in scholars... Skipping! ({email})")
                    if DEBUGPAUSE:
                        user_entry = input("Is this ok to skip? (y/n)")
                        if user_entry != "y" and user_entry != "":
                            exit(1)
        
    if VERBOSE:
        if nFacChoices[email] > 0:
            print(f"\t{FacChoices[email]} added for faculty {email}")

    # Now check availability
    availSlots = 0
    for i in SlotAvail:
        colid = i - SlotAvail[0]  # map to 0-7
        if len(sp) > colid:
            availability = clean_string(str(sp[i])).strip().lower()
            if VERBOSE:
                print(f"\tChecking availability for slot {colid}/{i} ({availability})")
            if re.search(r'\bno\b', availability, re.IGNORECASE):  # Use word boundary anchors to match exact word "no"
                FacAvail[email][colid] = 0
                nNotAvail += 1
            else:
                availSlots += 1
        else:
            availSlots += 1
        if DEBUG:
            print(f"\tAvail({colid}): {FacAvail[email][colid]}")

    if availSlots == 0:
        facCancel[email] = 1
        del Faculty[email]
        nFacWithChoices -= 1
        if VERBOSE:
            print(f"faculty {email} not available at all...")
    else:
        nTotAvail += availSlots

if VERBOSE:
    i = 0
    for s in sorted(scholarCNT, key=scholarCNT.get, reverse=True):
        i += 1
        print(f"{i}: {s} -> {scholarCNT[s]}")

NEWnFac = 0
for fac in Faculty:
    if fac not in facCancel:
        NEWnFac += 1

print(f"There are now {NEWnFac} faculty able to participate")
print(f"{nFacWithChoices} faculty chose {totScholarCNT} scholars ({nscholarCNT} unique) and are not available for {nNotAvail} slots out of {nNotAvail+nTotAvail} ({nNotAvail/(nNotAvail+nTotAvail)})")

# Initialize dictionaries to store scholar choices and related counts
ScholChoices = {}
nScholChoices = {}
nScholWithChoices = 0
facultyCNT = {}  # Number of times faculty selected
nfacultyCNT = 0  # Number of unique faculty selected
totFacultyCNT = 0  # Total number of faculty selections

# File path for the scholar choices data
scholChoiceFile = os.path.join(DIR, "Data/Scholar_Choices.xlsx")

# Read the Excel file into a pandas DataFrame
df_schol_choices = pd.read_excel(scholChoiceFile, skiprows=1, header=None)  # Skip first line, header on 1st line, data starts on 2nd line

nlines = 0
for index, row in df_schol_choices.iterrows():
    nlines += 1
    sp = row.tolist()

    email = clean_string(str(sp[3])).strip().lower()
    if DEBUG:
        print(f"working on {email}")

    if email not in Scholars:
        if VERBOSE:
            print(f"did not find Scholar[{email}]... on line {nlines}")
            if DEBUGPAUSE:
                user_entry = input("Continue? (y/n)")
                if user_entry != "y" and user_entry != "":
                    exit(1)
        continue

    if email not in nScholChoices:
        nScholWithChoices += 1
    else:
        if VERBOSE:
            print(f"Scholar {email} entered info twice! Overwritting...")
            if DEBUGPAUSE:
                user_entry = input("Continue? (y/n)")
                if user_entry != "y" and user_entry != "":
                    exit(1)


    nScholChoices[email] = 0
    for i in range(5, 9):
        if len(sp) > i:
            name = clean_string(str(sp[i])).split(" - ")[0].strip().lower().replace(" ","")

            if name == "" or name not in facultyNAMEtoEMAIL:
                continue
            if facultyNAMEtoEMAIL[name] in facCancel:
                continue

            if DEBUG:
                print(f"\tchoice {i-5}: {name}")

            if name in facultyNAMEtoEMAIL and facultyNAMEtoEMAIL[name] in Faculty:
                ScholChoices.setdefault(email, []).append(facultyNAMEtoEMAIL[name])
                nScholChoices[email] += 1
                if facultyNAMEtoEMAIL[name] not in facultyCNT:
                    nfacultyCNT += 1
                facultyCNT[facultyNAMEtoEMAIL[name]] = facultyCNT.get(facultyNAMEtoEMAIL[name], 0) + 1
                totFacultyCNT += 1
            else:
                if VERBOSE:
                    print(f"did not find faculty :{name}: in scholars interest list... ({email})")
                if DEBUGPAUSE:
                    response = input("Is this ok to skip? (y/n)")
                    if response == "y" or response == "":
                        continue
                    else:
                        quit()

print(f"{nScholWithChoices} scholars chose {totFacultyCNT} faculty ({nfacultyCNT} unique)")

if VERBOSE:
    i = 0
    for s in sorted(facultyCNT, key=facultyCNT.get, reverse=True):
        i += 1
        print(f"{i}: {s} -> {facultyCNT[s]}")

# Now check which faculty were not chosen by any scholar
nFacNotChosen = 0
for fac in Faculty:
    if fac in facCancel:
        continue
    if VERBOSE:
        print(f"checking faculty {fac}: {facultyEMAILtoNAME[fac]}...")
    if fac not in facultyCNT:
        nFacNotChosen += 1
        print(f"faculty {facultyEMAILtoNAME[fac]} was not chosen by any scholar")
print(f"{nFacNotChosen} faculty were not chosen by any scholar")
input("Press enter to continue")

# Now build schedules!!

# Initialize variables and dictionaries for schedules
FAILED = 1
TOTTRIES = 0
MAXTRIES = 100
MINScholInt = 3  # minimum interviews for each scholar
MINFacInt = 3  # minimum interviews for each faculty
MAXInt = 8  # maximum number of interview slots

FACsched = {}
FACschedWhy = {} # store why a slot was filled: FC=faculty choice; SC=scholar choice; RA=random assignment
nFACsched = {}
FACslots = {}
nFACslots = {}
SCHOLsched = {}
SCHOLschedWhy = {} # store why a slot was filled: FC=faculty choice; SC=scholar choice; RA=random assignment
nSCHOLsched = {}
SCHOLslots = {}
nSCHOLslots = {}

def check_schedules():
    minINTS = 8
    nFacFull = 0
    nFacLow = 0
    nFacAvail = 0
    nScholFull = 0
    nScholLow = 0
    nScholAvail = 0
    nScheduled = 0
    for fac in nFACsched:
        if fac in facCancel:
            continue
        if nFACsched[fac] < minINTS:
            minINTS = nFACsched[fac]
        if nFACsched[fac] < MINFacInt:
            nFacLow += 1
        if nFACsched[fac] == MAXInt:
            nFacFull += 1
        else:
            nFacAvail += 1
    
    print(f"{nFacFull} faculty with full schedules; {nFacLow} faculty with too few; as low as {minINTS}; {nFacAvail} still have availability")
    minINTS = 8
    for schol in nSCHOLsched:
        nScheduled += nSCHOLsched[schol]
        if nSCHOLsched[schol] < minINTS:
            minINTS = nSCHOLsched[schol]
        if nSCHOLsched[schol] < MINScholInt:
            nScholLow += 1
        if nSCHOLsched[schol] == MAXInt:
            nScholFull += 1
        else:
            nScholAvail += 1

    print(f"{nScholFull} scholars with full schedules; {nScholLow} scholars with too few; as low as {minINTS}; {nScholAvail} still have availability")
    print(f"total of {nScheduled} interviews scheduled")

MINtooFewInts = 1000 #store how many have too few interviews
while FAILED: #this block will rerun everything if it fails, re-initialiing all variables/dictionaries
    TOTTRIES += 1
    if VERBOSE:
        print(f"Try number {TOTTRIES}")
        if DEBUGPAUSE:
            user_entry = input("Continue? [y/n]")
            if user_entry != 'y' and user_entry != "":
                exit(1)

    # Reset schedules
    FACsched = {}
    FACschedWhy = {}
    nFACsched = {}
    FACslots = {}
    nFACslots = {}
    SCHOLsched = {}
    SCHOLschedWhy = {}
    nSCHOLsched = {}
    SCHOLslots = {}
    nSCHOLslots = {}
    F2S = {}
    S2F = {}

    # Initialize faculty schedules
    for fac in Faculty:
        if fac in facCancel:
            continue
        nFACslots[fac] = MAXInt  # default to max available
        nFACsched[fac] = 0
        FACsched[fac] = []
        FACschedWhy[fac] = []

        if DEBUG:
            print(f"initializing faculty {fac}")
        for i in range(MAXInt):
            if FacAvail[fac][i] == 0:
                FACsched[fac].append("NA")
                FACschedWhy[fac].append("NA")
                nFACslots[fac] -= 1
            else:
                FACsched[fac].append("")
                FACschedWhy[fac].append("")
                FACslots.setdefault(fac, {})[i] = 1
    
    # Initialize scholar schedules
    for schol in Scholars:
        nSCHOLsched[schol] = 0
        nSCHOLslots[schol] = MAXInt
        SCHOLsched[schol] = ["" for _ in range(MAXInt)]
        SCHOLschedWhy[schol] = ["" for _ in range(MAXInt)]
        for i in range(MAXInt):
            SCHOLslots.setdefault(schol, {})[i] = 1


    # First start with faculty interests
    facmiss = {}
    nfacmiss = 0
    totfacmiss = 0
    
    for i in range(5):
        facOrd = randomize_order(Faculty.keys())
        for fac in facOrd:
            if fac not in FacChoices:
                continue
            if i < len(FacChoices[fac]):
                schol = FacChoices[fac][i]
                if schol not in Scholars:
                    continue
                if nSCHOLsched[schol] < MAXInt:
                    foundMatch = 0
                    numbers = list(range(MAXInt))
                    random.shuffle(numbers)
                    for try_slot in numbers:
                        if try_slot in SCHOLslots[schol] and try_slot in FACslots[fac]:
                            FACsched[fac][try_slot] = schol
                            FACschedWhy[fac][try_slot] = "FC"
                            SCHOLsched[schol][try_slot] = fac
                            SCHOLschedWhy[schol][try_slot] = "FC"
                            del FACslots[fac][try_slot]
                            nFACslots[fac] -= 1
                            nFACsched[fac] += 1
                            del SCHOLslots[schol][try_slot]
                            nSCHOLslots[schol] -= 1
                            nSCHOLsched[schol] += 1
                            F2S.setdefault(fac, {})[schol] = True
                            S2F.setdefault(schol, {})[fac] = True
                            foundMatch = 1
                            break
                    if foundMatch == 0:
                        if VERBOSE:
                            print(f"{schol} can't match with {fac}, nSCHOLsched={nSCHOLsched[schol]}, nFACsched={nFACsched[fac]}")
                        if fac not in facmiss:
                            nfacmiss += 1
                        facmiss[fac] = facmiss.get(fac, 0) + 1
                        totfacmiss += 1

    if VERBOSE:
        print(f"{nfacmiss} faculty didn't get a slot with one of their top choices, {totfacmiss} overall:")
        for fac in facmiss:
            print(f"\t{fac}")

    print("Incorporated faculty choices")
    check_schedules()

    # Now fill in schedule from scholar interests in the same way, but start with scholars with fewest interviews
    scholmiss = {}
    nrecipChoice = 0
    for i in range(5):
        #sort scholars with fewest interviews first
        scholOrd = sorted(Scholars.keys(), key=lambda x: nSCHOLsched[x])
        #scholOrd = randomize_order(Scholars.keys())
        for schol in scholOrd:
            if schol not in ScholChoices:
                continue
            if i < len(ScholChoices[schol]):
                fac = ScholChoices[schol][i]
                if fac not in Faculty or fac in facCancel:
                    continue
                if nFACslots[fac] == 0:
                    continue
                if fac in SCHOLsched[schol]: #already have an interview!
                    nrecipChoice += 1
                    continue
                if nFACslots[fac] > 0:
                    foundMatch = 0
                    numbers = list(range(MAXInt))
                    random.shuffle(numbers)
                    for try_slot in numbers:
                        if try_slot in SCHOLslots[schol] and try_slot in FACslots[fac]:
                            FACsched[fac][try_slot] = schol
                            FACschedWhy[fac][try_slot] = "SC"
                            SCHOLsched[schol][try_slot] = fac
                            SCHOLschedWhy[schol][try_slot] = "SC"
                            del FACslots[fac][try_slot]
                            nFACslots[fac] -= 1
                            nFACsched[fac] += 1
                            del SCHOLslots[schol][try_slot]
                            nSCHOLslots[schol] -= 1
                            nSCHOLsched[schol] += 1
                            F2S.setdefault(fac, {})[schol] = True
                            S2F.setdefault(schol, {})[fac] = True
                            foundMatch = 1
                            break

                    if foundMatch == 0:
                        if VERBOSE:
                            print(f"{schol} can't match with {fac}, nFACsched={nFACsched[fac]}")
                            print(f"schol {schol}:{SCHOLsched[schol]}")
                            print(f"fac {fac}:{FACsched[fac]}")

                        if schol not in scholmiss:
                            scholmiss[schol] = 0
                        scholmiss[schol] += 1

    if VERBOSE:
        print(f"{len(scholmiss)} scholars didn't match with one of their top choices, {sum(scholmiss.values())} overall")
    
    print(f"Incorporated scholar choices. {nrecipChoice} recipricol choice interviews")
    check_schedules()

    # Randomly fill in remaining slots according to interests, starting with scholars, based on university of interest, and those with fewest interviews
    for TRIES in range(MAXInt):
        #sort scholars with fewest interviews first
        for schol in sorted(Scholars.keys(), key=lambda x: nSCHOLsched[x]):
            if nSCHOLsched[schol] > TRIES: #skip those who already have more than TRIES interviews this round
                continue

            for fac in sorted(SFintMap[schol].keys(), key=lambda x: SFintMap[schol][x], reverse=True):     
                if fac in facCancel:
                    continue
                if fac in SCHOLsched[schol] or nFACslots[fac] == 0:
                    continue
                found_match = False
                numbers = list(range(MAXInt))
                random.shuffle(numbers)
                for try_slot in numbers:
                    if try_slot in SCHOLslots[schol] and try_slot in FACslots[fac]:
                        FACsched[fac][try_slot] = schol
                        FACschedWhy[fac][try_slot] = "RA"
                        SCHOLsched[schol][try_slot] = fac
                        SCHOLschedWhy[schol][try_slot] = "RA"
                        del FACslots[fac][try_slot]
                        nFACslots[fac] -= 1
                        nFACsched[fac] += 1
                        del SCHOLslots[schol][try_slot]
                        nSCHOLslots[schol] -= 1
                        nSCHOLsched[schol] += 1
                        F2S.setdefault(fac, {})[schol] = True
                        S2F.setdefault(schol, {})[fac] = True
                        found_match = True
                        break
                if found_match:
                    break

    print(f"Added random matches based on interests: Scholars")
    check_schedules()

    for fac in sorted(nFACsched.keys(), key=lambda x: nFACsched[x]):
        if nFACsched[fac] > MINFacInt:
            continue
        if fac in facCancel:
            continue
        for schol in sorted(FSintMap[fac].keys(), key=lambda x: FSintMap[fac][x], reverse=True):
            if schol in FACsched[fac] or nSCHOLslots[schol] == 0:
                continue
            if VERBOSE:
                print(f"{fac}={nFACslots[fac]}; {schol}={nSCHOLslots[schol]}")
            numbers = list(range(MAXInt))
            random.shuffle(numbers)
            for try_slot in numbers:
                if try_slot in SCHOLslots[schol] and try_slot in FACslots[fac]:
                    FACsched[fac][try_slot] = schol
                    FACschedWhy[fac][try_slot] = "RA"
                    SCHOLsched[schol][try_slot] = fac
                    SCHOLschedWhy[schol][try_slot] = "RA"
                    del FACslots[fac][try_slot]
                    nFACslots[fac] -= 1
                    nFACsched[fac] += 1
                    del SCHOLslots[schol][try_slot]
                    nSCHOLslots[schol] -= 1
                    nSCHOLsched[schol] += 1
                    F2S.setdefault(fac, {})[schol] = True
                    S2F.setdefault(schol, {})[fac] = True
                    break
            if nFACsched[fac] > MINFacInt:
                break

    print(f"Added random matches based on interests: faculty")
    check_schedules()

    ## now lets add more random matches, ignoring university of interest for scholars. This will use the backup interest map
    for TRIES in range(MAXInt):
        for schol in sorted(Scholars.keys(), key=lambda x: nSCHOLsched[x]):
            if nSCHOLsched[schol] > TRIES:
                continue

            for fac in sorted(SFintMapBK[schol].keys(), key=lambda x: SFintMapBK[schol][x], reverse=True):            
                if fac in facCancel:
                    continue
                if fac in SCHOLsched[schol] or nFACslots[fac] == 0:
                    continue
                found_match = False
                numbers = list(range(MAXInt))
                random.shuffle(numbers)
                for try_slot in numbers:
                    if try_slot in SCHOLslots[schol] and try_slot in FACslots[fac]:
                        FACsched[fac][try_slot] = schol
                        FACschedWhy[fac][try_slot] = "RA"
                        SCHOLsched[schol][try_slot] = fac
                        SCHOLschedWhy[schol][try_slot] = "RA"
                        del FACslots[fac][try_slot]
                        nFACslots[fac] -= 1
                        nFACsched[fac] += 1
                        del SCHOLslots[schol][try_slot]
                        nSCHOLslots[schol] -= 1
                        nSCHOLsched[schol] += 1
                        F2S.setdefault(fac, {})[schol] = True
                        S2F.setdefault(schol, {})[fac] = True
                        found_match = True
                        break
                if found_match:
                    break
    print(f"Added random matches based on interests including universities Scholars did not select")
    check_schedules()

    # Check how many interviews everyone got
    facinthist = [0] * (MAXInt + 1)
    scholinthist = [0] * (MAXInt + 1)
    facfillhist = [0] * (MAXInt + 1)
    factoofew = {}
    scholtoofew = {}

    for fac in Faculty:
        if fac in facCancel:
            continue
        facinthist[nFACsched[fac]] += 1
        empty = MAXInt - nFACslots[fac]
        facfillhist[empty] += 1
        if nFACsched[fac] < MINFacInt:
            factoofew[fac] = True

    for schol in Scholars:
        scholinthist[nSCHOLsched[schol]] += 1
        if nSCHOLsched[schol] < MINScholInt:
            scholtoofew[schol] = True

    if VERBOSE:
        print("\nN\t#FacInt\t#Schols")
        for i in range(MAXInt + 1):
            print(f"{i}\t{facfillhist[i]}\t{scholinthist[i]}")
        print(f"\n\n{len(factoofew)} faculty and {len(scholtoofew)} scholars with too few interviews")
        print("Faculty with too few interviews:")
        for fac in sorted(factoofew):
            print(f"\t{fac}")
        print("Scholars with too few interviews:")
        for schol in sorted(scholtoofew):
            print(f"\t{schol}")

    if (len(factoofew) > 0 or len(scholtoofew) > 0) and TOTTRIES < MAXTRIES:
        if VERBOSE:
            print(f"FAILED! factoofew={len(factoofew)}; scholtoofew={len(scholtoofew)}")
        FAILED = 1  # Try again
    #else:
        #FAILED = 0  # Done

        if VERBOSE:
            print(f"\n\n{len(factoofew)} faculty and {len(scholtoofew)} scholars with too few interviews")
            print("Faculty with too few interviews:")
            for fac in sorted(factoofew):
                print(f"\t{fac}")
            print("Scholars with too few interviews:")
            for schol in sorted(scholtoofew):
                print(f"\t{schol}")
            print("N\t#Faculty\t#Scholars")
            for i in range(MAXInt + 1):
                print(f"{i}\t{facfillhist[i]}\t{scholinthist[i]}")


        # Check how many interviews everyone got
        facinthist = [0] * (MAXInt + 1)
        scholinthist = [0] * (MAXInt + 1)
        facfillhist = [0] * (MAXInt + 1)
        factoofew = {}
        scholtoofew = {}

        for fac in Faculty:
            facinthist[nFACsched[fac]] += 1
            empty = MAXInt - nFACslots[fac]
            facfillhist[empty] += 1
            if nFACsched[fac] < MINFacInt:
                factoofew[fac] = True

        for schol in Scholars:
            scholinthist[nSCHOLsched[schol]] += 1
            if nSCHOLsched[schol] < MINScholInt:
                scholtoofew[schol] = True
        
        if VERBOSE:
            print(f"\n\n{len(factoofew)} faculty and {len(scholtoofew)} scholars with too few interviews")
            print("Faculty with too few interviews:")
            for fac in sorted(factoofew):
                print(f"\t{fac}")
            print("Scholars with too few interviews:")
            for schol in sorted(scholtoofew):
                print(f"\t{schol}")
            print("\nHere is a table showing the number of faculty and scholars who have N interviews:")
            print("\nN\t#Faculty\t#Scholars")
            for i in range(MAXInt + 1):
                print(f"{i}\t{facfillhist[i]:8d}\t{scholinthist[i]:9d}")
        
    if len(factoofew) + len(scholtoofew) < MINtooFewInts:
        MINtooFewInts = len(factoofew) + len(scholtoofew)
        if TOTTRIES >= MAXTRIES:
            print(f"Try number {TOTTRIES}, faculty with too few interviews={len(factoofew)}; scholars with too few={len(scholtoofew)}; total={MINtooFewInts}")
            print("\nHere is a table showing the number of faculty and scholars who have N interviews:")
            print("\nN\t#Fac\t#Scholars")
            #format so output columns line up by forcing entries to have defined width
            for i in range(MAXInt + 1):
                print(f"{i}\t{facfillhist[i]:8d}\t{scholinthist[i]:9d}")

            answer = input("Try again? (y/n)")
            if answer == "n":
                FAILED = 0
                break
            MAXTRIES += 20
    #input("Press enter to continue...")



print("Made a schedule...")

# Let's print out the schedules!

# Create a Pandas DataFrame for the schedules and save to .xlsx files
totScholInts = 0
totFacInts = 0
GAPCHR = ","

try:
    # Faculty schedule: print with scholar names on first sheet, and reasons why on second sheet: FC=faculty choice; SC=scholar choice; RA=random assignment
    facSchedFile = os.path.join(DIROUT, "CompleteFacultySchedule.xlsx")
    with pd.ExcelWriter(facSchedFile, engine='xlsxwriter') as writer:
        fac_schedule = []
        fac_schedule_why = []
        for fac in sorted(Faculty.keys(), key=lambda x: facultyEMAILtoNAME[x]):
            if fac in facCancel:
                continue
            row = [facultyEMAILtoNAME[fac]]
            row_why = [facultyEMAILtoNAME[fac]]
            numInt = 0
            for i in range(MAXInt):
                schol = FACsched[fac][i]
                reason = FACschedWhy[fac][i]
                if schol == "NA" or schol == "":
                    row.append("NA")
                    row_why.append("NA")
                else:
                    row.append(scholE2N[schol])
                    row_why.append(f"{reason}")
                    totFacInts += 1
                    numInt += 1
            fac_schedule.append(row)
            fac_schedule_why.append(row_why)
        df_fac_schedule = pd.DataFrame(fac_schedule, columns=["FacultyName"] + TIMES)
        df_fac_schedule.to_excel(writer, index=False, sheet_name="FacultySchedule")

        df_fac_schedule_why = pd.DataFrame(fac_schedule_why, columns=["FacultyName"] + TIMES)
        df_fac_schedule_why.to_excel(writer, index=False, sheet_name="AssignmentReasons")

        # Access the xlsxwriter workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets["FacultySchedule"]

        # Set column widths to match header lengths (with small padding)
        for col_idx, col_name in enumerate(df_fac_schedule.columns):
            width = max(len(str(col_name)), 8) + 2
            worksheet.set_column(col_idx, col_idx, width)
        # Left-align header row for readability
        header_format = workbook.add_format({"align": "left"})
        worksheet.set_row(0, None, header_format)
        
        if "AssignmentReasons" in writer.sheets:
            ws_reasons = writer.sheets["AssignmentReasons"]
            ws_reasons.set_row(0, None, header_format)
            #also adjust widths for the AssignmentReasons sheet
            for col_idx, col_name in enumerate(df_fac_schedule_why.columns):
                width = max(len(str(col_name)), 8) + 2
                ws_reasons.set_column(col_idx, col_idx, width)

        # Add the link to the meeting links spreadsheet
        worksheet.write_url(len(df_fac_schedule) + 2, 0, "https://docs.google.com/spreadsheets/d/1BByK0i4PcavbzGxVeRPs66lmdhZwxXjQNVwW0iQpCV4/edit?usp=sharing", string='Please find all meeting links here')

    print(f"There are a total of {totFacInts} faculty interviews")
except Exception as e:
    print(f"Failed to create faculty schedule. Error: {e}")



# Scholar schedule, also with reasons why on second sheet
try:
    scholSchedFile = os.path.join(DIROUT, "CompleteScholarSchedule.xlsx")
    with pd.ExcelWriter(scholSchedFile, engine='xlsxwriter') as writer:
        schol_schedule = []
        schol_schedule_why = []
        for schol in sorted(Scholars.keys()):
            row = [scholE2N[schol]]
            row_why = [scholE2N[schol]]
            numInt = 0
            for i in range(MAXInt):
                fac = SCHOLsched[schol][i]
                reason = SCHOLschedWhy[schol][i]
                if fac not in facultyEMAILtoNAME:
                    row.append("NA")
                    row_why.append("NA")
                else:
                    row.append(facultyEMAILtoNAME[fac])
                    row_why.append(f"{reason}")
                    totScholInts += 1
                    numInt += 1
            schol_schedule.append(row)
            schol_schedule_why.append(row_why)
        df_schol_schedule = pd.DataFrame(schol_schedule, columns=["ScholarName"] + TIMES)
        df_schol_schedule.to_excel(writer, index=False, sheet_name="ScholarSchedule")

        df_schol_schedule_why = pd.DataFrame(schol_schedule_why, columns=["ScholarName"] + TIMES)
        df_schol_schedule_why.to_excel(writer, index=False, sheet_name="AssignmentReasons")

        # Access the xlsxwriter workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets["ScholarSchedule"]
        
        # Set column widths to match header lengths (with small padding)
        for col_idx, col_name in enumerate(df_schol_schedule.columns):
            width = max(len(str(col_name)), 8) + 2
            worksheet.set_column(col_idx, col_idx, width)
        #left-align header row for readability
        header_format = workbook.add_format({"align": "left"})
        worksheet.set_row(0, None, header_format)

        # Also adjust widths for the AssignmentReasons sheet
        if "AssignmentReasons" in writer.sheets:
            ws_reasons = writer.sheets["AssignmentReasons"]
            for col_idx, col_name in enumerate(df_schol_schedule_why.columns):
                width = max(len(str(col_name)), 8) + 2
                ws_reasons.set_column(col_idx, col_idx, width)
            # Left-align header row for readability
            header_format = workbook.add_format({"align": "left"})
            ws_reasons.set_row(0, None, header_format)

        # Left-align header row for the ScholarSchedule sheet
        header_format = workbook.add_format({"align": "left"})
        worksheet.set_row(0, None, header_format)

        # Add the link to the meeting links spreadsheet
        worksheet.write_url(len(df_schol_schedule) + 2, 0, "https://docs.google.com/spreadsheets/d/1BByK0i4PcavbzGxVeRPs66lmdhZwxXjQNVwW0iQpCV4/edit?usp=sharing", string='Please find all meeting links here')

    print(f"There are a total of {totScholInts} scholar interviews")
except Exception as e:
    print(f"Failed to create scholar schedule. Error: {e}")

# Print out individual scholar schedules
try:
    scholar_schedules_dir = os.path.join(DIROUT, "ScholarSchedules")
    if not os.path.exists(scholar_schedules_dir):
        os.makedirs(scholar_schedules_dir)

    for schol in Scholars.keys():
        scholSchedFile = os.path.join(scholar_schedules_dir, f"{scholE2N[schol]}.xlsx")
        with pd.ExcelWriter(scholSchedFile, engine='xlsxwriter') as writer:
            scholar_schedule = [[scholE2N[schol]] + [facultyEMAILtoNAME[SCHOLsched[schol][i]] if SCHOLsched[schol][i] in facultyEMAILtoNAME else "NA" for i in range(MAXInt)]]
            df_scholar_schedule = pd.DataFrame(scholar_schedule, columns=["ScholarName"] + TIMES)
            df_scholar_schedule.to_excel(writer, index=False, sheet_name="ScholarSchedule")

            # Access the xlsxwriter workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets["ScholarSchedule"]

            # Set column widths to match header lengths (with small padding)
            for col_idx, col_name in enumerate(df_scholar_schedule.columns):
                width = max(len(str(col_name)), 8) + 2
                worksheet.set_column(col_idx, col_idx, width)

            # Left-align header row for readability
            header_format = workbook.add_format({"align": "left"})
            worksheet.set_row(0, None, header_format)

            # Add the link to the meeting links spreadsheet
            worksheet.write_url(len(df_scholar_schedule) + 2, 0, "https://docs.google.com/spreadsheets/d/1BByK0i4PcavbzGxVeRPs66lmdhZwxXjQNVwW0iQpCV4/edit?usp=sharing", string='Please find all meeting links here')

            # Add list of other faculty of interest who were not scheduled
            unscheduled_faculty = []
            if VERBOSE:
                print(f"Checking for unscheduled faculty for scholar {scholE2N[schol]} {schol}...")
            if schol in ScholChoices:
                for fac in ScholChoices[schol]:
                    if fac not in SCHOLsched[schol]:
                        unscheduled_faculty.append(fac)
                        if VERBOSE:
                            print(f"\tFound scholar choice faculty of interest: {fac}")
                # now add faculty who selected the scholar but were not scheduled
            for fac in FacChoices:
                if schol not in FacChoices[fac]:
                    continue
                if fac not in SCHOLsched[schol] and fac not in unscheduled_faculty:
                    unscheduled_faculty.append(fac)
                    if VERBOSE:
                        print(f"\tFound faculty who selected scholar but were not scheduled: {fac}")
            if unscheduled_faculty:
                worksheet.write(len(df_scholar_schedule) + 4, 0, "Faculty of Interest Not Scheduled:")
                #print alphabetically sorted list of faculty names and emails
                unscheduled_faculty.sort(key=lambda x: facultyEMAILtoNAME[x])
                for idx, fac in enumerate(unscheduled_faculty):
                    worksheet.write(len(df_scholar_schedule) + 5 + idx, 0, facultyEMAILtoNAME[fac])
                    worksheet.write(len(df_scholar_schedule) + 5 + idx, 1, fac)
            else:
                if VERBOSE:
                    print(f"\tNo unscheduled faculty of interest found for scholar {scholE2N[schol]} {schol}")
except Exception as e:
    print(f"Failed to create individual scholar schedule for {schol}. Error: {e}")

# Print out individual faculty schedules
try:
    faculty_schedules_dir = os.path.join(DIROUT, "FacultySchedules")
    if not os.path.exists(faculty_schedules_dir):
        os.makedirs(faculty_schedules_dir)

    for fac in Faculty.keys():
        if fac in facCancel:
            continue
        facSchedFile = os.path.join(faculty_schedules_dir, f"{facultyEMAILtoNAME[fac]}.xlsx")
        with pd.ExcelWriter(facSchedFile, engine='xlsxwriter') as writer:
            faculty_schedule = [[facultyEMAILtoNAME[fac]] + [scholE2N[FACsched[fac][i]] if FACsched[fac][i] in scholE2N else "NA" for i in range(MAXInt)]]
            df_faculty_schedule = pd.DataFrame(faculty_schedule, columns=["FacultyName"] + TIMES)
            df_faculty_schedule.to_excel(writer, index=False, sheet_name="FacultySchedule")

            # Access the xlsxwriter workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets["FacultySchedule"]

            # Set column widths to match header lengths (with small padding)
            for col_idx, col_name in enumerate(df_faculty_schedule.columns):
                width = max(len(str(col_name)), 8) + 2
                worksheet.set_column(col_idx, col_idx, width)

            # Left-align header row for readability
            header_format = workbook.add_format({"align": "left"})
            worksheet.set_row(0, None, header_format)

            # Add the link to the meeting links spreadsheet
            worksheet.write_url(len(df_faculty_schedule) + 2, 0, "https://docs.google.com/spreadsheets/d/1BByK0i4PcavbzGxVeRPs66lmdhZwxXjQNVwW0iQpCV4/edit?usp=sharing", string='Please find all meeting links here')

            # now add list of scholars who selected the faculty or selected by faculty but were not scheduled
            unscheduled_scholars = []
            if VERBOSE:
                print(f"Checking for unscheduled scholars for faculty {facultyEMAILtoNAME[fac]} {fac}...")
            if fac in FacChoices:
                for schol in FacChoices[fac]:
                    if schol not in FACsched[fac]:
                        unscheduled_scholars.append(schol)
                        if VERBOSE:
                            print(f"\tFound faculty choice scholar of interest: {schol}")
            # now add scholars who selected the faculty but were not scheduled
            for schol in ScholChoices:
                if fac not in ScholChoices[schol]:
                    continue
                if schol not in FACsched[fac] and schol not in unscheduled_scholars:
                    unscheduled_scholars.append(schol)
                    if VERBOSE:
                        print(f"\tFound scholar who selected faculty but were not scheduled: {schol}")
            if unscheduled_scholars:
                worksheet.write(len(df_faculty_schedule) + 4, 0, "Scholars of Interest Not Scheduled:")
                #print alphabetically sorted list of scholar names and emails
                unscheduled_scholars.sort(key=lambda x: scholE2N[x])
                for idx, schol in enumerate(unscheduled_scholars):
                    worksheet.write(len(df_faculty_schedule) + 5 + idx, 0, scholE2N[schol])
                    worksheet.write(len(df_faculty_schedule) + 5 + idx, 1, schol)
            else:
                if VERBOSE:
                    print(f"\tNo unscheduled scholars of interest found for faculty {facultyEMAILtoNAME[fac]} {fac}")

except Exception as e:
    print(f"Failed to create individual faculty schedules. Error: {e}")

# # Print final meeting link list
# try:
#     facmeeting = os.path.join(DIROUT, "facMeetingLinks.xlsx")
#     with pd.ExcelWriter(facmeeting, engine='xlsxwriter') as writer:
#         meeting_links = [[facultyEMAILtoNAME[fac], fac, Faculty[fac][1]] for fac in sorted(Faculty.keys())]
#         df_meeting_links = pd.DataFrame(meeting_links, columns=["Faculty", "Email", "MeetingLink"])
#         df_meeting_links.to_excel(writer, index=False, sheet_name="MeetingLinks")
# except Exception as e:
#     print(f"Failed to create meeting link list. Error: {e}")

print("Schedules and meeting links have been saved.")

## Changes needed:
# break the big block of code down into more functions followed by a main function
# when adding interviews based on interests, only add for individuals with the fewest interviews (e.g. all the 2s, then all the 3s, etc)
# possibly emphasize scholars with least research experience?
# pull out the scholars who are selected by the most faculty vs the least?
# report the scholars and faculty who requested eachother, but didn't get paired together. Put this in individual schedules.
