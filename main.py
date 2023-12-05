#!usr/bin/env Python3


"""
Importing the following modules:
OS for directory navigation, re for Regular Expression checks, shutil to create archives/remove folders,
sys for a clean exit out of the program, date to get to today's date for labeling purposes
"""
import os
import re
import shutil
import sys
from datetime import date


"""
Function clean_transcripts
Variable id_reg_ex creates a regular expression pattern looking for a 3 letter, 3 number combination in the files
os.chdir changes the working directory to the transcripts folder, as this is where we will be doing our modifications
Iterating over the items in the directory, we look for files that contain UTC's ID pattern of 3 letter/3 number and end
with the .pdf file extension.
If the item does not have a UTC ID attached to it, the program first checks to see if a NOID folder exists. If it does,
the item is moved to the folder. If not, the folder is first created and then the item is moved.
Additionally, if a file exists within the NOID folder that has the same name, document type, and ID, the program with
append a '2' on the file in order to differentiate it.
"""
def clean_transcripts():
    # create the UTC ID matching pattern
    id_reg_ex = re.compile(r'\w{3}\d{3}')

    # change the working directory to where the transcripts are located
    os.chdir("c:\\users\\cjj714\\desktop\\transcripts")

    # iterate over the files in the 'transcripts' folder
    for item in os.listdir():

        try:

            # if the file doesn't have a UTC and it ends with the .pdf extension, move it to the 'NOID' folder
            if not id_reg_ex.search(item) and item.endswith(".pdf"):

                # check to see if the 'NOID' folder exists. If not, create it.
                if not os.path.isdir("NOID"):
                    print("No NOID folder! Creating one...")
                    os.mkdir("NOID")

                # after the item is moved, display that it was moved. Ex: "Moved: Thomas Helms - CJJ714 - College.pdf to NOID"
                print("Moved: ", item, " to NOID.")

                # if a file exists in the NOID folder that has the same name, append a '2' to the name to differentiate it.
                # TODO: Update the duplicate file renaming to be able to go beyond just '2' copies (3, 4, etc...)
                if os.path.exists("c:\\users\\cjj714\\desktop\\transcripts\\noid\\" + item):
                    broken = item.split()
                    broken[-2] += "(2)"
                    fixed = " ".join(broken)
                    shutil.move(item, "c:\\users\\cjj714\\desktop\\transcripts\\noid\\" + fixed)
                else:
                    shutil.move(item, "c:\\users\\cjj714\\desktop\\transcripts\\noid\\")
                    
        except Exception as e:
                print("An occurred while moving the file:", str(e))

                
"""
---DEFUNCT---
Function prep_noid
This function is run at the end of the day and creates archives of the NOID folder (for send out via Teams) and the Recommendations
folder (for send out via email).
variable today gets today's date and time
variable day formats today specifically for date only, in the day(numberical) month(abbv name) year format. ex: 16Feb2022
"""
def prep_noid():
    # get today's date and formate is as ex: 16Feb2022
    today = date.today()
    day = today.strftime("%d%b%Y")

    # create a folder with today's date as the name. This will be the parent folder for all of today's documents.
    try:
        os.mkdir("c:\\users\\cjj714\\desktop\\transcripts\\" + day)
        
    # if folder already exits, contiue
    except Exception:
        print("That folder already exits. All good.")

    # from the NOID folder, creates an archive labeled with today's date and NOID. ex: 16Feb2022NOID.zip
    # And then empties and removes the source folder.
    try:
        shutil.make_archive("c:\\users\\cjj714\\desktop\\transcripts\\" + day + "NOID",
                            "zip", "c:\\users\\cjj714\\desktop\\transcripts\\noid\\")
        shutil.rmtree("c:\\users\\cjj714\\desktop\\transcripts\\noid\\")
        print("NOID Zip created and source folder cleared.")
        
    # in the rare case we don't have any NOIDs, print the message and continue
    except Exception:
        print("NOID folder doesn't exists. Oh well, carry on.")

    # from the Recommendations folder, creates an archive labeled with today's date and Recommendations. ex: 16Feb2022Recommendations.zip
    # And then empties and removes the source folder.
    try:
        shutil.make_archive("c:\\users\\cjj714\\desktop\\transcripts\\" + day + "Recommendations", "zip",
                            "c:\\users\\cjj714\\desktop\\transcripts\\Recommendations\\")
        shutil.rmtree("c:\\users\\cjj714\\desktop\\transcripts\\Recommendations\\")
        
    # in the rare case that the Recommendations folder can't be removed (which isn't a big deal), continue
    except Exception:
        print("Can't remove recommendation folder... Whatever. Onward!")
    print("Recommendation Zip created and source folder cleared.")


"""
helper function to detect pre-existing files and rename them using the naming convention
"""

def get_new_filename(folder, name, extension):
    parts = name.split(' - ')
    original_second_part = parts[1]

    # If the part after " - " ends with a digit (indicating an existing counter), split it from the base name
    if original_second_part[-1].isdigit():
        base_part = original_second_part.rstrip('0123456789')  # The base name
        existing_counter = int(''.join(filter(str.isdigit, original_second_part)))  # The existing counter
    else:
        # If there's no existing counter, the base name is the entire part after " - "
        base_part = original_second_part
        existing_counter = 1

    counter = existing_counter + 1

    # Generate initial filename without appending a new counter
    name = f"{parts[0]} - {base_part}{'' if existing_counter == 1 else existing_counter} - {parts[2]}"

    # If filename with counter already exists, then increment the counter until we find a filename that doesn't exist
    while os.path.exists(os.path.join(folder, f"{name}{extension}")):
        name = f"{parts[0]} - {base_part}{counter} - {parts[2]}"
        counter += 1

    return name

"""
function move
Sorts labeled documents into their related folders
"""

def move():
    # creating each documents pattern for sorting purposes
    # catches all documents with HS, SH (user error in typing), Test in the name, ignoring capitalization.
    high_school_re = re.compile(r"\sHS\d?\s|\sSH\d?\s|\sTest\d?\s", re.I)

    # catches alld ocuments with college in the name
    college_re = re.compile(r"\sCollege\d?\s", re.I)

    # catches all documents with waiver in the name
    waiver_re = re.compile(r"\sWaiver\d?\s", re.I)

    # catches all documents with rec (recommendations) in the name
    rec_re = re.compile(r"\sRec\d?\s", re.I)

    # catches all documents with info (general information) or citizenship in the name
    info_re = re.compile(r"\sInfo\d?\s|\sCitizenship\d?\s|\sAppeal\d?\s", re.I)

    # create a dict to match folder names and regex patterns
    document_types = {"HS": high_school_re, "College": college_re, "Waivers": waiver_re, "Recommendations": rec_re, "Misc": info_re}

    # start an incremential counter for final display
    count = 0

    # set base folder and target
    base_folder = "c:\\users\\cjj714\\desktop\\transcripts\\"
    os.chdir(base_folder)

    # iterate over the keys and values in the document_types dict
    for folder, pattern in document_types.items():
        for item in os.listdir():
            if item.endswith(".pdf") and pattern.search(item):
                if not os.path.isdir(folder):
                    print(f"No {folder} folder! Creating one...")
                    os.mkdir(folder)
                    
                if os.path.isdir(folder):
                    # split the filename and extension
                    file_name, file_ext = os.path.splitext(item)
                    # get the new filename
                    new_file_name = get_new_filename(os.path.join(base_folder, folder), file_name, file_ext)
                    # generate destination path
                    dest_path = os.path.join(base_folder, folder, f"{new_file_name}{file_ext}")
                    
                    try:
                        shutil.move(item, dest_path)
                    except Exception as error:
                        print(error)

                    print("Moved: ", item)

                    count += 1

    print("Moved", str(count), "files.")
    

def move_old():
    # creating each documents pattern for sorting purposes
    # catches all documents with HS, SH (user error in typing), Test in the name, ignoring capitalization.
    high_school_re = re.compile(r"\sHS\d?\s|\sSH\d?\s|\sTest\d?\s", re.I)

    # catches alld ocuments with college in the name
    college_re = re.compile(r"\sCollege\d?\s", re.I)

    # catches all documents with waiver in the name
    waiver_re = re.compile(r"\sWaiver\d?\s", re.I)

    # catches all documents with rec (recommendations) in the name
    rec_re = re.compile(r"\sRec\d?\s", re.I)

    # catches all documents with info (general information) in the name
    info_re = re.compile(r"\sInfo\d?\s", re.I)

    # sets working directory to the transcripts folder
    os.chdir("c:\\users\\cjj714\\desktop\\transcripts\\")

    # creates a informational counter to show how many files are moved
    count = 0

    # iterates over the files in the transcripts folder
    for item in os.listdir():

        # if the document is a high school document and ends in the .pdf extension
        if item.endswith(".pdf") and high_school_re.search(item):

            # checks to see if the 'HS' folder is in the directory. If not, creates it.
            if not os.path.isdir("HS"):
                print("No HS folder! Creating one...")
                os.mkdir("HS")

            # if the 'HS' folder is in the directory, move the currently matched HS file into it.
            if os.path.isdir("HS"):
                try:
                    shutil.move(item, "c:\\users\\cjj714\\desktop\\transcripts\\HS\\")

                # in case of error, displays the error
                except Exception as error:
                    print(error)

                # displays what document was moved into the HS folder
                print("Moved: ", item)

                # incriments the file count indicator
                count += 1

        # if the document is a college document and ends in the .pdf extension
        if item.endswith(".pdf") and college_re.search(item):

            # checks to see if the 'College' folder is in the directory. If not, creates it.
            if not os.path.isdir("College"):
                print("No College folder! Creating one...")
                os.mkdir("College")

            # if the 'College' folder is in the directory, move the currently matched College file into it.
            if os.path.isdir("College"):
                try:
                    shutil.move(item, "c:\\users\\cjj714\\desktop\\transcripts\\college\\")

                # in case of error, displays the error
                except Exception as error:
                    print(error)

                # displays what document was moved into the College folder
                print("Moved: ", item)

                # incriments the file count indicator
                count += 1

        # if the document is a recommendation document and ends in the .pdf extension
        if item.endswith(".pdf") and rec_re.search(item):

            # checks to see if the 'Recommendations' folder is in the directory. If not, creates it.
            if not os.path.isdir("Recommendations"):
                print("No Recommendations folder! Creating one...")
                os.mkdir("Recommendations")

            # if the 'Recommendations' folder is in the directory, moves the currently matched Recommendation file into it.
            if os.path.isdir("Recommendations"):
                try:
                    shutil.move(item, "c:\\users\\cjj714\\desktop\\transcripts\\recommendations\\")

                # in case of error, displays the error
                except Exception as error:
                    print(error)

                # displays what document was moved into the Recommendations folder
                print("Moved: ", item)

                # incriments the file count indicator
                count += 1

        # if the document is a waiver document and ends in the .pdf extension
        if item.endswith(".pdf") and waiver_re.search(item):

            # checks to see if the 'Waivers' folder is in the directory. If not, creates it.
            if not os.path.isdir("Waivers"):
                print("No Waivers folder! Creating one...")
                os.mkdir("Waivers")

            # if the 'Waivers' directory exists, moves the currently matched Waiver file into it.
            if os.path.isdir("Waivers"):
                try:
                    shutil.move(item, "c:\\users\\cjj714\\desktop\\transcripts\\waivers\\")

                # in case of error, displays the error
                except Exception as error:
                    print(error)

                # displays what document was moved into the Info folder.
                print("Moved: ", item)

                # incriments the file count indicator
                count += 1

        # if the document is a information document and ends in the .pdf extension
        if item.endswith(".pdf") and info_re.search(item):

            # checks to see if the 'Misc' folder extists. If not, creates it.
            if not os.path.isdir("Misc"):
                print("No Misc folder! Creating one...")
                os.mkdir("Misc")

            # if the 'Misc' folder exists, move the currently matched informational file into it.
            if os.path.isdir("Misc"):
                try:
                    shutil.move(item, "c:\\users\\cjj714\\desktop\\transcripts\\Misc\\")

                # in case of error, display the error
                except Exception as error:
                    print(error)

                # displays what document was moved into the Info folder.
                print("Moved: ", item)

                # incriments the file count indicator
                count += 1

    # Displays the number of files moved to their associated folder.
    print("Moved", str(count), "files.")

    # TODO: I think I can refactor these since a lot of the code is repeated. I think I just need to pass folder name to the function.
    # example: if item.endswith (".pdf") and info_re.search(item):
    #   dest_folder("Info")
    #(SNIP)
    #   dest_folder("College)
    # etc


"""
function prompting
Creates a text based menu system for the user to makes selections to run different processes or exit the program
"""
def prompting():
    # Displays the actual menu system to the user
    print("What would you like to do?\n"
          + "1. Clean up the transcripts folders by moving NOIDS\n"
          + "2. Move HS and College to their folders\n"
          + "3. Count file sizes for load\n"
          + "4. Exit")

    # takes the users input from the menu system
    choice = input()


    # depending on the users' choice, run the specific function
    # runs the clean_transcripts function and redisplays the menu
    if choice == "1":
        clean_transcripts()
        prompting()

    ''' Defucnt
    # runs the prep_noid function and redisplays the menu
    if choice == "2":
        prep_noid()
        prompting()
    '''

    # runs the move function and redisplays the menu
    if choice == "2":
        move()
        prompting()

    # runs the counter function and redisplays the menu
    if choice == "3":
        counter()
        prompting()

    # exits the program
    if choice == "4":
        sys.exit()

    # if the users inputs anything besides 1-5, show an error message and have them try again
    else:
        end = input("I'm sorry, that is an invalid entry. Try again? (y/N)")

        # if they choose to try again by typing either Y/y/yes/Yes, redisplay the menu
        if end and end[0].lower() == 'y':
            prompting()

        # if they either purposfully or accidently end something else, terminate the program
        else:
            print("Good-bye")
            sys.exit()


"""
function counter
Because the AppXtender (indexer) for UTC has a document upload limit of 10MB (10485760), this counts
and adds files size together in order to create document batches that come in under the upload limit, coming as close
as possible to the limit for effeciency.
"""
def counter():
    # sets working directory to where the transcripts are located
    os.chdir("C:\\Users\\cjj714\\Desktop\\Transcripts")
    directory = os.getcwd()

    # creates a lists of files in the directory for iteration
    items = os.listdir(directory)

    totalsize = 0

    # iterates over the items in the list
    for item in items:

        # checks to see if the document is a pdf
        if item.endswith(".pdf"):

            # create a variable for os.path.getsize to operate on in order to get the file's size
            file_path = directory + "\\" + item
            filesize = os.path.getsize(file_path)

            # if current sum of files is greater than 10MB (changed to 9.96 MB for headroom)
            if totalsize + filesize > 9961472:

                # tell me what file will start the next indexing batch
                print("Start at " + item)

                # reset the counter for the next batch
                totalsize = 0

            # add current file size to current running total of all files counted so far
            totalsize += filesize



# runs the main program
prompting()
