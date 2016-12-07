#!/usr/bin/env python3

# Author: Huon Imberger
# Date: 2016-12-08
# Description:
#     This script explodes an exported PDF (BulkPrint) from Springboard (http://www.springboard.com.au/)
#     into a CSV listing candidate details and individual resume files.
#
#     It creates a CSV file containing candidate's title, first name, last name, and email address.
#     It also has an ID field, which associates each candidate with their resume. Note this is not the same as
#     Springboard's internal candidate ID, which isn't included in the exported data.
#
#     Resumes are named [candidate_id].pdf where [candidate_id] corresponds to the 'id' field in the CSV.
#
# Notes:
#     - Candidates that are missing their resume (don't have one, or there was an error in the BulkPrint) will
#       simply not have an associated resume in OUTPUT_PDF_DIR
#     - There seems to be a bug in the BulkPrint causing incorrect ordering of pages when there are two
#       candidates with the same name next to each other. The script should handle this fine, since the resume
#       always comes after the "Candidate Details" page.
#
# PyPDF2: See here for licencing: https://github.com/mstamy2/PyPDF2
#


from PyPDF2 import PdfFileWriter, PdfFileReader
import re, csv, os, sys

INPUT_PDFS = ['bulkprint_huge.pdf',]
OUTPUT_DIR = 'output'

# Any candidates with a title not in this list will be incorrectly parsed
# and end up with a title as their first name
TITLES = ['Mr', 'Mrs', 'Ms', 'Miss', 'Dr', 'Professor']


# Check if all input pdfs exist
for input_pdf in INPUT_PDFS:
    if not os.path.isfile(input_pdf):
        print("At least one input file doesn't exist: {0}".format(input_pdf))
        sys.exit()

# Create output folder structure
try:
    resume_output_dir = "{0}/resumes".format(OUTPUT_DIR)
    os.makedirs(resume_output_dir, exist_ok=True)
except OSError as e:
    print("Could not create output directories: {0}".format(e))
    sys.exit()

# Prepare CSV writer
csvfile = open("{0}/candidates.csv".format(OUTPUT_DIR), 'w', newline='')
fieldnames = ['id', 'title', 'first_name', 'last_name', 'email']
csvwriter = csv.DictWriter(csvfile, fieldnames=fieldnames)
csvwriter.writeheader()

# Stats
resume_count = 0
resume_errors = 0

# Start at 0 so first candidate is 1
candidate_id = 0

# Process each input PDF
for input_pdf in INPUT_PDFS:
    print("Processing file: {0}...".format(input_pdf))

    # Open input PDF as binary
    doc = PdfFileReader(open(input_pdf, "rb"))

    # Loop through outline (list of Destination objects, aka bookmarks)
    for i, dest in enumerate(doc.outlines):
        # Ignore sub-destinations, which are always lists (some resumes will have their own bookmarks)
        if isinstance(dest, list):
            continue

        # Get page details associated with this outline (Destination object)
        page_num = doc.getDestinationPageNumber(dest)
        page = doc.getPage(page_num)

        # Is this page the start of a new candidate?
        if re.match(r'.* Candidate Details', dest.title):
            candidate_id += 1

            # Extract candidate details
            details = page.extractText().split('\n')
            c_fullname = details[0].strip().split()
            print(" ".join(c_fullname))
            if c_fullname[0] in TITLES:
                # Title must exist
                c_title = c_fullname[0]
                c_firstname = c_fullname[1]
            else:
                # No title
                c_title = ""
                c_firstname = c_fullname[0]

            c_lastname = c_fullname[-1]  # Last element, to ignore first/middle names       
            c_email = details[details.index('Email Address:')+1].strip()

            # Write details to CSV
            csvwriter.writerow({'id': candidate_id,
                                'title': c_title,
                                'first_name': c_firstname,
                                'last_name': c_lastname,
                                'email': c_email})

        # Is the page the start of a resume?
        # Note: matches end of string, since resumes that failed for whatever reason
        #     will have (Error) or (Pending) or (Embedded) at the end. We want to ignore these pages.
        if re.match(r'.*\.pdf$', dest.title):
            resume_count += 1
            print("  Found resume")

            # Find the next destination that isn't a sub-destination (which are always lists)
            k = i
            while True:
                k += 1
                next_dest = doc.outlines[k]
                if not isinstance(next_dest, list):
                    break

            # Extract pages between current and next destination, which should be the whole resume
            next_dest_pagenum = doc.getDestinationPageNumber(next_dest)

            resume = PdfFileWriter()
            for j in range(page_num, next_dest_pagenum):
                resume.addPage(doc.getPage(j))
                
            # Create the file
            with open("{0}/resumes/{1}.pdf".format(OUTPUT_DIR, candidate_id), "wb") as file:
                resume.write(file)

        elif re.match(r'.*\.pdf \(.*\)$', dest.title):
            # If this matches, it's a resume with an error
            resume_errors += 1

# Close CSV
csvfile.close()
print("\n---Done!---")
print("Candidates: {0}".format(candidate_id))
print("Resumes: {0}".format(resume_count))
print("Resume errors: {0}".format(resume_errors))