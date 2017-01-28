.. contents:: **Table of Contents**

************************
generate_presentation.py
************************

Generates a ppt presentation from csv dump of admission lottery
results. Call with two positional args, eg.:

.. code:: bash

  $ python generate_presentation.py <lottery results csv file> <desired output file name>

Assumes the csv is ordered by status: accepted students first,
followed by waitlist students.  Column headers expected in the csv:

- id
- lottery_number (value will 'Offered' if student has been offered enrollment, otherwise 'WL# XXXX' if student on waitlist)
- first_name
- last_name
- Elementary


**************
analyze_ppt.py
**************

Taken from http://pbpython.com/creating-powerpoint.html.

Requires https://python-pptx.readthedocs.org/en/latest/index.html.  Program
takes a PowerPoint input file and generates a marked up version that shows
the various layouts and placeholders in the template.

Parameters:
  * *infile_name* -- ppt file to analyze.
  * *outfile_name* -- filename to save output ppt as.
