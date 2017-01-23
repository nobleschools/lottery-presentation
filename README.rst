.. contents:: Table of Contents

************************
generate_presentation.py
************************

Generates a ppt presentation from csv dump of admission lottery
results. Call with two positional args:

.. code:: bash

  generate_presentation.py <lottery results csv file> <desired output file name>

Assumes the csv is ordered by status: accepted students first,
followed by waitlist students.  Column headers expected in the csv:

- id 
- lottery_number
- first_name
- last_name
- Elementary

class generate_presentation.PresentationMaker(*infile_name*, *outfile_name*):
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

A PresentationMaker instance is a manager object used to read data from a
csv of lottery results and add them to a pptx.Presentation
object.  The necessary steps are wrapped in its make_presentation method.

Parameters:
  * *infile_name* -- name of csv lottery results file.
  * *outfile_name* -- filename to save ppt presentation as.

**make_presentation()**:

  Read in data from the infile csv, and save a pptx under the outfile name.

  Creates two sections (a title and set of following body slides)
  -- one for admitted students, and one for waitlist students.


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
