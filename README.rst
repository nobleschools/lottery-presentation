Utility to generate a pptx admission lottery presentation from a csv of lottery results.

generate_presentation.py
^^^^^^^^^^^^^^^^^^^^^^^^

Generates a ppt presentation from csv dump of admission lottery
results. Call with two positional args, eg.:

.. code:: bash

  $ python generate_presentation.py <lottery results csv file> <desired output file name>

Assumes the csv is ordered by enrollment status (value of the 'lottery_number' field): accepted students first,
followed by waitlist students.  Column headers expected in the csv are:

- id
- lottery_number (value is 'Offered' if student has been offered enrollment, otherwise 'WL# XXXX' if student on waitlist)
- first_name
- last_name
- Elementary


analyze_ppt.py
^^^^^^^^^^^^^^

Taken from http://pbpython.com/creating-powerpoint.html.

Requires https://python-pptx.readthedocs.org/en/latest/index.html.  Program
takes a PowerPoint input file and generates a marked up version that shows
the various layouts and placeholders in the template.

Usage:

.. code:: bash

  $ python analyze_ppt.py <ppt file to analyze> <desired output file name>
