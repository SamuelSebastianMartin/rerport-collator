Report Collator
===============

The file report_merge.py in this folder is for preparing the report data
before the mail-merge. The output is an excel spreadsheet with all the
necessary data in it.

Use
---

  1. Ensure data files are in the directory:
        acengmarks.csv, umwmarks.csv (qmmarks.csv if needed)
        plus the names-emails csv

  2. `python3 report_merge.py`

Data files
----------

  * the `.csv` files downloaded from Moodle. These provide the grades,
    and should be unchanged from their downloaded format.
    They must be named:
        ```
        acengmarks.csv, umwmarks.csv (qmmarks.csv if needed)
        ```

  * the `csv` files containing the names of students. These should have
    3 columns, named Email address,Surname,Forenames.
    Note that "Email address" must be written exactly (Capital E,
    no white space). The first row is the column headings, the
    remaining rows are students. For example:
        ```
        Email address,forename,surname
        654732@soas.ac.uk,John,Smith
        etc.
        ```

Note to self: Please leave a copy of last terms .csv files in the folder as an aide mémoire for the correct filenames - data files must be named correctly.


STILL TO DO
-----------

  * allow option of what dbs to combine - umw, eng (+qm - optional)

  * capitalise surname, and Camel  first-name

  * change umw, eng, qm filenames to moodleqm.csv, moodleumw.csv etc.

  * visual file-picker for opening

  * more forgiving and fuzzy recognition of email column.

  * display list of columns - click to select/reject (or some way to
    deal with the long string-literals of column values.

