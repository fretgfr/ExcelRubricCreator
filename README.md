# About

This project started when I needed to create a grading rubric for a class that didn't update grades regularly but did provide me with the grades I had for completed assignments. 

### Obtain/Build

Either download one of the packaged builds that I've provided or download the script and run it. Please note that it requires the `openpyxl` Python library that can be installed using pip.

You can use `pip install -r requirements.txt` to get all dependencies.

## Using

Run the file, either using one of the provided binaries or launching the script yourself.

The script will guide you through creating an Excel file to calculate your grades.

It will walk you through the following sections:

### Points vs Percentages

The script will start by asking you whether you want to generate a file that calculates using points or percentages. Points should be used if you get points on assignments and your total is calculated using `Points Earned / Total Points`. 


### Enter Possible Grades

The script will then ask youu to enter the possible grades you can earn for the class. A typical example would be
```
A
A-
B+
B
B-
C+
C
C-
D
F
```

When you're done entering the different values, submit a blank one to move to the next section, which will ask you what the minimum grade to obtain that grade value would be. A typical example would be 
```
A       ->      96
A-      ->      90
B+      ->      85
B       ->      80
B-      ->      77
C+      ->      74
C       ->      70
C-      ->      65
D       ->      60
F       ->      0
```

### Assignments

The script will then ask you about the assignments that you are assigned. It will ask you for the name and the number of points the assignment is worth or the weight it has on your grade as a percentage.

#### Points

Enter the assignment name and it will ask you how many points that assignment is worth. Enter these values until you have no more assignments, where you'll leave the assignment name blank and hit enter.

#### Percentage

Enter the assignment name and it will ask you what the weight of that assignment is. This should be expressed as a decimal percentage (i.e. 10% -> 0.10). Enter these values until you have no more assignments, where you'll leave the assignment name blank and hit enter.

### Name the Output File

Enter a name for the output file and a file with that name will be created in the same folder you ran the script in.

### Open the File

Open the file that was created in whatever spreadsheet app you prefer. I have tested the following applications and confirmed that the output works in them:

- Excel
- Apple Numbers
- LibeOffice Calc
- Gnumeric
- Google Sheets

Input your earned grades into the appropriate cells and the sheet will output your final letter grade as well as your percentage in the class. You can also utilize Excel's `Goal Seek` function to determine what you need on an assignement to get a grade you're looking for (e.x. determine what you need to get on a final to get an A in a class).

# Contributing

Feel free to submit a pull request and make any improvements you see fit.