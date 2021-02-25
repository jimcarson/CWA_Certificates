This script will generate completion certificates for CW Academy using a template 

Python 3.7 and higher.  Libraries used:
     Python-docx - write Microsoft Word docx files
     Docx2PDF - convert a Microsoft Word docx file to a PDF file
     Pandas - read files.

`usage: certificates.py [-h] -c CLASSLIST [-l LEVEL] [-d DATE] [-D]`
                      `[-o OUTPUTDIR] [-a ADVISOR] [-ac ADVISOR_CALL]`
                      `[-aa ASSOCIATE_ADVISOR] [-aac ASSOCIATE_ADVISOR_CALL]`

optional arguments:
* -h, --help            
show this help message and exit
* -c CLASSLIST, --classlist CLASSLIST 
File containing a list of students (call sign, name)
* -l LEVEL, --level LEVEL
Level of class, Beginner, Basic, Intermediate or Advanced
* -d DATE, --date DATE  
Use a specific date for the report (Default: Sep-Oct 2020)
* -D, --debug           
Toggle debugging mode
* -o OUTPUTDIR, --outputdir OUTPUTDIR
Output directory (Default: output)
* -a ADVISOR, --advisor ADVISOR
Advisor (Default: Jim Carson)
* -ac ADVISOR_CALL, --advisor_call ADVISOR_CALL
Advisor Call Sign (Default; WT8P)
* -aa ASSOCIATE_ADVISOR, --associate_advisor ASSOCIATE_ADVISOR
Associate Advisor (Default: None)
* -aac ASSOCIATE_ADVISOR_CALL, --associate_advisor_call ASSOCIATE_ADVISOR_CALL
Associate Advisor Call Sign (Default: None)
