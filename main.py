__author__ = '@thebongy (Rishit Bansal) <rishit.bansal0@gmail.com>'
__license__ = 'MIT'
__status__ = 'Development'

import subprocess,os,sys
from docx import Document
from docx.shared import Pt,Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

FONT = 'Courier New'
SIZE = Pt(11)
TOP = 0.3
BOTTOM = 0.3
LEFT = 0.7
RIGHT = 0.7

if os.name == 'nt':
    PYTHON='C:\Python27\python.exe'
else:
    PYTHON='/usr/bin/python'
print 'Is your python executable located at',PYTHON,'?'

PYTHON = raw_input('If yes, just press enter, otherwise enter the exectuable path').strip() or PYTHON 
DIR = raw_input('Enter the directory with your assignment programs:\n')

FILES = raw_input('Enter filenames (without .py) in order, seperated by commas').split(',')

heading = raw_input('Enter the heading for the Outputs Word File:\n')

document = Document()
head_p = document.add_paragraph()
head_p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
head_r = head_p.add_run(heading)

head_font = head_r.font
head_font.name = FONT
head_font.size = SIZE

print '_______________________'
print '''Now the program will attempt to execute your files one by one. Note that
the output shown will be written to the word file simultaneously. If you make a mistake
while running a file, don't worry, you will be asked for the option of re-running the 
file in the end, and erasing your old output.

Note: Ensure that all your raw_input() prompt messages end with a '\\n' or output may not 
format correctly near raw_input() statements!!!!
'''

for f in FILES:
    print '____________________ Running ' + f + '.py'
    f_name = PYTHON + ' -u ' + os.path.join(DIR,f+'.py')
    done = False
    while not done:
        proc = subprocess.Popen(f_name,
        shell=True,
        stdin = sys.stdin,
        stdout = subprocess.PIPE)
        output = f+')\n'

        while proc.poll() == None:
            data = proc.stdout.readline()
            print data.rstrip()
            output += data
        
        output+=proc.communicate()[0]
        print '____________________ File Execution Complete
        done = raw_input('Would you like to rerun the file? (y for yes, n for no)')

    prog_p = document.add_paragraph()
    prog_r = prog_p.add_run(output)
    prog_font = prog_r.font
    prog_font.name = FONT
    prog_font.size = SIZE

sections = document.sections
for section in sections:
    section.top_margin = TOP
    section.bottom_margin = BOTTOM
    section.left_margin = LEFT
    section.right_margin = RIGHT
print 'ALL FILES HAVE BEEN RUN SUCCESFULLY!!'
x = raw_input('Enter the name to save the outputs word file as')
document.save(x+'.docx')
