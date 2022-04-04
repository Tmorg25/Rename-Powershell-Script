# rename powershell script
Dear Mason, I have no fucking clue if this is going to work. It did only take a few hours so if it doesn't i really don't give a fuck and I learned how to write some basic powershell code so that was kinda fun. I wrote this fancy lil markdown readme so you could see what happens and ideally make this work for your use case. Obviously ask if you have any questions.

# actually using the script
run this as a powershell script. My laptop didn't like me running powershell scripts, so if your computer also doesn't then you may have to launch powershell as an administrator or change execution rules -- if this happens lmk and heres a semi useful stack exchange link that could help fix it https://superuser.com/questions/106360/how-to-enable-execution-of-powershell-scripts
## expects
so to use this you need a file directory that kinda looks like what is below

```
grandparent
├───p1
│   └───bid
│           excel_sheet.xlsx
│           word_doc.docx
│
├───p2
│   └───bid
│           excel_sheet.xlsx
│           word_doc.docx
│
├───p3
│   └───bid
│           excel_sheet.xlsx
│           word_doc.docx
│
├───p4
│   └───bid
│           excel_sheet.xlsx
│           word_doc.docx
│
└───p5
    └───bid
            excel_sheet.xlsx
            word_doc.docx
```
## what happens
this script is going to rename every file named ``excel_sheet.xlsx`` and ``word_doc.docx`` with ``excel_sheetpX.xlsx`` and ``word_docpX.docx`` where the pX in each is the folders ``p1 ... p5`` as shown below

```
grandparent
├───p1
│   └───bid
│           excel_sheetp1.xlsx
│           word_docp1.docx
│
├───p2
│   └───bid
│           excel_sheetp2.xlsx
│           word_docp2.docx
│
├───p3
│   └───bid
│           excel_sheetp3.xlsx
│           word_docp3.docx
│
├───p4
│   └───bid
│           excel_sheetp4.xlsx
│           word_docp4.docx
│
└───p5
    └───bid
            excel_sheetp5.xlsx
            word_docp5.docx
```
it only searches for files named ``excel_sheet.xlsx`` and ``word_doc.docx`` that are one level down from the folder called ``bid`` so any other files shouldn't matter. If the file names are not exactly ``excel_sheet.xlsx`` and ``word_doc.docx`` and they are not exactly one level down from the folder called ``bid`` it shouldn't touch the files, but I didn't test this too much so could have easily missed something.

## customizing for your case
i allowed you to input the excel sheet name, the wrod doc name and the grandparent directory name as whatever you want as command line parameters. the command for including this is shown below
```
.\rename.ps1 -base_dir GRANDPARENT_PATH -sheet_name SHEET_NAME -doc_name DOC_NAME
```
where `GRANDPARENT_PATH` is the full path (starting at the drive - ex. `D:\workspace\grandparent`) to the grandparent folder, `SHEET_NAME` is the exact name of the excel sheet and `DOC_NAME` is the exact name of the word doc before the names are changed to include their parent folders

You have to run the file from whichever directory it is located in, but this lets run modify different directory trees.

It defaults to use whatever directory you run it from as the `GRANDPARENT_PATH`, ``excel_sheet.xlsx`` as `SHEET_NAME` and ``word_doc.docx`` as `DOC_NAME`

please do whatever you want with this and mess with / modify the code if you actually do end up wanting to use it. if not and you never look at this again that's cool too, it gave me something to do while bored on a sunday evening.
