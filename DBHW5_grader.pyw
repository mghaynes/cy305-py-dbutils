import os, tkinter, pypyodbc, tkinter.messagebox
import sys, csv, datetime

# sys.path.append(r"\\usmasvddeecs\eecs\S&F\Courses\IT305\libraries")
import dbUtils as db

pypyodbc.lowercase = False
tk = tkinter.Tk()
displayFrame = tkinter.Frame(tk)
displayFrame.pack()

cdtDict = {}
sections = []
sec = ""
debug = 1  # Set from 0 or 2 to get varying levels of output; 0=no output, 2=very verbose

'''Some global variables we should change from HW to HW'''
xNumFile = r"\\usmasvddeecs\eecs\S&F\Courses\IT305\ay181\admin\sectioning\xNums181.csv"
db_file_name = "program_tracker_hw5.accdb"  # student db file name
dbPath = r"\\usmasvddeecs\eecs\S&F\Courses\IT305\ay181\lessons\Lsn18-Databases5\HW\solution\program_tracker_hw5_soln.accdb"  # solution file
solnTableNames = []  # Queries we are interested in checking
solnQueryNames = ['EmployeeBioBrief', 'TopSalesFigures']  # Tables we are interested in checking

helpText = '''This program checks CY305 AY18-1 DB HW 5. 

File (0 pts) - Successfully open "program_tracker_hw5.accdb
 in the Cadet's proper path
===============================================

QUERIES - 100% (3 points total)
-----------------------------------------------
H, N, Z Tbl - Checks for tables named "HighVolumeSales" 
           and "NonSupportSales" and "ZeroSaleEmployees" 
           (.5 pt each)
         -NOTE: Will give 0 for entire query if not named correctly.
-----------------------------------------------
H, N, Z R/C - Checks HighVolumeSales has 4 rows and 
           3 columns. (.25 pt each)
         - Checks NonSupportSales has 4 rows and
           2 columns. (.25 pt each)
         - Checks ZeroSaleEmployees has 4 rows 
           and 2 columns. (.25 pt each)
-----------------------------------------------
H, N, Z Out - Checks row by row output of trackers and makes
           sure same as correct output.
         - Row ordering matters, column ordering does not 
           matter. Extra columns do not matter. (.5 pt each)

===============================================
Total - Shows percent received.
Points - Converts to suggested HW points (out of 3 possible)'''


def helpDiag():
    tkinter.messagebox.showinfo("Grading Info", helpText)


if os.path.exists(xNumFile):
    fExists = True
    csvFObj = open(xNumFile)
    csvContents = csvFObj.readlines()
    csvFObj.close()
    for line in csvContents:
        line = line.strip().split(',')
        if len(line[3].strip()) != 1:
            line[2] += "(" + line[5].strip()[0:3] + ")"
        cdtDict[line[0]] = line[1:]
        if line[3].strip() + line[4].strip() not in sections:
            sections.append(line[3].strip() + line[4].strip())
else:
    print("XNums.csv not found")


def DisplayTableScore(scoreVector, cdtRes):
    for cnt, score in enumerate(scoreVector):
        if cnt != 1:
            if score == 0:
                cdtRes.append('--'.center(6))
            else:
                cdtRes.append('Good'.center(6))
        else:
            if score == 0:
                cdtRes.append('--'.center(6))
            elif score == 1:
                cdtRes.append('Ok'.center(6))
            elif score == 2:
                cdtRes.append('Good'.center(6))
            else:
                print('Unexpected score:', score, 'position:', cnt)


def ScoreToGrade(scoreVector, rubric):
    grade = 0.0
    for cnt, score in enumerate(scoreVector):
        if cnt != 1:
            grade += score * rubric[cnt]
        else:
            grade += score * rubric[cnt] / 2
    return grade


def setSection(section, tk):
    global displayFrame
    global sec
    tk.title("Retrieving " + section + ": HW 5")
    tk.update()
    sec = section
    newFrame = tkinter.Frame(tk)
    # display column names in Tkinter
    labels = ['Name', 'File', 'EB Tbl', 'EB R/C', 'EB Out', 'TS Tbl', 'TS R/C', 'TS Out', 'Total', 'Points']
    for cnt, label in enumerate(labels):
        label = tkinter.Label(newFrame, text=label, font=("Courier New", 14, "bold"))
        label.grid(row=0, column=cnt)
    cdtCount = 0
    cdtOutput = []

    for cdt in cdtDict:
        if cdtDict[cdt][2].strip() + cdtDict[cdt][3].strip() == section.strip():
            grade = 0
            total = 3
            cdtRes = []
            cdtCount += 1
            cadet_name = cdtDict[cdt][0] + ',' + cdtDict[cdt][1]
            print(cadet_name)
            cdtRes.append((cdtDict[cdt][0].strip() + ", " + cdtDict[cdt][1].strip())[:20].ljust(22))
            workPath = r"\\usmasvddeecs\eecs\Cadet\Courses\CY305"
            workPath = os.path.join(workPath,
                                    cdtDict[cdt][4].strip(),
                                    section,
                                    cdtDict[cdt][0] + "." + cdtDict[cdt][1].strip(),
                                    "database",
                                    "hw5",
                                    db_file_name)
            if "(" in workPath:  # strip out instructor names from test locations
                begin = workPath.find("(")
                end = workPath.find(")")
                workPath = workPath[:begin] + workPath[end + 1:]
            # print(workPath)

            wpg = True
            try:  # Try to connect to the database
                conn = pypyodbc.connect(
                    r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};" + "Dbq={0};".format(workPath))
            except:
                print('Cannot open DB')
                wpg = False

            if wpg:  # if we have a good connection to the database
                cur = conn.cursor()
                cdtRes.append("Good".center(6))
                # grade += 1 # no points for file found

                try:
                    studentTables = db.GetTableNames(cur)
                    studentQueries = db.GetQueryNames(cur)
                    if debug:
                        print('TABLES:', studentTables, '\nQUERIES:', studentQueries)
                    '''--------------------------------------------------------------------------------------
                        Check query tables
                    --------------------------------------------------------------------------------------'''
                    total = 3
                    rubric = [.5, .5, .5]
                    goodStudentNames = set(studentQueries).intersection(set(solnQueryNames))
                    badStudentNames = set(studentQueries).difference(set(goodStudentNames))
                    errorTables = []
                    print('Good Names:', goodStudentNames)
                    print('Bad Names:', badStudentNames)
                    # Loop through the solution tables/queries
                    for tableName in solnQueryNames:
                        print('ANALYZING:', tableName)
                        scoreVector = [0, 0, 0]
                        solnTable = db.Table(dbPath, tableName, type='QUERY')
                        if tableName in goodStudentNames:
                            bestBadTableName = tableName
                            try:
                                studentTable = db.Table(workPath, tableName, type='QUERY')
                                scoreVector = db.GradeTables(solnTable, studentTable)
                            except Exception as e:
                                print('TABLE ERROR:', e)
                        else:
                            maxScore = 0
                            bestBadTableName = ''
                            for badTableName in badStudentNames:
                                if badTableName not in errorTables:
                                    try:
                                        studentTable = db.Table(workPath, badTableName, type='QUERY')
                                        scoreVector = db.GradeTables(solnTable, studentTable)
                                    except Exception as e:
                                        print('TABLE ERROR:', e)
                                        errorTables.append(badTableName)
                                if sum(scoreVector) > maxScore:
                                    bestBadTableName = badTableName
                        # if sum(scoreVector) > 0:
                        #     studentTablesRemaining.remove(bestTableName)
                        # for badTableName in badTableNames:
                        #     studentTablesRemaining.remove(badTableName)
                        print('SOLUTION TABLE:',tableName,'\tBEST MATCH:',bestBadTableName,'\tSCORE:',scoreVector)
                        # Workaround for this HW due to discrepancy between solution query name
                        # and name used in written HW document
                        if bestBadTableName == "TopSaleFigures" and tableName == "TopSalesFigures":
                            scoreVector[0] = 1
                        grade += ScoreToGrade(scoreVector, rubric)
                        DisplayTableScore(scoreVector, cdtRes)
                except Exception as e:
                    print('Problem', e)

            # perc = str(round(grade / total * 3, 2)).rjust(5)
            score = int(round(grade * 100 / total, 0))
            perc = ''.join([str(score), ('%')]).rjust(5)
            cdtRes.append(perc.center(6))
            # create the points column
            if (score >= 70):
                cdtRes.append("3".center(6))
            elif (score < 70) and (score >= 50):
                cdtRes.append("2".center(6))
            elif (score < 50) and (score > 0):
                cdtRes.append("1".center(6))
            elif (score <= 0):
                cdtRes.append("0".center(6))

            cdtOutput.append(cdtRes)

    rowNum = 1
    cdtOutput.sort()
    for cadet in cdtOutput:
        colNum = 0
        for item in cadet:
            color = 'black'
            if len(item.strip()) <= 6:
                # if item.strip() == 'Good' or item.strip() in ['5/5','4/5'] or score>=0.7:
                if (item.strip() == 'Good') or (item.strip() in ["5/5", "4/5"]):
                    color = 'dark green'
                # elif item.strip() == 'Ok' or item.strip() in ['1/5','2/5','3/5'] or (score <0.7 and score>=0.5):
                elif (item.strip() == 'Ok') or (item.strip() in ["1/5", "2/5", "3/5"]):
                    color = '#F39C12'  # a dark orange color
                # elif item.strip() == '--' or item.strip() == '0/5' or score <= 0.5:
                elif (item.strip() == '--') or (item.strip() in ["0/5"]):
                    color = 'red'
                # Have to do this series of elif's since item.strip() is not always a number
                elif ('%' not in item):
                    color = 'black'
                elif (int(item.strip().strip('%')) >= 70):
                    color = 'dark green'
                elif (int(item.strip().strip('%')) < 70) and (int(item.strip().strip('%')) >= 50):
                    color = '#3498DB'  # a light blue color
                elif (int(item.strip().strip('%')) < 50) and (int(item.strip().strip('%')) > 0):
                    color = '#F39C12'  # a dark orange color
                elif (int(item.strip().strip('%')) <= 0):
                    color = 'red'
                else:
                    color = 'black'
            label = tkinter.Label(newFrame,
                                  text=item,
                                  font=("Courier New", 14, "bold"),
                                  fg=color,
                                  relief="ridge")
            label.grid(row=rowNum, column=colNum)
            colNum += 1
        rowNum += 1

    label = tkinter.Label(newFrame,
                          text=str(cdtCount) + " Cadets",
                          font=("Courier New", 14, "bold"),
                          relief="ridge")
    label.grid(row=rowNum + 1, column=0)

    displayFrame.forget()
    newFrame.pack()
    for slave in displayFrame.grid_slaves():
        slave.grid_remove()
        slave.destroy()
    displayFrame = newFrame
    if section == "":
        section = "Instructors"
    tk.title("Section " + section + ": DB HW 5")
    tk.update()


def makeSectionButton(section, frame, tk):
    labelsection = section
    if labelsection == "":
        labelsection = "Inst"
    sectionMenu.add_radiobutton(label=labelsection,
                                command=(lambda: setSection(section, tk)))


menuBar = tkinter.Menu(tk)
sectionMenu = tkinter.Menu(tk)
menuBar.add_cascade(label="Section", menu=sectionMenu)

refreshMenu = tkinter.Menu(tk)
menuBar.add_cascade(label="Refresh", menu=refreshMenu)

menuBar.add_command(label="Help", command=helpDiag)

refreshRate = 60000


def setRefresh(refresh):
    global refreshRate
    refreshRate = refresh


def makeRefreshButton(refresh, selected=False):
    refreshMenu.add_radiobutton(label=str(refresh / 1000) + " sec",
                                command=(lambda: setRefresh(refresh)))


makeRefreshButton(5000)
makeRefreshButton(10000, True)
makeRefreshButton(20000)
makeRefreshButton(30000)
makeRefreshButton(60000)

for sect in sections:
    makeSectionButton(sect, displayFrame, tk)


def refresh():
    setSection(sec, tk)
    tk.after(refreshRate, refresh)


tk.after(refreshRate, refresh)

tk.config(menu=menuBar)

tk.mainloop()
