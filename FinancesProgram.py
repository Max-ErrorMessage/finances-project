import pyodbc  # Module for interaction with the database
import datetime  # Module for up-to-date information
from datetime import datetime  # Apparently necessary
import matplotlib  # module for creation of charts
import matplotlib.pyplot as plt  # module for creation of charts
import numpy as np  # creation of windows I think
from tkinter import *  # module for creation of GUI


# outputs a list containing the current week and month of the year
def currentInfo():
    currentdate = datetime.today()  # current info in date Type
    year, week_num, day_of_week = currentdate.isocalendar()  # only way I could find online to get the week number
    if day_of_week > 6:  # datetime module starts the week on Monday whereas Access starts the week on Sunday so a
        # correction is required
        week_num = week_num + 1
    return [week_num + 1, currentdate]  # format of currentdate is as 2022-09-24 12:22:21.578622


# runs the sqlcommand in Spending.accdb and returns the output generated
def selectFromFile(sqlcommand):  # runs the SQL command in Spending.accdb and returns the result
    conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                r'DBQ=' + filelocation + '\Finances Project\Spending.accdb;')  # fetches data from the Spending
    # database in the folder
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    # print(sqlcommand)  # useful debugging tool
    cursor.execute(sqlcommand)  # executes any command but only should be used for SELECT commands
    return cursor.fetchall()  # returns the value of the select command


# collects information about one field from the database
def getInfo(targetinformation):  # the parameter targetinformation is only inputted as either 'Purchaser' or
    # 'PurchaseType' in this program, so I have not tested if other parameters work
    allinfo = selectFromFile(f"SELECT {targetinformation} FROM qryPurchases ORDER BY {targetinformation};")  # orders
    # the data alphabetically to ensure consistency otherwise the file is read from the bottom up, so the orientation
    # of data slices around the pie charts will change
    info = []
    for i in range(len(allinfo)):
        if allinfo[i][0] not in info:  # checks for the data type in the list already to avoid duplicates
            # does not use .lower() to allow for more detail in naming conventions ('F' different to 'f' perhaps)
            info.append(allinfo[i][0])
    return info


# returns a float value of the sum of all costs within the data specified by the parameters
def sumPrice(timevalue, purchasetypevalue, purchaservalue, communalvalue):
    if timevalue.lower() == "week" or timevalue.lower() == "w":  # many options are given for the input as this input
        # originates from the user
        timesql = f'((qryPurchases.WeekOfPurchase) = {str(currentInfo()[0])})'  # collects data from the current week

    elif timevalue.lower() == "last week" or timevalue.lower() == "lastweek" or timevalue.lower() == "lw":  #
        # collects data from last week
        timesql = f'((qryPurchases.WeekOfPurchase) = {str(currentInfo()[0] - 1)})'

    elif timevalue.lower() == "month" or timevalue.lower() == "m":
        timesql = f'(qryPurchases.MonthOfPurchase) = {str(currentInfo()[1][5:7])})'  # the value outputted by
        # currentInfo()[1] is always in the same format and the month is always in positions 5 - 7

    elif timevalue.lower() == "all" or timevalue.lower() == "a":
        timesql = f'(qryPurchases.PurchaseID > 0)'  # some sql statement is required as AND statements are later
        # used. IDs are automatically given to new values starting at 1, so this requirement is guaranteed to be True

    elif timevalue[0:2] == "wk":
        targetweek = int(timevalue[2:len(timevalue)])
        timesql = f'((qryPurchases.WeekOfPurchase) = {str(targetweek)})'  # collects data only from the specified week

    elif timevalue[0:2] == "dy":
        date = timevalue[2:len(timevalue)]
        timesql = f'((qryPurchases.DayOfPurchase) = {str(date)})'  # collects data only from the specified date

    else:
        try:
            timesql = f"((qryPurchases.DateOfPurchase) Between Date() And Date() - {str(int(timevalue))})"  # if no
            # keywords are used to specify the timevalue then the program checks if an integer is applicable (to
            # view the last x days where x is the given integer)
        except TypeError:  # if the inputted timeframe cannot be represented as an integer an error is necessary
            error('Please use "week", "month", or input the amount of days you want to look at.', 40)
            timesql = '(WeekOfPurchase = -1)'  # no record will have a WeekOfPurchase of -1 so this returns no values

    if purchasetypevalue != "":  # an empty purchasetypevalue signifies that the return value should add up all
        # records specified by the other parameters regardless of the purchaseType
        purchasetypesql = f"And ((qryPurchases.PurchaseType) LIKE '{purchasetypevalue}')"  # even if an incorrect
        # value is inputted as purhasetypevalue this will not create an error - just return an empty value
    else:
        purchasetypesql = ""  # if no specification is given then the SQL command requires no criteria for the
        # PurchaseType field

    if purchaservalue != "":  # similar concept to PurchaseType above
        purchasersql = f" And((qryPurchases.Purchaser) = '{purchaservalue}')"
    else:
        purchasersql = ""

    if str(communalvalue) != "True" and str(communalvalue) != "False":  # str() value of communalvalue is used to
        # ensure that it doesn't matter whether the parameter is inputted as Boolean or str, despite the program only
        # inputting as Boolean
        communalsql = ""  # if the parameter is neither True nor False then the program considers it an indication
        # that eny value is acceptable
    else:
        communalsql = f" And((qryPurchases.Communal) = {communalvalue})"

    costovertime = selectFromFile(f'SELECT qryPurchases.Cost '
                                  f'FROM qryPurchases '
                                  f'WHERE ({str(timesql)} '
                                  f'{str(purchasetypesql)}'
                                  f'{purchasersql}'
                                  f'{communalsql}'
                                  f') '
                                  f'ORDER BY qryPurchases.Cost DESC;')  # only one selectFromFile command is
    # necessary as the components of the command are deiced earlier by a series of if statement that interrogate the
    # parameters
    totalcost = 0
    for i in range(len(costovertime)):  # adds up all the values to return a sum (as specified in the name of the
        # function) and not a list of values
        totalcost = totalcost + +costovertime[i][0]
    return totalcost  # returns the sum of the costs of all records within the criteria specified by the parameters


# takes in an array of decimal numbers and returns an array of the same numbers formatted as pounds and pence
def niceValues(values):
    for i in range(len(values)):  # the input parameter is a list, so it must be searched one by one
        if values[i] != 0:  # the function is used for making the output of decimal values appear nicely on pie
            # charts, so a 0 or single £ sign is often not the desired result when displaying this data, so values of
            # 0 are replaced with empty string
            values[i] = '{:.2f}'.format(values[i])
            values[i] = "£" + values[i]
        else:
            values[i] = ""
    return values  # returns the list still in list form


# compares personal and communal spending for different purchasers
def compareSpending1(timevalue):
    try:
        matplotlib.pyplot.close()  # if the previous plot is not closed, matplotlib draws over the plot and creates
        # very confusing images
    except NameError:  # if this is the first instance of a pie chart in the running of the program, pyplot will not
        # yet have been defined so an error will occur if the closing is not part of a try and except with NameError
        pass  # nothing actually needs to be done if there is no existing pyplot

    purchasers = getInfo("Purchaser")  # creates a list of every Purchaser with no duplicates

    values = []  # will be used for representing slices in the pie chart
    labelsarray = []  # will be used to labelling slices in the pie chart

    for i in range(len(purchasers)):  # adds communal and non-communal spending for every purchaser
        values.append(sumPrice(timevalue, "", purchasers[i], True))
        labelsarray.append(purchasers[i] + "'s Communal Spending")  # labels are added at the same time as values for
        # simplicity
        values.append(sumPrice(timevalue, "", purchasers[i], False))
        labelsarray.append(purchasers[i] + "'s Personal Spending")

    total = niceValues([sum(values)])  # the total variable is for display purposes only so niceValues() is applied

    spendingchart = np.array(values)  # creates the pie chart

    plt.pie(spendingchart, labels=niceValues(values))  # plots the pie chart

    plt.legend(labels=labelsarray, loc=1)  # displays the key in the top right
    plt.title("Total : " + str(total[0]))  # total is saved as a list to input it as a parameter into niceValues()
    # but only the first and only value in the list needs to be outputted, without the [] around the string
    plt.show()  # shows the pie chart


# compares purchase types
def compareSpending2(timevalue):
    try:
        matplotlib.pyplot.close()  # if the previous plot is not closed, matplotlib draws over the plot and creates
        # very confusing images
    except NameError:  # if this is the first instance of a pie chart in the running of the program, pyplot will not
        # yet have been defined so an error will occur if the closing is not part of a try and except with NameError
        pass  # nothing actually needs to be done if there is no existing pyplot

    purchasers = getInfo("Purchaser")  # creates a list of every Purchaser with no duplicates
    purchasetypes = getInfo("PurchaseType")  # creates a list of every PurchaseType with no duplicates

    values = []  # will be used for representing slices in the pie chart
    labelsarray = []  # will be used to labelling slices in the pie chart

    for i in range(len(purchasetypes)):  # I wanted to group the spending of PurchaseTypes so the PurchaseType is the
        # outer loop
        for j in range(len(purchasers)):
            values.append(sumPrice(timevalue, purchasetypes[i], purchasers[j], ""))  # adds all spending for Purchasers
            labelsarray.append(purchasers[j] + "'s " + purchasetypes[i] + " Spending")  # adds a label for the key

    total = niceValues([sum(values)])  # the total variable is for display purposes only so niceValues() is applied

    spendingchart = np.array(values)  # creates the pie chart

    plt.pie(spendingchart, labels=niceValues(values))  # plots the pie chart
    plt.legend(labels=labelsarray, loc=1)  # creates a key for the chart in the top right
    plt.title("Total : " + str(total[0]))  # total is saved as a list to input it as a parameter into niceValues()
    # but only the first and only value in the list needs to be outputted, without the [] around the string
    plt.show()  # shows the pie chart


# line graph showing spending over the given timeframe
def compareSpending3(timevalue, booleancumulative):
    try:
        matplotlib.pyplot.close()  # if the previous plot is not closed, matplotlib draws over the plot and creates
        # very confusing images
    except NameError:  # if this is the first instance of a pie chart in the running of the program, pyplot will not
        # yet have been defined so an error will occur if the closing is not part of a try and except with NameError
        pass  # nothing actually needs to be done if there is no existing pyplot

    if timevalue.lower() == "week" or timevalue.lower() == "w":  # checks to see if the given parameter specifies the
        # function should generate a graph for weeks or for days
        prefix = "wk"  # the prefix is used to communicate with the sumPrice() function
        titlevalue = "Week"  # used to display on the title of the line graph
        alldates = selectFromFile("SELECT qryPurchases.WeekOfPurchase "
                                  "FROM qryPurchases "
                                  "ORDER BY qryPurchases.WeekofPurchase ASC")  # collects all weeks from the database

    elif timevalue.lower() == "day" or timevalue.lower() == "d":  # checks to see if the given parameter specifies the
        # function should generate a graph for weeks or for days
        prefix = "dy"  # the prefix is used to communicate with the sumPrice() function
        titlevalue = "Day"  # used to display on the title of the line graph
        alldates = selectFromFile("SELECT qryPurchases.DayOfPurchase "
                                  "FROM qryPurchases "
                                  "ORDER BY qryPurchases.DayOfPurchase ASC")  # collects all days from the database.
        # The says are numbered as an integer is the qryPurchases.DayOfPurchase field

    else:
        error("Please use either 'day' or 'week' to indicate the frequency on the x axis", 65)  # if neither day nor
        # week are entered the function returns an error message as only day or week are accepted as timeframes for
        # this function
        return

    earliestvalue = alldates[0][0]  # finds the earliest value from the list
    latestvalue = alldates[len(alldates) - 1][0]  # finds the latest value from the list

    timedifference = int(latestvalue) - int(earliestvalue)  # calculates how many times the program will have to repeat

    xvalues = []  # defines values for x and y on the graph
    yvalues = []

    if booleancumulative:  # if the parameter for booleancumulative is true then the program outputs a cumulative
        # version of the graph

        cumulativecost = 0  # starts the cost at 0

        for i in range(timedifference):  # repeats for as many of the timeframe lie between the earliest value and
            # the latest one
            cumulativecost = cumulativecost + sumPrice(f"{prefix}{str(earliestvalue + i)}", "", "", "")  # adds the
            # sum of the cost for the given time to the running total cost
            yvalues.append(cumulativecost)  # adds the cost on the y-axis and increments the value on the x-axis by
            # one each time
            xvalues.append(i + 1)

    else:  # runs a non-cumulative version of the same function
        for i in range(timedifference):
            yvalues.append(sumPrice(f"{prefix}{str(earliestvalue + i)}", "", "", ""))
            xvalues.append(i + 1)

    plt.plot(xvalues, yvalues)  # plots the graph
    plt.title('Spending')  # titles the graph
    plt.xlabel(titlevalue)  # uses the label for the x-axis
    plt.ylabel('Amount Spending')  # titles the y-axis
    plt.show()  # shows the line graph


# calculates money owed over the inputted timeframe
def calculatePayments(timevalue):
    purchasers = getInfo("Purchaser")  # gets a complete list of all purchasers

    communalspending = []  # creates a list of the communal spending done by each purchaser within the timeframe
    for i in range(len(purchasers)):
        exec(f'communalspending.append(sumPrice("{timevalue}", "", "{purchasers[i]}", True))')  # appends the sum of
        # each purchaser to the communalspending list

    totalspend = (sum(communalspending))  # sums communal spending for a total value so the database does not need to
    # be interrogated again to determine the total

    assets = []  # a list of money that is owed to the purchaser and the purchaser
    liabilities = []  # a list of money that is owed by the purchaser and the purchaser

    try:
        meanspend = totalspend / len(communalspending)  # calculates the mean spend amount
    except ZeroDivisionError:  # if no money has been spent in the timeframe an error is avoided here
        meanspend = 0

    for i in range(len(communalspending)):
        if communalspending[i] > meanspend:  # if the spending of a given purchaser is greater than the mean spending
            # then they are entitled to money to compensate for spending on communal items so the value of the money
            # they spent over the mean spend of the group is added to the assets list
            assets.append([communalspending[i] - meanspend, purchasers[i]])
        elif communalspending[i] < meanspend:  # of the communal spending of the purchaser is less than the mean
            # spending then the purchaser owed someone money for spending above the mean on items that are shared
            liabilities.append([meanspend - communalspending[i], purchasers[i]])

    paymentoutput = ""  # the output does not fit in the error() function that normally displays outputs as it can be
    # an infinitely long output with infinite people, so it is outputted in the terminal
    paymentsuncalculated = True
    while paymentsuncalculated:  # repeats the process of adding to the output until all payments have been calculated
        if assets[0][0] > liabilities[0][0]:  # if the current asset is greater than the current liability then the
            # asset will become shorter but the liability can be popped from the list
            paymentoutput = paymentoutput + f"""{liabilities[0][1]} owes {niceValues([liabilities[0][0]])[0]} to {assets[0][1]} 
"""  # adds a statement of debt to the current payments
            assets[0][0] = assets[0][0] - liabilities[0][0]  # current asset must be reduced as the debt owed has
            # already been factored into the paymentoutput
            liabilities.pop(0)  # remove the first value of the liabilities list to update the current liabilities value
        elif liabilities[0][0] > assets[0][0]:  # similar principle to above but reverse the assets and liabilities.
            paymentoutput = paymentoutput + f"""{liabilities[0][1]} owes {niceValues([assets[0][0]])[0]} to {assets[0][1]}
"""  # despite the swapping of assets and liabilities in this elif statement, liabilities are still owed to assets to
            # the output is only changed by the reference to the amount owed
            liabilities[0][0] = liabilities[0][0] - assets[0][0]
            assets.pop(0)
        else:  # current asset must be equal to current liability
            paymentoutput = paymentoutput + f"""{liabilities[0][1]} owes {niceValues([liabilities[0][0]])[0]} to {assets[0][1]}
"""
            liabilities.pop(0)  # if both assets and liabilities are the same then both can be removed once they have
            # been factored into the paymentoutput value
            assets.pop(0)
        if len(assets) == 0:  # when the length of either list reaches 0 then all payments have been calculated. The
            # lists should reach length 0 at the same point, so it is only necessary to check one; assets
            paymentsuncalculated = False

    print(paymentoutput)


# shows an error message on the mainloop GUI on the bottom row in red
def error(errormessage, xposition):
    blankmessage = ""
    for i in range(256):
        blankmessage = blankmessage + " "  # creates a blank message to wipe the previous message off the GUI
    errorlbl = Label(buttonmenu, text=blankmessage, fg='#af2323', bg='#ffffff', )  # overwrites any previous message
    # with the blank message
    errorlbl.place(x=0, y=220)
    errorlbl = Label(buttonmenu, text=errormessage, fg='#af2323', bg='#ffffff',
                     font=('JetBrains Mono', 9))
    errorlbl.place(x=xposition, y=220)  # places the error message on the lowest row with the specified x-position


# mainloop for program with tkinter GUI
def main():
    global buttonmenu  # buttonmenu must be global because the error() function calls it
    buttonmenu = Tk()
    buttonmenu.title("Finances")  # titles the GUI
    buttonmenu.configure(width=500, height=260)
    buttonmenu.configure(bg='#ffffff')

    winwidth = buttonmenu.winfo_reqwidth()
    winwheight = buttonmenu.winfo_reqheight()
    posright = int(buttonmenu.winfo_screenwidth() / 2 - winwidth / 2)
    posdown = int(buttonmenu.winfo_screenheight() / 2 - winwheight / 2)
    buttonmenu.geometry("+{}+{}".format(posright, posdown))  # configuring the window

    lbl = Label(buttonmenu, text="Timeframe for calculations and graphs:", fg='#000000', bg='#ffffff',
                font=('JetBrains Mono', 9))  # input for all timeframes
    lbl.place(x=150, y=12)

    timeframe = StringVar()  # links the form to the variable timeframe

    timeframeentry = Entry(buttonmenu, bd=2, textvariable=timeframe)
    timeframeentry.place(x=200, y=30)

    btn1 = Button(buttonmenu, text="Compare Spending On Personal and Communal Items",
                  command=lambda: compareSpending1(timeframe.get()),
                  bg='#ffffff')  # creates a button for creating a pie chart comparing communal and personal spending
    btn1.place(x=110, y=70)
    btn2 = Button(buttonmenu, text="      Compare Spending On Different Types of Items       ",
                  command=lambda: compareSpending2(timeframe.get()),
                  bg='#ffffff')  # creates a button for creating a pie chart comparing spending on PurchaseTypes
    btn2.place(x=110, y=100)
    btn3 = Button(buttonmenu, text="                        Look at Spending Over Time                      ",
                  command=lambda: compareSpending3(timeframe.get(), False),
                  bg='#ffffff')  # creates a button for creating a line graph showing non-cumulative spending
    btn3.place(x=110, y=130)
    btn4 = Button(buttonmenu, text="           Look at Spending Over Time (Cumulative)           ",
                  command=lambda: compareSpending3(timeframe.get(), True),
                  bg='#ffffff')  # creates a button for creating a line graph showing cumulative spending
    btn4.place(x=110, y=160)
    btn5 = Button(buttonmenu, text="                              Calculate Payments                               ",
                  command=lambda: calculatePayments(timeframe.get()),
                  bg='#ffffff')  # Calculates money owed between users based on communal spending
    btn5.place(x=110, y=190)

    buttonmenu.mainloop()


global filelocation  # addition for users other than myself - enter the location of the folder (not the document)
# filelocation is used in line 24 to specify the document area if you would like to hardcode in your file location
filelocation = input("Please specify the location of the 'Finances Project' Folder ")


main()
