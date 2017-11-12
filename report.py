import xlrd
import datetime


def open_file(path):
    """
    #######################################
    #     Open and read an Excel file     #
    #                                     #
    #              Coded By.              #
    #            Ashraf Amer              #
    #            Oct,30 2017              #
    #                                     #
    #######################################
    """

    # open file
    book = xlrd.open_workbook(path)
 
    # get the first worksheet
    first_sheet = book.sheet_by_index(0)

    # print number of sheets
    print ("The number of sheets: =>", (book.nsheets))
 
    # print sheet names
    print ("sheet names => ", (book.sheet_names()))

    # output reading data array
    read_output = []
    # read excel file
    for x in range(1, 15708):
        # read seasons [k]
        Seasons = first_sheet.col_values(10)[x]
        # Output times for each season
        date = (first_sheet.col_values(2)[x])
        # convert xltime to timedate => from numeric value to real time
        t_value = (xlrd.xldate_as_tuple(date,datemode=0))
        # convert time to mins =>(hours*60 , mins/60)
        time_min = (t_value[3]*60) + t_value[4] + (t_value[5] / 60)
        # add reading data to read output array
        read_output.append((Seasons,time_min))
    

    # deal with output
    all_season = []
    for key,val in read_output: 
        all_season.append(key)

    # count all similar seassons
    my_dict = {i:all_season.count(i) for i in all_season}

    # print seasons' count , count of seassons
    print ("Seasons' Count: =>", my_dict)
    print ("Number of Seasons: =>", len(my_dict))

    # save count values for each season
    my_values = []
    # the average length of each season
    avg_len = []
    for item in my_dict:
        my_values.append(my_dict[item])
        # sum hours for each season
        sum_mins = []
        for k,v in read_output:
            if (k == item):
                sum_mins.append(v)
        avg_min = sum(sum_mins) / len (sum_mins)
        # conver total_time to hours:mins:sec
        time_hours = datetime.timedelta(minutes=avg_min)
        avg_len.append(time_hours)     


    # index inside for loop
    x = 0
    longestvalue = 0
    shortestvalue = 0

    # print average lenght for each seasson [Time]
    for value in my_dict:
        if my_dict[value] == max(my_values):
            longestvalue = value
        elif my_dict[value] == min(my_values):
            shortestvalue = value
        # print average length for each season    
        print("The Average length for [", value, "] is: =>", avg_len[x])
        x = x+1
    
    # print shortest and longest val. [Seasons' Count]
    print("Longest Season is: =>", longestvalue)
    print("Shortest Season is: =>", shortestvalue)


  
if __name__ == "__main__":
    path = "test.xls"
    open_file('[File-Name].xlsx')
