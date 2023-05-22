def MedAssociatesLOCO_Script():

    import xlwt
    from tempfile import TemporaryFile
    book = xlwt.Workbook()

    #TextFile = input("Enter the Med Associates text (.txt) file name to me converted into an Exel file:")
    ExcelFile = input("Enter the name of the new Excel file to be created that will contain the converted data:")
    
    sheet1 = book.add_sheet('sheet1')
    #TextFile can be changes to meet testing needs
    TextFile = 'MT14NBNJLOCOSummary2'

    sheet1.write(0,0, TextFile)
    sheet1.write(1,0, 'RunDate')
    sheet1.write(1,1, 'RatID')
    sheet1.write(1,2, 'BoxID')
    sheet1.write(1,15, 'TotDist')


    # The number of lines the text file are shifted: line_shift = s.
    # For each subject data gets collected, the shift increases by 1 or s = s + 1 (s += 1).
    # Because of the way the data is structured, there is an additional shift within each subjet's data
    # not present in the first subject's data but must also be compensated for by secondary line_shift = t

    s = 0
    line_shift = s

    anchor = 41
    bins = 12
    anc_bin = anchor+bins

    specific_lines_dist_trav_shift = []
    for line in range(anchor, anc_bin):
        line += s
        specific_lines_dist_trav_shift.append(line)

    # Subject 1
    medtxtfile = open(TextFile+'.txt')
    five_min_time_bin_List = []

    def remove_space(slicedp1):
        return "".join(slicedp1.split())

    print("Sum of dist. trav. (cm per 5min)")
    for pos, A_num in enumerate(medtxtfile):
        A_num = A_num.rstrip()
        if pos in specific_lines_dist_trav_shift:
            #print(A_num)
            slicedp1 = A_num [0:9]
            slcflt = (float(remove_space(slicedp1)))
            five_min_time_bin_List.append(slcflt)
            print(slcflt)

    for i, e in enumerate(five_min_time_bin_List,start=3):
        sheet1.write(2,i,e)

    medtxtfile = open(TextFile+'.txt')
    content = medtxtfile.readlines()

    RunDate_line = content[2]
    RDS = slice(10,21)
    sheet1.write(2,0,RunDate_line[RDS])

    RatID_line = content[27 + s]
    IDS = slice(30,38)
    sheet1.write(2,1,RatID_line[IDS])

    Box_line = content[21]
    BS = slice(30,38)
    sheet1.write(2,2, Box_line[BS])

    TotDist_line = content[72 + s]
    TD = slice(21,32)
    sheet1.write(2,15,TotDist_line[TD])

    for col in range(3,15,1):
        binnum = str(col - 2)
        sheet1.write(1,col,'Bin'+binnum)

    # Subject 2

    if line_shift > 0:
        s = line_shift + 1
        t = s
    else: 
        s = 0
        t = 1

    anchor = 126
    bins = 12
    anc_bin = anchor+bins

    specific_lines_dist_trav_shift = []
    for line in range(anchor, anc_bin):
        line += s
        specific_lines_dist_trav_shift.append(line)

    medtxtfile = open(TextFile+'.txt')
    five_min_time_bin_List = []

    def remove_space(slicedp2):
        return "".join(slicedp2.split())

    print("Sum of dist. trav. (cm per 5min)")
    for pos, A_num in enumerate(medtxtfile):
        A_num = A_num.rstrip()
        if pos in specific_lines_dist_trav_shift:
            #print(A_num)
            slicedp2 = A_num [0:9]
            slcflt2 = (float(remove_space(slicedp2)))
            five_min_time_bin_List.append(slcflt2)
            print(slcflt2)

    for i, e in enumerate(five_min_time_bin_List, start=3):
        sheet1.write(3,i,e)

    medtxtfile = open(TextFile+'.txt')
    content = medtxtfile.readlines()

    RunDate_line = content[87 + (t - 1)]
    RDS = slice(10,21)
    sheet1.write(3,0, RunDate_line[RDS])

    RatID_line = content[112 + s]
    IDS = slice(30,38)
    sheet1.write(3,1, RatID_line[IDS])

    Box_line = content[106 + (t - 1)]
    BS = slice(30,38)
    sheet1.write(3,2, Box_line[BS])

    TotDist_line = content[157 + s]
    TD = slice(21,32)
    sheet1.write(3,15, TotDist_line[TD])


    # Subject 3

    if line_shift > 0:
        s = line_shift + 2
        t = s
    else: 
        s = 0
        t = 1
    anchor = 211
    bins = 12
    anc_bin = anchor+bins

    specific_lines_dist_trav_shift = []
    for line in range(anchor, anc_bin):
        line += s
        specific_lines_dist_trav_shift.append(line)
    medtxtfile = open(TextFile+'.txt')
    five_min_time_bin_List = []

    def remove_space(slicedp3):
        return "".join(slicedp3.split())

    print("Sum of dist. trav. (cm per 5min)")
    for pos, A_num in enumerate(medtxtfile):
        A_num = A_num.rstrip()
        if pos in specific_lines_dist_trav_shift:
            #print(A_num)
            slicedp3 = A_num [0:9]
            slcflt3 = (float(remove_space(slicedp3)))
            five_min_time_bin_List.append(slcflt3)
            print(slcflt3)

    for i, e in enumerate(five_min_time_bin_List, start=3):
        sheet1.write(4,i,e)

    medtxtfile = open(TextFile+'.txt')
    content = medtxtfile.readlines()

    RunDate_line = content[172 + (t - 1)]
    RDS = slice(10,21)
    sheet1.write(4,0, RunDate_line[RDS])

    RatID_line = content[197 + s]
    IDS = slice(30,38)
    sheet1.write(4,1, RatID_line[IDS])

    Box_line = content[191 + (t - 1)]
    BS = slice(30,38)
    sheet1.write(4,2, Box_line[BS])

    TotDist_line = content[242 + s]
    TD = slice(21,32)
    sheet1.write(4,15, TotDist_line[TD])


    # Subject 4
    if line_shift > 0:
        s = line_shift + 3
        t = s
    else: 
        s = 0
        t = 1

    anchor = 296
    bins = 12
    anc_bin = anchor+bins

    specific_lines_dist_trav_shift = []
    for line in range(anchor, anc_bin):
        line += s
        specific_lines_dist_trav_shift.append(line)
    medtxtfile = open(TextFile+'.txt')
    five_min_time_bin_List = []

    def remove_space(slicedp4):
        return "".join(slicedp4.split())

    print("Sum of dist. trav. (cm per 5min)")
    for pos, A_num in enumerate(medtxtfile):
        A_num = A_num.rstrip()
        if pos in specific_lines_dist_trav_shift:
            #print(A_num)
            slicedp4 = A_num [0:9]
            slcflt4 = (float(remove_space(slicedp4)))
            five_min_time_bin_List.append(slcflt4)
            print(slcflt4)

    for i, e in enumerate(five_min_time_bin_List, start=3):
        sheet1.write(5,i,e)

    medtxtfile = open(TextFile+'.txt')
    content = medtxtfile.readlines()

    RunDate_line = content[257 + (t - 1)]
    RDS = slice(10,21)
    sheet1.write(5,0, RunDate_line[RDS])

    RatID_line = content[282 + s]
    IDS = slice(30,38)
    sheet1.write(5,1, RatID_line[IDS])

    Box_line = content[276 +  (t - 1)]
    BS = slice(30,38)
    sheet1.write(5,2, Box_line[BS])

    TotDist_line = content[327 + s]
    TD = slice(21,32)
    sheet1.write(5,15, TotDist_line[TD])


    # Subject 5

    if line_shift > 0:
        s = line_shift + 4
        t = s
    else: 
        s = 0
        t = 1

    anchor = 381
    bins = 12
    anc_bin = anchor+bins

    specific_lines_dist_trav_shift = []
    for line in range(anchor, anc_bin):
        line += s
        specific_lines_dist_trav_shift.append(line)
    medtxtfile = open(TextFile+'.txt')
    five_min_time_bin_List = []

    def remove_space(slicedp4):
        return "".join(slicedp4.split())

    print("Sum of dist. trav. (cm per 5min)")
    for pos, A_num in enumerate(medtxtfile):
        A_num = A_num.rstrip()
        if pos in specific_lines_dist_trav_shift:
            #print(A_num)
            slicedp4 = A_num [0:9]
            slcflt4 = (float(remove_space(slicedp4)))
            five_min_time_bin_List.append(slcflt4)
            print(slcflt4)

    for i, e in enumerate(five_min_time_bin_List, start=3):
        sheet1.write(6,i,e)

    medtxtfile = open(TextFile+'.txt')
    content = medtxtfile.readlines()

    RunDate_line = content[342 + (t - 1)]
    RDS = slice(10,21)
    sheet1.write(6,0, RunDate_line[RDS])

    RatID_line = content[367 + s]
    IDS = slice(30,38)
    sheet1.write(6,1, RatID_line[IDS])

    Box_line = content[361 + (t - 1)]
    BS = slice(30,38)
    sheet1.write(6,2, Box_line[BS])

    TotDist_line = content[412 + s]
    TD = slice(21,32)
    sheet1.write(6,15, TotDist_line[TD])


    # Subject 6

    if line_shift > 0:
        s = line_shift + 5
        t = s
    else: 
        s = 0
        t = 1

    anchor = 466
    bins = 12
    anc_bin = anchor+bins

    specific_lines_dist_trav_shift = []
    for line in range(anchor, anc_bin):
        line += s
        specific_lines_dist_trav_shift.append(line)
    medtxtfile = open(TextFile+'.txt')
    five_min_time_bin_List = []

    def remove_space(slicedp4):
        return "".join(slicedp4.split())

    print("Sum of dist. trav. (cm per 5min)")
    for pos, A_num in enumerate(medtxtfile):
        A_num = A_num.rstrip()
        if pos in specific_lines_dist_trav_shift:
            #print(A_num)
            slicedp4 = A_num [0:9]
            slcflt4 = (float(remove_space(slicedp4)))
            five_min_time_bin_List.append(slcflt4)
            print(slcflt4)

    for i, e in enumerate(five_min_time_bin_List, start=3):
        sheet1.write(7,i,e)

    medtxtfile = open(TextFile+'.txt')
    content = medtxtfile.readlines()

    RunDate_line = content[427 + (t - 1)]
    RDS = slice(10,21)
    sheet1.write(7,0, RunDate_line[RDS])

    RatID_line = content[452 + s]
    IDS = slice(30,38)
    sheet1.write(7,1, RatID_line[IDS])

    Box_line = content[446 + (t - 1)]
    BS = slice(30,38)
    sheet1.write(7,2, Box_line[BS])

    TotDist_line = content[497 + s]
    TD = slice(21,32)
    sheet1.write(7,15, TotDist_line[TD])

    # Subject 7

    if line_shift > 0:
        s = line_shift + 6
        t = s
    else: 
        s = 0
        t = 1

    anchor = 551
    bins = 12
    anc_bin = anchor+bins

    specific_lines_dist_trav_shift = []
    for line in range(anchor, anc_bin):
        line += s
        specific_lines_dist_trav_shift.append(line)
    medtxtfile = open(TextFile+'.txt')
    five_min_time_bin_List = []

    def remove_space(slicedp4):
        return "".join(slicedp4.split())

    print("Sum of dist. trav. (cm per 5min)")
    for pos, A_num in enumerate(medtxtfile):
        A_num = A_num.rstrip()
        if pos in specific_lines_dist_trav_shift:
            #print(A_num)
            slicedp4 = A_num [0:9]
            slcflt4 = (float(remove_space(slicedp4)))
            five_min_time_bin_List.append(slcflt4)
            print(slcflt4)

    for i, e in enumerate(five_min_time_bin_List, start=3):
        sheet1.write(8,i,e)

    medtxtfile = open(TextFile+'.txt')
    content = medtxtfile.readlines()

    RunDate_line = content[512 + (t - 1)]
    RDS = slice(10,21)
    sheet1.write(8,0, RunDate_line[RDS])

    RatID_line = content[537 + s]
    IDS = slice(30,38)
    sheet1.write(8,1, RatID_line[IDS])

    Box_line = content[531 + (t - 1)]
    BS = slice(30,38)
    sheet1.write(8,2, Box_line[BS])

    TotDist_line = content[582 + s]
    TD = slice(21,32)
    sheet1.write(8,15, TotDist_line[TD])

    # Subject 8

    if line_shift > 0:
        s = line_shift + 7
        t = s
    else: 
        s = 0
        t = 1
    anchor = 636
    bins = 12
    anc_bin = anchor+bins

    specific_lines_dist_trav_shift = []
    for line in range(anchor, anc_bin):
        line += s
        specific_lines_dist_trav_shift.append(line)
    medtxtfile = open(TextFile+'.txt')
    five_min_time_bin_List = []

    def remove_space(slicedp4):
        return "".join(slicedp4.split())

    print("Sum of dist. trav. (cm per 5min)")
    for pos, A_num in enumerate(medtxtfile):
        A_num = A_num.rstrip()
        if pos in specific_lines_dist_trav_shift:
            #print(A_num)
            slicedp4 = A_num [0:9]
            slcflt4 = (float(remove_space(slicedp4)))
            five_min_time_bin_List.append(slcflt4)
            print(slcflt4)

    for i, e in enumerate(five_min_time_bin_List, start=3):
        sheet1.write(9,i,e)

    medtxtfile = open(TextFile+'.txt')
    content = medtxtfile.readlines()

    RunDate_line = content[597 + (t - 1)]
    RDS = slice(10,21)
    sheet1.write(9,0, RunDate_line[RDS])

    RatID_line = content[622 + s]
    IDS = slice(30,38)
    sheet1.write(9,1, RatID_line[IDS])

    Box_line = content[616 + (t - 1)]
    BS = slice(30,38)
    sheet1.write(9,2, Box_line[BS])

    TotDist_line = content[667 + s]
    TD = slice(21,32)
    sheet1.write(9,15, TotDist_line[TD])

    # Subject 9

    if line_shift > 0:
        s = line_shift + 8
        t = s
    else: 
        s = 0
        t = 1

    anchor = 721
    bins = 12
    anc_bin = anchor+bins

    specific_lines_dist_trav_shift = []
    for line in range(anchor, anc_bin):
        line += s
        specific_lines_dist_trav_shift.append(line)
    medtxtfile = open(TextFile+'.txt')
    five_min_time_bin_List = []

    def remove_space(slicedp4):
        return "".join(slicedp4.split())

    print("Sum of dist. trav. (cm per 5min)")
    for pos, A_num in enumerate(medtxtfile):
        A_num = A_num.rstrip()
        if pos in specific_lines_dist_trav_shift:
            #print(A_num)
            slicedp4 = A_num [0:9]
            slcflt4 = (float(remove_space(slicedp4)))
            five_min_time_bin_List.append(slcflt4)
            print(slcflt4)

    for i, e in enumerate(five_min_time_bin_List, start=3):
        sheet1.write(10,i,e)

    medtxtfile = open(TextFile+'.txt')
    content = medtxtfile.readlines()

    RunDate_line = content[682 + (t - 1)]
    RDS = slice(10,21)
    sheet1.write(10,0, RunDate_line[RDS])

    RatID_line = content[707 + s]
    IDS = slice(30,38)
    sheet1.write(10,1, RatID_line[IDS])

    Box_line = content[701 + (t - 1)]
    BS = slice(30,38)
    sheet1.write(10,2, Box_line[BS])

    TotDist_line = content[752 + s]
    TD = slice(21,32)
    sheet1.write(10,15, TotDist_line[TD])

    # Subject 10

    if line_shift > 0:
        s = line_shift + 9
        t = s
    else: 
        s = 0
        t = 1

    anchor = 806
    bins = 12
    anc_bin = anchor+bins

    specific_lines_dist_trav_shift = []
    for line in range(anchor, anc_bin):
        line += s
        specific_lines_dist_trav_shift.append(line)
    medtxtfile = open(TextFile+'.txt')
    five_min_time_bin_List = []

    def remove_space(slicedp4):
        return "".join(slicedp4.split())

    print("Sum of dist. trav. (cm per 5min)")
    for pos, A_num in enumerate(medtxtfile):
        A_num = A_num.rstrip()
        if pos in specific_lines_dist_trav_shift:
            #print(A_num)
            slicedp4 = A_num [0:9]
            slcflt4 = (float(remove_space(slicedp4)))
            five_min_time_bin_List.append(slcflt4)
            print(slcflt4)

    for i, e in enumerate(five_min_time_bin_List, start=3):
        sheet1.write(11,i,e)

    medtxtfile = open(TextFile+'.txt')
    content = medtxtfile.readlines()

    RunDate_line = content[767 + (t - 1)]
    RDS = slice(10,21)
    sheet1.write(11,0, RunDate_line[RDS])

    RatID_line = content[792 + s]
    IDS = slice(30,38)
    sheet1.write(11,1, RatID_line[IDS])

    Box_line = content[786 + (t - 1)]
    BS = slice(30,38)
    sheet1.write(11,2, Box_line[BS])

    TotDist_line = content[837 + s]
    TD = slice(21,32)
    sheet1.write(11,15, TotDist_line[TD])

    name = ExcelFile+'.xls'
    book.save(name)
    book.save(TemporaryFile())

    e = input("To convert another Med Associates Activity Tracker text file, Enter 'y'\n"
    "To exit this program, press any other key.")
    if e == "y":            
            MedAssociatesLOCO_Script()
MedAssociatesLOCO_Script()