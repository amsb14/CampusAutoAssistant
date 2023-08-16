import xlsxwriter, os
import pandas as pd
import io

parent_path = os.path.join(os.getcwd())
sheet_name = 'جداول الطلاب'
# excel_file = rf'{parent_path}\{sheet_name}.xlsx'
time_cells_dict = {
    "08": "BC", "09": "DE", "10": "FG", "11": "HI",
    "12": "JK", "13": "LM", "14": "NO", "15": "PQ", "16": "RS", "17": "TU"
}
day_cells_dict = {
    "الأحد": "14", "الإثنين": "18", "الثلاثاء": "22", "الاربعاء": "26", "الخميس": "30"
}
    
    
def get_term_text(term):
    if term[-2:-1] == '1':
        return "الأول"
    elif term[-2:-1] == '2':
        return "الثاني"
    elif term[-2:-1] == '3':
        return "الثالث"

def studentName(userID):
    df_new = dfsf24[dfsf24['رقم الطالب'] == (userID)]
    return df_new['إسم الطالب'].iloc[0]

# def comIDs():
#     IDs = list(set(dfsf24['رقم الطالب']))
#     return IDs

def get_lab_department(department='all'):
    """Return computer ids"""
    if department != "all":
        IDs = dfsf24.loc[dfsf24['القسم'] == department, 'رقم الطالب'].unique().tolist()
    else:
        IDs = dfsf24['رقم الطالب'].unique().tolist()
    return IDs


def split(txt):
    res = [i.split('\n') for i in txt][0]
    return list(map(str.strip, res))

def removeNonValidTimeSlot(timeslots, *arguments):
    Error = '18'
    for i in range(len(timeslots) - 1, -1, -1):
        if Error not in timeslots[i]:
            pass
        else:
            for x in arguments:
                x.pop(i)
            timeslots.pop(i)
    return timeslots

def merge_cells(timeslot, day):
    x = timeslot.split("-")
    start_time = x[1].strip()
    end_time = x[0].strip()
    s = start_time[:2]
    e = end_time[:2]
    starting_cell = time_cells_dict[s][0]
    ending_cell = time_cells_dict[e][-1]
    day_row = day_cells_dict[day]
    start_and_end = starting_cell + ending_cell
    return f"{starting_cell}{day_row}:{ending_cell}{int(day_row)+3}", start_and_end

  
def ss01Details(userID):
    teachers, subjects, crn, classrooms, days, times = [], [], [], [], [], []
    df_new = dfsf24[dfsf24['رقم الطالب'] == (userID)]
    ref_subject_id = df_new['الرقم المرجعي'].tolist()
    for ref_id in ref_subject_id:
        df_temp = dfss01[dfss01['الرقم المرجعي'] == ref_id]
        teachers.extend(df_temp['اسم المدرب'].tolist())
        subjects.extend(df_temp['اسم المقرر'].tolist())
        crn.extend(df_temp['الرقم المرجعي'].tolist())
        times.extend(df_temp['أوقات'].tolist())
        classrooms.extend(df_temp['قاعة'].tolist())
        days.extend(df_temp['أيام'].tolist())
        days = [d.strip() for d in days]
    major = df_new['القسم'].iloc[0]
    return teachers, subjects, crn, classrooms, times, days, major


def create_excel_file(student_id):

    worksheet = workbook.add_worksheet(f"{student_id}") # add a new worksheet
    
    
    teachers, subs, refs, labs, times, days, major = ss01Details(student_id)
    result = removeNonValidTimeSlot(times, subs, refs, labs, days)


    def total_hours(s, e):
        sum = ord(e) - ord(s) + 1
        sum //= 2 
        return sum
    
    def merge_format(back_color, size, font='black'):
        merge_format = workbook.add_format({
        'bold':     True,
        'font_name': 'Calibri',
        'font_size': f'{size}',
        'border':   6,
        'font_color':f'{font}',
        'align':    'center',
        'valign':   'vcenter',
        'fg_color': f'{back_color}',
        'text_wrap': True
        })
        return merge_format
    
    def no_Border(back_color, size, font='black'):
        no_border = workbook.add_format({
        'bold':     True,
        'font_name': 'Calibri',
        'font_size': f'{size}',
        'border':   0,
        'font_color':f'{font}',
        'align':    'center',
        'valign':   'vcenter',
        'fg_color': f'{back_color}',
        })
        return no_border
    
    def merge_format2(back_color, size, font='black', border=3):
        merge_format = workbook.add_format({
        'bold':     True,
        'font_name': 'Calibri',
        'font_size': f'{size}',
        'border':   int(border),
        'font_color':f'{font}',
        'align':    'center',
        'valign':   'vcenter',
        'fg_color': f'{back_color}',
        'text_wrap': True
        })
        return merge_format
    
    timeslots = []
    totalhours = []
    for t, d in zip(times, days):
        whatcell, traininghours = merge_cells(t, d)
        timeslots.append(whatcell)
        totalhours.append(total_hours(traininghours[0], traininghours[1]))

    list_of_alphabets = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U"]
    for letter in list_of_alphabets:
        for i in range(14,34+1):
            worksheet.write(f"{letter}{i}:{letter}{i}", "", merge_format2("#FFFFFF", 9))

    for teacher, sub, ref, lab, slot in zip(teachers, subs, refs, labs, timeslots):
        worksheet.merge_range(f"{slot}", f'{sub}\n{ref}\n{str(lab)[-3:]}\n{teacher}',  merge_format('#ADD5B0', '9'))

    
    
    colV = workbook.add_format()
    colV.set_left(6)
    
    line29 = workbook.add_format()
    line29.set_top(6)
    
    
    # Set column size
    worksheet.set_column('A:A', 5.71) 
    worksheet.set_column('B:U', 2.86) 

    # Insert image (tvtc)
    worksheet.insert_image('G1', 'icon/tvtc.jpg', {'x_scale': 0.3, 'y_scale': 0.3, 'x_offset': 0,'y_offset': 0})
    

    worksheet.merge_range("A1:F1", "المملكة العربية السعودية" , no_Border('#FFFFFF', 9))
    worksheet.merge_range("A2:F2", "المؤسسة العامة للتدريب التقني والمهني" , no_Border('#FFFFFF', 9))
    worksheet.merge_range("A3:F3", "الكلية التقنية بمحافظة حقل" , no_Border('#FFFFFF', 9))

    worksheet.merge_range("M1:U1", "Kingdom of Saudi Arabia" ,no_Border('#FFFFFF', 9))
    worksheet.merge_range("M2:U2", "Technical and Vocational Training Corporation" ,no_Border('#FFFFFF', 9))
    worksheet.merge_range("M3:U3", "College of Technology in Haql" ,no_Border('#FFFFFF', 9))
    
    worksheet.merge_range("A5:U6", "جدول المتدرب" ,merge_format('#636467', 14, font='white'))
    
    worksheet.merge_range("A8:B8", "اسم المتدرب" ,merge_format('#2FBAB3', '9'))
    worksheet.merge_range("C8:J8", f"{studentName(student_id)}" ,merge_format('#FFFFFF', '9'))
    worksheet.merge_range("K8:M8", "رقم المتدرب" ,merge_format('#2FBAB3', '9'))
    worksheet.merge_range("N8:P8", f"{student_id}" ,merge_format('#FFFFFF', '9'))
    worksheet.merge_range("Q8:S8", "الفصل التدريبي" ,merge_format('#2FBAB3', '9'))
    worksheet.merge_range("T8:U8", "الثالث" ,merge_format('#FFFFFF', '9'))

    worksheet.merge_range("A9:B9", "القسم" ,merge_format('#2FBAB3', '9'))
    worksheet.merge_range("C9:J9", major ,merge_format('#FFFFFF', '9'))
    worksheet.merge_range("K9:M9", "المؤهل" ,merge_format('#2FBAB3', '9'))
    worksheet.merge_range("N9:U9", "الدبلوم" ,merge_format('#FFFFFF', '9'))

    
    worksheet.write("A11", "المحاضرة",merge_format('#2FBAB3', '9'))
    worksheet.merge_range("A12:A13", "الوقت",merge_format('#2FBAB3', '9'))
    worksheet.merge_range("A14:A17", "الأحد",merge_format('#2FBAB3', '9'))
    worksheet.merge_range("A18:A21", "الإثنين",merge_format('#2FBAB3', '9'))
    worksheet.merge_range("A22:A25", "الثلاثاء",merge_format('#2FBAB3', '9'))
    worksheet.merge_range("A26:A29", "الأربعاء",merge_format('#2FBAB3', '9'))
    worksheet.merge_range("A30:A33", "الخميس",merge_format('#2FBAB3', '9'))
    
    worksheet.merge_range("V14:V33", "",colV)

            
    worksheet.merge_range("A34:U34", "",line29)
    
    worksheet.merge_range("A35:C35", "ساعات الإتصال",merge_format('#2FBAB3', '9'))
    worksheet.merge_range("D35:H35", f"{sum(totalhours)}",merge_format('#FFFFFF', '9'))
    
    worksheet.merge_range("A36:C36", "اخر تعديل",merge_format('#2FBAB3', '9'))
    import datetime
    now = datetime.datetime.now()
    worksheet.merge_range("D36:H36", now.strftime("%I:%M%p - %d/%m/%Y"),merge_format('#FFFFFF', '9'))
    
    row = 11
    t_start = ["08:00","09:00","10:00","11:00","12:00","13:00","14:00","15:00","16:00","17:00"]	
    t_end = ["08:40","09:40","10:40","11:40","12:40","13:40","14:40","15:40","16:40","17:40"]	
    


    for index, i in enumerate(time_cells_dict, start =1 ):
        worksheet.merge_range(f"{time_cells_dict[i][0]}{row}:{time_cells_dict[i][1]}{row}", f'{index}',  merge_format('#2FBAB3', '9')) 


            
    row += 1
    for i, s, e in zip(time_cells_dict, t_start, t_end ):
        worksheet.merge_range(f"{time_cells_dict[i][0]}{row}:{time_cells_dict[i][1]}{row}", f'{s}',  merge_format('#2FBAB3', '9'))
        worksheet.merge_range(f"{time_cells_dict[i][0]}{row+1}:{time_cells_dict[i][1]}{row+1}", f'{e}',  merge_format('#2FBAB3', '9'))
         
        
def run(file1, file2, department):
    global workbook, dfss01, dfsf24
    dfss01 = pd.read_csv(file1)
    dfsf24 = pd.read_csv(file2)
    # Check if the number of columns is 20, otherwise raise an exception
    if len(dfss01.columns) != 20:
        raise ValueError("Make sure you upload the correct file (SS01) from Rayat!")
    if len(dfsf24.columns) != 20:
        raise ValueError("Make sure you upload the correct file (SF24) from Rayat!")
    LIST_OF_STUDENT_ID = get_lab_department(department)  
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})  
    for STUDENT_ID in LIST_OF_STUDENT_ID:
        create_excel_file(STUDENT_ID) 
    workbook.close()
    output.seek(0)
    return output
    

    
    
    
    
    
    
    
    
    
    
    
    
    
