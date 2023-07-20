import xlsxwriter, os
import pandas as pd
import io
    
global TIME_CELLS_DICT, DAY_CELLS_DICT , RAYAT_FILE , EXCEL_FILE , PARENT_PATH

PARENT_PATH = os.path.join(os.getcwd())
RAYAT_FILE  = rf'{PARENT_PATH}\SS01.csv'
SHEET_NAME  = 'جداول المدربين'
EXCEL_FILE  = rf'{PARENT_PATH}\{SHEET_NAME }.xlsx'
TIME_CELLS_DICT = {"08":"BC", "09":"DE", "10":"FG", "11":"HI",
                   "12":"JK", "13":"LM", "14":"NO", "15":"PQ", "16":"RS","17":"TU"}

DAY_CELLS_DICT = {"الأحد":"14","الإثنين":"17","الثلاثاء":"20","الاربعاء":"23","الخميس":"26"}
  

def teacherName(userID): 
    """Return a teacher name based on his computer ID number"""
    df_new = df[df['رقم المدرب'] == (userID)]
    teacher_name = df_new['اسم المدرب'].iloc[0]
    
    return teacher_name


def get_lab_department(department='all'):
    """Return computer ids"""
    if department != "all":
        IDs = df.loc[df['القسم'] == department, 'رقم المدرب'].unique().tolist()
    else:
        IDs = df['رقم المدرب'].unique().tolist()
    return IDs

    
def split(txt):
    res = [i.split('\n') for i in txt][0]
    stripped = list(map(str.strip, res))

    return stripped

def removeNonValidTimeSlot(timeslots, *arguments):
    Error = '18'
    for i in range(len(timeslots)-1, -1, -1):
        if not Error in timeslots[i]: 
            pass
        else:
            for x in arguments:
                x.pop(i)
            timeslots.pop(i)
    return timeslots



def merge_cells(timeslot, day):
    x = timeslot.split("-")
    start_time   = x[1].strip()
    end_time     = x[0].strip()
    
    s = start_time[:2]
    e = end_time[:2]
    starting_cell = TIME_CELLS_DICT[s][0]
    ending_cell   = TIME_CELLS_DICT[e][-1]
    day_row = DAY_CELLS_DICT[day]
    
    start_and_end = starting_cell + ending_cell
    merge = f"{starting_cell}{day_row}:{ending_cell}{int(day_row)+2}"
    
    return merge, start_and_end

      
def ss01Details(userID):
    subjects, subject_reference, lab_id, days, times = [], [], [], [], []
    
    df_new = df[df['رقم المدرب'] == (userID)]
    subject_name = df_new['اسم المقرر'].to_string(index=False).strip()
    ref_subject_id = df_new['الرقم المرجعي'].to_string(index=False).strip()
    laboratories = df_new['قاعة'].to_string(index=False).strip()
    lecture_times = df_new['أوقات'].to_string(index=False).strip()
    lecture_days = df_new['أيام'].to_string(index=False).strip()
    
    
    subjects.append(subject_name)
    subjects = split(subjects)
    
    subject_reference.append(ref_subject_id)
    subject_reference =split(subject_reference)
    
    lab_id.append(laboratories)
    lab_id = split(lab_id)
    
    times.append(lecture_times)
    times = split(times)
    
    days.append(lecture_days)
    days = split(days)
    
    return subjects, subject_reference, lab_id, times, days


def create_excel_file(teacher, computer_id):
    
    worksheet = workbook.add_worksheet(f"{teacher}") # add a new worksheet
    
    subs, refs, labs, times, days = ss01Details(computer_id)
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
        for i in range(14,28+1):
            worksheet.write(f"{letter}{i}:{letter}{i}", "", merge_format2("#FFFFFF", 9))

    for sub, ref, lab, slot in zip(subs, refs, labs, timeslots):
        worksheet.merge_range(f"{slot}", f'{sub}\n{ref}\n{lab}',  merge_format('#E0E0E0', '9'))
  
    
    colV = workbook.add_format()
    colV.set_left(6)
    
    line29 = workbook.add_format()
    line29.set_top(6)
    
    worksheet.set_column('A:A', 5.71) 
    worksheet.set_column('B:U', 2.86) 

    worksheet.insert_image('G1', '/icon/tvtc.jpg', {'x_scale': 0.3, 'y_scale': 0.3, 'x_offset': 0,'y_offset': 0})
    

    worksheet.merge_range("A1:F1", "المملكة العربية السعودية" , no_Border('#FFFFFF', 9))
    worksheet.merge_range("A2:F2", "المؤسسة العامة للتدريب التقني والمهني" , no_Border('#FFFFFF', 9))
    worksheet.merge_range("A3:F3", "الكلية التقنية بمحافظة حقل" , no_Border('#FFFFFF', 9))

    worksheet.merge_range("M1:U1", "Kingdom of Saudi Arabia" ,no_Border('#FFFFFF', 9))
    worksheet.merge_range("M2:U2", "Technical and Vocational Training Corporation" ,no_Border('#FFFFFF', 9))
    worksheet.merge_range("M3:U3", "College of Technology in Haql" ,no_Border('#FFFFFF', 9))
    
    worksheet.merge_range("A5:U6", "الجدول التدريبي" ,merge_format('#FFFFFF', 14))
    
    worksheet.merge_range("A8:B8", "اسم المدرب" ,merge_format('#E0E0E0', '9'))
    worksheet.merge_range("C8:J8", f"{teacherName(computer_id)}" ,merge_format('#FFFFFF', '9'))
    worksheet.merge_range("K8:M8", "رقم الحاسب" ,merge_format('#E0E0E0', '9'))
    worksheet.merge_range("N8:P8", f"00{computer_id}" ,merge_format('#FFFFFF', '9'))
    worksheet.merge_range("Q8:S8", "رقم المكتب" ,merge_format('#E0E0E0', '9'))
    worksheet.merge_range("T8:U8", "" ,merge_format('#FFFFFF', '9'))

    worksheet.merge_range("A9:B9", "الوظيفة" ,merge_format('#E0E0E0', '9'))
    worksheet.merge_range("C9:D9", "" ,merge_format('#FFFFFF', '9'))
    worksheet.merge_range("E9:F9", "المؤهل" ,merge_format('#E0E0E0', '9'))
    worksheet.merge_range("G9:J9", "" ,merge_format('#FFFFFF', '9'))
    worksheet.merge_range("K9:L9", "التخصص" ,merge_format('#E0E0E0', '9'))
    worksheet.merge_range("M9:P9", "" ,merge_format('#FFFFFF', '9'))
    worksheet.merge_range("Q9:S9", "الفصل التدريبي" ,merge_format('#E0E0E0', '9'))
    worksheet.merge_range("T9:U9", "" ,merge_format('#FFFFFF', '9'))
    
    worksheet.write("A11", "المحاضرة",merge_format('#E0E0E0', '9'))
    worksheet.merge_range("A12:A13", "الوقت",merge_format('#E0E0E0', '9'))
    worksheet.merge_range("A14:A16", "الأحد",merge_format('#E0E0E0', '9'))
    worksheet.merge_range("A17:A19", "الإثنين",merge_format('#E0E0E0', '9'))
    worksheet.merge_range("A20:A22", "الثلاثاء",merge_format('#E0E0E0', '9'))
    worksheet.merge_range("A23:A25", "الأربعاء",merge_format('#E0E0E0', '9'))
    worksheet.merge_range("A26:A28", "الخميس",merge_format('#E0E0E0', '9'))
    
    worksheet.merge_range("V14:V28", "",colV)

            
    worksheet.merge_range("A29:U29", "",line29)
    
    worksheet.merge_range("A30:C30", "ساعات التدريب",merge_format('#E0E0E0', '9'))
    worksheet.merge_range("D30:G30", f"{sum(totalhours)}",merge_format('#FFFFFF', '9'))
    worksheet.merge_range("A31:C31", "ساعات مكتبية",merge_format('#E0E0E0', '9'))
    worksheet.merge_range("D31:G31", "",merge_format('#FFFFFF', '9'))
    worksheet.merge_range("A32:C32", "ساعات إدارية",merge_format('#E0E0E0', '9'))
    worksheet.merge_range("D32:G32", "",merge_format('#FFFFFF', '9'))
    worksheet.merge_range("A33:C33", "مجموع الساعات",merge_format('#E0E0E0', '9'))
    worksheet.merge_range("D33:G33", "",merge_format('#FFFFFF', '9'))
    
    row = 11
    t_start = ["08:00","09:00","10:00","11:00","12:00","13:00","14:00","15:00","16:00","17:00"]	
    t_end = ["08:40","09:40","10:40","11:40","12:40","13:40","14:40","15:40","16:40","17:40"]	
    


    for index, i in enumerate(TIME_CELLS_DICT, start =1 ):
        worksheet.merge_range(f"{TIME_CELLS_DICT[i][0]}{row}:{TIME_CELLS_DICT[i][1]}{row}", f'{index}',  merge_format('#E0E0E0', '9')) 


            
    row += 1
    for i, s, e in zip(TIME_CELLS_DICT, t_start, t_end ):
        worksheet.merge_range(f"{TIME_CELLS_DICT[i][0]}{row}:{TIME_CELLS_DICT[i][1]}{row}", f'{s}',  merge_format('#E0E0E0', '9'))
        worksheet.merge_range(f"{TIME_CELLS_DICT[i][0]}{row+1}:{TIME_CELLS_DICT[i][1]}{row+1}", f'{e}',  merge_format('#E0E0E0', '9'))

        
def run(file,department):
    global workbook, df
    df = pd.read_csv(file)
    LIST_OF_TEACHERS_ID = get_lab_department(department)
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})  
    for TEACHER_ID in LIST_OF_TEACHERS_ID:
        create_excel_file(teacherName(TEACHER_ID), TEACHER_ID)
    workbook.close()
    output.seek(0)
    return output


    
    
    
    
    
    
    
    
    
    
    
    