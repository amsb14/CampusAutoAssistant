import xlsxwriter, os
import pandas as pd
import io
global time_cells_dict, day_cells_dict , rayat_file, excel_file, parent_path
parent_path = os.path.join(os.getcwd())
rayat_file = rf'{parent_path}\SS01.csv'
sheet_name = 'جداول القاعات'
excel_file = rf'{parent_path}\{sheet_name}.xlsx'
time_cells_dict = {"08": "BC", "09": "DE", "10": "FG", "11": "HI",
                   "12": "JK", "13": "LM", "14": "NO", "15": "PQ", "16": "RS", "17": "TU"}
day_cells_dict = {"الأحد": "14", "الاثنين": "17", "الثلاثاء": "20", "الأربعاء": "23", "الخميس": "26"}



def get_unique_labIDs(department='all'):
    IDs = list(set(df['قاعة']))
    return IDs

def get_lab_department(department='all'):
    if department != "all":
        IDs = df.loc[df['القسم'] == department, 'قاعة'].unique().tolist()
    else:
        IDs = df['قاعة'].unique().tolist()
    return IDs  

def split(txt):
    res = [i.split('\n') for i in txt][0]
    stripped = list(map(str.strip, res))
    return stripped


def get_term(labID):
    df_new = df[df['قاعة'] == int(labID)]
    term = df_new['الفصل التدريبي'].iloc[0]
    return str(term)


def get_term_text(term):
    if term[-2:-1] == '1': return "الأول"
    if term[-2:-1] == '2': return "الثاني"
    if term[-2:-1] == '3': return "الثالث"


def removeNonValidTimeSlot(timeslots, *arguments):
    Error = '18'
    for i in range(len(timeslots) - 1, -1, -1):
        if not Error in timeslots[i]:
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
    merge = f"{starting_cell}{day_row}:{ending_cell}{int(day_row)+2}"

    return merge, start_and_end


def ss01Details(labID):
    subjects, subject_reference, teachernames, days, times = [], [], [], [], []

    df_new = df[df['قاعة'] == int(labID)]
    subject_name = df_new['اسم المقرر'].to_string(index=False).strip()
    ref_subject_id = df_new['الرقم المرجعي'].to_string(index=False).strip()
    teachers = df_new['اسم المدرب'].to_string(index=False).strip()
    lecture_times = df_new['الوقت'].to_string(index=False).strip()
    lecture_days = df_new['اليوم'].to_string(index=False).strip()

    subjects.append(subject_name)
    subjects = split(subjects)

    subject_reference.append(ref_subject_id)
    subject_reference = split(subject_reference)

    teachernames.append(teachers)
    teachernames = split(teachernames)

    times.append(lecture_times)
    times = split(times)

    days.append(lecture_days)
    days = split(days)

    return subjects, subject_reference, teachernames, times, days

def create_excel_file(lab):

    worksheet = workbook.add_worksheet(lab) # add a new worksheet
    subs, refs, teachernames, times, days =  ss01Details(lab)
    result =  removeNonValidTimeSlot(times, subs, refs, teachernames, days)

    def total_hours(s, e):
        sum = ord(e) - ord(s) + 1
        sum //= 2 
        return sum
    
    def merge_format( back_color, size, font='black'):
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
    
    def no_Border( back_color, size, font='black'):
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
        whatcell, traininghours =  merge_cells(t, d)
        timeslots.append(whatcell)
        totalhours.append(total_hours(traininghours[0], traininghours[1]))

    list_of_alphabets = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U"]
    for letter in list_of_alphabets:
        for i in range(14,28+1):
            worksheet.write(f"{letter}{i}:{letter}{i}", "", merge_format2("#FFFFFF", 9))

    for sub, ref, teacher, slot in zip(subs, refs, teachernames, timeslots):
        worksheet.merge_range(f"{slot}", f'{sub}\n{ref}\n{teacher}',  merge_format( '#E0E0E0', '9'))

    
    
    colV = workbook.add_format()
    colV.set_left(6)
    
    line29 = workbook.add_format()
    line29.set_top(6)
    
    
    # Set column size
    # worksheet.set_column('A:A', 5.71) 
    worksheet.set_column('A:U', 6.29) 

    # Insert image (tvtc)
    worksheet.insert_image('G1', 'icon/tvtc.jpg', {'x_scale': 0.3, 'y_scale': 0.3, 'x_offset': 100,'y_offset': 0})
    

    worksheet.merge_range("A1:F1", "المملكة العربية السعودية" , no_Border( '#FFFFFF', 12))
    worksheet.merge_range("A2:F2", "المؤسسة العامة للتدريب التقني والمهني" , no_Border( '#FFFFFF', 12))
    worksheet.merge_range("A3:F3", "الكلية التقنية بمحافظة حقل" , no_Border( '#FFFFFF', 12))

    worksheet.merge_range("M1:U1", "Kingdom of Saudi Arabia" ,no_Border( '#FFFFFF', 12))
    worksheet.merge_range("M2:U2", "Technical and Vocational Training Corporation" ,no_Border( '#FFFFFF', 12))
    worksheet.merge_range("M3:U3", "College of Technology in Haql" ,no_Border( '#FFFFFF', 12))
    
    worksheet.merge_range("A5:U7", f"الجدول التدريبي قاعة ({str(lab)[-3:]})" ,merge_format( '#FFFFFF', 14, font='black'))
    
    
    worksheet.write("A11", "المحاضرة",merge_format( '#E0E0E0', '9'))
    worksheet.merge_range("A12:A13", "الوقت",merge_format( '#E0E0E0', '9'))
    worksheet.merge_range("A14:A16", "الأحد",merge_format( '#E0E0E0', '9'))
    worksheet.merge_range("A17:A19", "الإثنين",merge_format( '#E0E0E0', '9'))
    worksheet.merge_range("A20:A22", "الثلاثاء",merge_format( '#E0E0E0', '9'))
    worksheet.merge_range("A23:A25", "الأربعاء",merge_format( '#E0E0E0', '9'))
    worksheet.merge_range("A26:A28", "الخميس",merge_format( '#E0E0E0', '9'))
    
    worksheet.merge_range("V14:V28", "",colV)
    worksheet.merge_range("A29:U29", "",line29)
    
    worksheet.merge_range("A30:C31", "ساعات التدريب",merge_format( '#E0E0E0', '12'))
    worksheet.merge_range("D30:G31", f"{sum(totalhours)}",merge_format( '#FFFFFF', '12'))

    
    row = 11
    t_start = ["08:00","09:00","10:00","11:00","12:00","13:00","14:00","15:00","16:00","17:00"]	
    t_end = ["08:50","09:50","10:50","11:40","12:45","13:40","14:40","15:40","16:40","17:50"]	

    for index, i in enumerate(time_cells_dict, start =1 ):
        worksheet.merge_range(f"{time_cells_dict[i][0]}{row}:{time_cells_dict[i][1]}{row}", f'{index}',  merge_format( '#E0E0E0', '9')) 
    
    row += 1
    for i, s, e in zip(time_cells_dict, t_start, t_end ):
        worksheet.merge_range(f"{time_cells_dict[i][0]}{row}:{time_cells_dict[i][1]}{row}", f'{s}',  merge_format( '#E0E0E0', '9'))
        worksheet.merge_range(f"{time_cells_dict[i][0]}{row+1}:{time_cells_dict[i][1]}{row+1}", f'{e}',  merge_format( '#E0E0E0', '9'))
    
        
def run(file,department):
    global workbook, df
    df = pd.read_csv(file)
    # Check if the number of columns is 24, otherwise raise an exception
    if len(df.columns) != 24:
        raise ValueError("Make sure you upload the correct file (SS01) from Rayat!")
    LIST_OF_LABS_ID = get_lab_department(department)    
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})  
    for LABS_ID in LIST_OF_LABS_ID:
        create_excel_file(str(LABS_ID))
    workbook.close()
    output.seek(0)
    return output
    
    

    
    
    
    
    
    
    
    
    
    
    
    