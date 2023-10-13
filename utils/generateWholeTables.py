import xlsxwriter, os
import pandas as pd
import io

parent_path = os.path.join(os.getcwd())
sheet_name = 'جدول المجمع (المدربين)'
excel_file = rf'{parent_path}\{sheet_name}.xlsx'

time_cells_dict = {"08":"4", "09":"5", "10":"6", "11":"7",
               "12":"8", "13":"9", "14":"10", "15":"11", "16":"12","17":"13"}

list_of_alphabets = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]

def teacherName(userID):
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

def split_without_strip(txt):
    res = [i.split('\n') for i in txt][0]
    return res

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

def removeNonValidString(*args):
    try:
        for a in args:
            a.remove("-")
    except ValueError:
        pass  # do nothing!
     
def day_column(s, e, d):
    start = int(s)
    end = int(e)
    if d == 'الأحد':
        return start, end
    elif d == "الإثنين":
        start += 10
        end += 10
        return start, end
    elif d == "الثلاثاء":
        start += 20
        end += 20
        return start, end
    elif d == "الأربعاء":
        start += 30
        end += 30
        return start, end
    elif d == "الخميس":
        start += 40
        end += 40
        return start, end
        
def get_subject_color(comID, subject):
    df_new = df[(df['اسم المقرر'] == (subject)) & (df['رقم المدرب'] == comID)]
    subject = df_new['القسم'].iloc[0]
    
    if subject == 'الدراسات العامة':
        return '#C5D9F1'
    if subject == 'الحاسب وتقنية المعلومات':
        return '#00B050'
    if subject == 'التقنية الالكترونية':
        return '#EF4360'
    if subject == 'أخرى':
        return '#A9A9A9'
        

# return current year term
def get_term(userID):
    df_new = df[df['رقم المدرب'] == (userID)]
    term = df_new['الفصل التدريبي'].iloc[0]
    
    return str(term)
        
# return a reformatted year term 
def get_term_text(term):
    if term[-2:-1] == '1': term = "الأول"
    elif term[-2:-1] == '2': term = "الثاني"
    elif term[-2:-1] == '3': term = "الثالث"
    
    return term
   
def merge_cells(timeslot, day, letter):
    x = timeslot.split("-")
    start_time   = x[1].strip()
    end_time     = x[0].strip()
    
    s = start_time[:2]
    e = end_time[:2]
    starting_cell = time_cells_dict[s]
    ending_cell   = time_cells_dict[e]

        
    start_column, end_column = day_column(starting_cell, ending_cell, day)
    
    hours = (end_column - start_column) + 1
    
    merge = f"{letter}{start_column}:{letter}{end_column}"
    
    return merge, str(hours) 

def ss01Details(userID):
    subjects, subject_reference, lab_id, days, times = [], [], [], [], []
    
    filtered_df  = df[df["اسم المدرب"].str.contains("-") == False]
    
    df_new = df[filtered_df['رقم المدرب'] == (userID)]
    subject_name = df_new['اسم المقرر'].to_string(index=False).strip()
    ref_subject_id = df_new['الرقم المرجعي'].to_string(index=False)
    laboratories = df_new['قاعة'].to_string(index=False).strip()
    lecture_times = df_new['الوقت'].to_string(index=False).strip()
    lecture_days = df_new['اليوم'].to_string(index=False).strip()
    
    
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

def create_excel_file():
     
    global merge_format
    def merge_format(back_color, size, font='black', border=1):
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


    worksheet.merge_range("A2:A3", "اليوم", merge_format("#808080", 12, font='white'))
    worksheet.merge_range("B2:C2", "أوقات المحاضرات", merge_format("#808080", 14, font='white'))
    worksheet.write("B3", "المحاضرة", merge_format("#808080", 12, font='white'))
    worksheet.write("C3", "الوقت", merge_format("#808080", 12, font='white'))

    lectures = ["الأولى","الثانية","الثالثة","الرابعة","الخامسة","السادسة","السابعة","الثامنة","التاسعة","العاشرة"]	
    times = ["8","9","10","11","12","1","2","3","4","5"]
    
    # set_lectures_times_columns = [woorksheet.write("B4")l , t for l , t in zip(lectures, times) ]
    col = 4
    for i in range(1, 5+1):
        for l, t in zip(lectures, times):
            worksheet.write(f"B{col}", f"{l}", (merge_format("#808080", 12, font='FF0000') if i%2==0 else merge_format("#D9D9D9", 12, font='FF0000')))
            worksheet.write(f"C{col}", f"{t}", (merge_format("#808080", 12, font='FF0000') if i%2==0 else merge_format("#D9D9D9", 12, font='FF0000')))
            col += 1
    
    worksheet.merge_range("A4:A13", "الأحد", merge_format("#D9D9D9", 12, font='FF0000'))
    worksheet.merge_range("A14:A23", "الإثنين", merge_format("#808080", 12, font='FF0000'))
    worksheet.merge_range("A24:A33", "الثلاثاء", merge_format("#D9D9D9", 12, font='FF0000'))
    worksheet.merge_range("A34:A43", "الأربعاء", merge_format("#808080", 12, font='FF0000'))
    worksheet.merge_range("A44:A53", "الخميس", merge_format("#D9D9D9", 12, font='FF0000'))
    
    
    worksheet.merge_range("A54:C54", "الساعات المعتمدة", merge_format("#808080", 12, font='white'))
    
    worksheet.merge_range("A55:C55", "التكاليف الداخلية", merge_format("#808080", 12, font='white'))
        
    
            
        # workbook.close()
        
def write(letter, last_letter, computer_id):
    global totalhours
    subs, refs, labs, times, days = ss01Details(computer_id)

    # remove_dash = self.removeNonValidString(subs, refs, labs, times, days)
    
    result = removeNonValidTimeSlot(times, subs, refs, labs, days)
    
    timeslots = []
    totalhours = []

    for time, day in zip(times, days):
        whatcell, traininghours = merge_cells(time, day, letter)
        timeslots.append(whatcell)
        totalhours.append(int(traininghours))
    
        
    
        
    worksheet.write(f"{letter}2", f"{teacherName(computer_id)}", merge_format("#636467", 12, font='white'))
    worksheet.write(f"{letter}3", f"{computer_id}", merge_format("#636467", 12, font='white'))
    
    for i in range(4, 55+1):
        worksheet.write(f"{letter}{i}:{letter}{i}", "", merge_format("#FFFFFF", 12))
    

    for sub, ref, lab, slot in zip(subs, refs, labs, timeslots): 

        subject_cell_color = get_subject_color(computer_id, f" {sub}")
        
        if slot[-3:] != slot[:3]:
            worksheet.merge_range(f"{slot}", f'{sub}\n{ref}\n{lab[-3:]}',  merge_format(subject_cell_color , 6))
        else:
            worksheet.write(f"{slot}", f'{sub}\n{ref}\n{lab[-3:]}',  merge_format(subject_cell_color, 6))
            
        worksheet.write(f"{letter}54:{letter}54", f"{sum(totalhours)}", merge_format("#D9D9D9", 6))
        
    
              
        
def run(file, department):

    global workbook, worksheet, df
    df = pd.read_csv(file)
    # Check if the number of columns is 20, otherwise raise an exception
    if len(df.columns) != 20:
        raise ValueError("Make sure you upload the correct file (SS01) from Rayat!")
        
    LIST_OF_TEACHERS_ID = get_lab_department(department)
    teacher_list_length = list_of_alphabets[:len(LIST_OF_TEACHERS_ID)]
    last_letter = teacher_list_length[-1]
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True}) 
    worksheet = workbook.add_worksheet('جدول المجمع (المدربين)') # add a new worksheet
    
    create_excel_file()
    # worksheet.merge_range(f"D5:{last_letter}55", "", merge_format(teachers, "#FFFFFF", 12))
    
    sumtotalhours = []
    for TEACHER_ID, first_letter in zip(LIST_OF_TEACHERS_ID, teacher_list_length):
        write(first_letter, last_letter, TEACHER_ID)
        sumtotalhours.append(sum(totalhours))
    
    
    term_in_number = get_term(TEACHER_ID)
    term_in_text = get_term_text(get_term(TEACHER_ID))
    worksheet.merge_range(f"A1:{last_letter}1", f"جدول المدربين المجمع  ( الفصل التدريبي {term_in_text} ) العام التدريبي {term_in_number}", merge_format("#FFFFFF", 22))
    worksheet.merge_range(f"A58:C58", "مجموع الساعات", merge_format("#808080", 12, font='white'))
    worksheet.write(f"D58", sum(sumtotalhours), merge_format("#FFFFFF", 12))
    
    worksheet.merge_range(f"A59:C59", "مقررات الحاسب الآلي", merge_format("#808080", 12, font='white'))
    worksheet.write(f"D59", "", merge_format("#00B050", 12))
    
    worksheet.merge_range(f"A60:C60", "مقررات الإلكترونيات", merge_format("#808080", 12, font='white'))
    worksheet.write(f"D60", "", merge_format("#EF4360", 12))
    
    worksheet.merge_range(f"A61:C61", "مقررات الدراسات العامة", merge_format("#808080", 12, font='white'))
    worksheet.write(f"D61", "", merge_format("#C5D9F1", 12))
    
    worksheet.set_column(f'A:{last_letter}', 12)    
    worksheet.set_row(0, 42)    
    
    
    workbook.close()
    output.seek(0)
    return output

    
    
    
                                                                                                                                              
                                                                                                                                              
        
        