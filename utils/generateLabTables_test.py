import xlsxwriter, os
import pandas as pd

class Generate_Table:
    
    global time_cells_dict, day_cells_dict , rayat_file, excel_file, parent_path
    
    parent_path = os.path.join(os.getcwd())
    rayat_file = rf'{parent_path}\SS01.csv'
    sheet_name = 'جداول القاعات'
    excel_file = rf'{parent_path}\{sheet_name}.xlsx'
    time_cells_dict = {"08":"BC", "09":"DE", "10":"FG", "11":"HI",
                       "12":"JK", "13":"LM", "14":"NO", "15":"PQ", "16":"RS","17":"TU"}
    day_cells_dict = {"الأحد":"14","الإثنين":"17","الثلاثاء":"20","الاربعاء":"23","الخميس":"26"}
    
    
    def __init__(self, lab_id):
        self.lab_id = lab_id
    
    def __str__(self):
        pass

    def get_unique_labIDs(self):
        df = pd.read_csv(rayat_file)
        IDs = list(set( df['قاعة'] ))
        # cleaned = [x for x in IDs if x.isdigit()]
        return IDs
        
    def split(self, txt):
        res = [i.split('\n') for i in txt][0]
        stripped = list(map(str.strip, res))

        return stripped

    # return current year term
    def get_term(self, labID):
        
        df = pd.read_csv(rayat_file)
        df_new = df[df['قاعة'] == (int(labID))]
        term = df_new['الفصل التدريبي'].iloc[0]
        
        return str(term)
    
    # return a reformatted year term 
    def get_term_text(self, term):
        if term[-2:-1] == '1': term = "الأول"
        elif term[-2:-1] == '2': term = "الثاني"
        elif term[-2:-1] == '3': term = "الثالث"
        
        return term

    def removeNonValidTimeSlot(self, timeslots, *arguments):
        Error = '18'
        for i in range(len(timeslots)-1, -1, -1):
            if not Error in timeslots[i]: 
                pass
            else:
                for x in arguments:
                    x.pop(i)
                timeslots.pop(i)
        return timeslots

    def merge_cells(self, timeslot, day):
        x = timeslot.split("-")
        start_time   = x[1].strip()
        end_time     = x[0].strip()
        
        s = start_time[:2]
        e = end_time[:2]
        starting_cell = time_cells_dict[s][0]
        ending_cell   = time_cells_dict[e][-1]
        day_row = day_cells_dict[day]
        
        start_and_end = starting_cell + ending_cell
        merge = f"{starting_cell}{day_row}:{ending_cell}{int(day_row)+2}"
        
        return merge, start_and_end
  
    def ss01Details(self, labID):
        
        df = pd.read_csv(rayat_file)
        subjects, subject_reference, teachernames, days, times = [], [], [], [], []
        
        df_new = df[df['قاعة'] == (int(labID))]
        subject_name = df_new['اسم المقرر'].to_string(index=False).strip()
        ref_subject_id = df_new['الرقم المرجعي'].to_string(index=False).strip()
        teachers = df_new['اسم المدرب'].to_string(index=False).strip()
        lecture_times = df_new['أوقات'].to_string(index=False).strip()
        lecture_days = df_new['أيام'].to_string(index=False).strip()
        
        
        subjects.append(subject_name)
        subjects = self.split(subjects)
        
        subject_reference.append(ref_subject_id)
        subject_reference =self.split(subject_reference)
        
        teachernames.append(teachers)
        teachernames = self.split(teachernames)
        
        times.append(lecture_times)
        times = self.split(times)
        
        days.append(lecture_days)
        days = self.split(days)
        
        return subjects, subject_reference, teachernames, times, days


    def xlsx(self, lab):

        worksheet = workbook.add_worksheet(lab) # add a new worksheet
        subs, refs, teachernames, times, days = self.ss01Details(self.lab_id)
        result = self.removeNonValidTimeSlot(times, subs, refs, teachernames, days)
        print(lab, self.lab_id)
        print(type(lab), type(self.lab_id))
        def total_hours(s, e):
            sum = ord(e) - ord(s) + 1
            sum //= 2 
            return sum
        
        def merge_format(self, back_color, size, font='black'):
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
        
        def no_Border(self, back_color, size, font='black'):
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
        
        timeslots = []
        totalhours = []
        for t, d in zip(times, days):
            whatcell, traininghours = self.merge_cells(t, d)
            timeslots.append(whatcell)
            totalhours.append(total_hours(traininghours[0], traininghours[1]))


        for sub, ref, teacher, slot in zip(subs, refs, teachernames, timeslots):
            worksheet.merge_range(f"{slot}", f'{sub}\n{ref}\n{teacher}',  merge_format(self, '#E0E0E0', '9'))

        
        
        colV = workbook.add_format()
        colV.set_left(6)
        
        line29 = workbook.add_format()
        line29.set_top(6)
        
        
        # Set column size
        # worksheet.set_column('A:A', 5.71) 
        worksheet.set_column('A:U', 6.29) 

        # Insert image (tvtc)
        worksheet.insert_image('G1', rf'{parent_path}\icon\tvtc.jpg', {'x_scale': 0.3, 'y_scale': 0.3, 'x_offset': 100,'y_offset': 0})
        
    
        worksheet.merge_range("A1:F1", "المملكة العربية السعودية" , no_Border(self, '#FFFFFF', 12))
        worksheet.merge_range("A2:F2", "المؤسسة العامة للتدريب التقني والمهني" , no_Border(self, '#FFFFFF', 12))
        worksheet.merge_range("A3:F3", "الكلية التقنية بمحافظة حقل" , no_Border(self, '#FFFFFF', 12))

        worksheet.merge_range("M1:U1", "Kingdom of Saudi Arabia" ,no_Border(self, '#FFFFFF', 12))
        worksheet.merge_range("M2:U2", "Technical and Vocational Training Corporation" ,no_Border(self, '#FFFFFF', 12))
        worksheet.merge_range("M3:U3", "College of Technology in Haql" ,no_Border(self, '#FFFFFF', 12))
        
        worksheet.merge_range("A5:U7", f"الجدول التدريبي قاعة ({str(lab)[-3:]})" ,merge_format(self, '#FFFFFF', 14, font='black'))
        
        
        worksheet.write("A11", "المحاضرة",merge_format(self, '#E0E0E0', '9'))
        worksheet.merge_range("A12:A13", "الوقت",merge_format(self, '#E0E0E0', '9'))
        worksheet.merge_range("A14:A16", "الأحد",merge_format(self, '#E0E0E0', '9'))
        worksheet.merge_range("A17:A19", "الإثنين",merge_format(self, '#E0E0E0', '9'))
        worksheet.merge_range("A20:A22", "الثلاثاء",merge_format(self, '#E0E0E0', '9'))
        worksheet.merge_range("A23:A25", "الأربعاء",merge_format(self, '#E0E0E0', '9'))
        worksheet.merge_range("A26:A28", "الخميس",merge_format(self, '#E0E0E0', '9'))
        
        worksheet.merge_range("V14:V28", "",colV)
        worksheet.merge_range("A29:U29", "",line29)
        
        worksheet.merge_range("A30:C31", "ساعات التدريب",merge_format(self, '#E0E0E0', '12'))
        worksheet.merge_range("D30:G31", f"{sum(totalhours)}",merge_format(self, '#FFFFFF', '12'))

        
        row = 11
        t_start = ["08:00","09:00","10:00","11:00","12:00","13:00","14:00","15:00","16:00","17:00"]	
        t_end = ["08:50","09:50","10:50","11:40","12:45","13:40","14:40","15:40","16:40","17:50"]	

        for index, i in enumerate(time_cells_dict, start =1 ):
            worksheet.merge_range(f"{time_cells_dict[i][0]}{row}:{time_cells_dict[i][1]}{row}", f'{index}',  merge_format(self, '#E0E0E0', '9')) 
        
        row += 1
        for i, s, e in zip(time_cells_dict, t_start, t_end ):
            worksheet.merge_range(f"{time_cells_dict[i][0]}{row}:{time_cells_dict[i][1]}{row}", f'{s}',  merge_format(self, '#E0E0E0', '9'))
            worksheet.merge_range(f"{time_cells_dict[i][0]}{row+1}:{time_cells_dict[i][1]}{row+1}", f'{e}',  merge_format(self, '#E0E0E0', '9'))
        
        
def main():


    lab = Generate_Table("")
    t_list = lab.get_unique_labIDs()
    
    string_labs = list(map(str, t_list))
    global workbook
    workbook = xlsxwriter.Workbook(excel_file) 
    
    for t in string_labs:
        lab = Generate_Table(str(t))
        lab.xlsx(str(t))
    
    workbook.close()

    
if __name__ == "__main__":
    main()
    
    
    
    
    
    
    
    
    
    
    
    
    