import pandas as pd
import xlsxwriter
import streamlit as st
import io
import json
import warnings
from  utils.generateTables import run as generateTables
from utils.generateLabTables import run as generateLabTables
from utils.generateStudentTables import run as generateStudentTables
from utils.generateWholeTables import run as generateWholeTables
from utils.generateWholeLabTables import run as generateWholeLabTables
from utils.analyze_compliance import analyze_compliance
from data.courses import course_level_mapping_it, course_level_mapping_elec


# Import your utility functions
from utils import searchFreeSlot 

# Load the JSON file
with open('data/data.json', encoding='utf-8') as file:
    data = json.load(file)

css = """
<style>
    .rtl {
        direction: rtl;
        text-align: right;
        
        
</style>
"""

hide_github_icon = """
    <style>
    .css-1jc7ptx, .e1ewe7hr3, .viewerBadge_container__1QSob,
    .styles_viewerBadge__1yB5_, .viewerBadge_link__1S137,
    .viewerBadge_text__1JaDK {
        display: none;
    }
    </style>
    """

font_link = """
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@700;1000&display=swap" rel="stylesheet">
"""

st.markdown(font_link, unsafe_allow_html=True)
st.markdown(css, unsafe_allow_html=True)
st.markdown(hide_github_icon, unsafe_allow_html=True)


# Setup Sidebar
st.sidebar.title("")
choice = st.sidebar.selectbox("Services Menu", ["تحميل الجداول التدريبية", "استعلامات الجداول", "احصائية الانتظام بالخطط التدريبية"])

def timetables_download():
    title = "تحميل جداول الأقسام التدريبية (مدربين/قاعات/الجداول المجمعة)"
    st.markdown(f'<div class="rtl" style="font-size: 30pt; font-family: Cairo; margin-bottom: 100px; text-align: center;">{title}</div>', unsafe_allow_html=True)
    
    
    upload_text = "قم برفع ملف SS01 من تقارير رايات بتنسيق (CSV)"
    st.markdown(f'<div class="rtl" style="font-size: 12pt; font-family: Cairo;">{upload_text}</div>', unsafe_allow_html=True)
    uploaded_file  = st.file_uploader("", type='csv', key=1)
    download_label = "تنزيل الملف"
    
    if uploaded_file is not None:
        
    
        
        select_box_text = "اختر نوع الجدول المراد تنزيله على جهازك"
        st.markdown(f'<div class="rtl" style="font-size: 12pt; font-family: Cairo;">{select_box_text}</div>', unsafe_allow_html=True)
        selected_type = st.selectbox("", options=data['Timetable_Type'])
    
        department_text = "اختر القسم التدريبي"
        upload_2nd_file_text = "قم برفع ملف SF24 من تقارير رايات بتنسيق (CSV)"
        
        
        if selected_type == data['Timetable_Type'][0]:
            
            
            st.markdown(f'<div class="rtl" style="font-size: 12pt; font-family: Cairo;">{department_text}</div>', unsafe_allow_html=True)
            department = st.selectbox("", data['Departments'])
            if department == data['Departments'][0]:
                words = [data['Timetable_Type'][0], data['Departments'][0]]
                try:
                    st.download_button(label=download_label, data=generateTables(uploaded_file, data['Departments'][0]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
                except ValueError as e:
                    st.warning(e)
                    
            elif department == data['Departments'][1]:
                words = [data['Timetable_Type'][0], data['Departments'][1]]
                try:
                    st.download_button(label=download_label, data=generateTables(uploaded_file, data['Departments'][1]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
                except ValueError as e:
                    st.warning(e)
        
            elif department == data['Departments'][2]:
                words = [data['Timetable_Type'][0], data['Departments'][2]]
                try:
                    st.download_button(label=download_label, data=generateTables(uploaded_file, data['Departments'][2]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
                except ValueError as e:
                    st.warning(e)
                    
            else:
                words = [data['Timetable_Type'][0], data['Departments'][3]]
                try:
                    st.download_button(label=download_label, data=generateTables(uploaded_file, department='all'), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
                except ValueError as e:
                    st.warning(e)    
            
                
        elif selected_type == data['Timetable_Type'][1]:
            st.markdown(f'<div class="rtl" style="font-size: 12pt; font-family: Cairo;">{department_text}</div>', unsafe_allow_html=True)
            department = st.selectbox("", [data['Departments'][0], data['Departments'][1], data['Departments'][3]])
            
            if department == data['Departments'][0]:
                words = [data['Timetable_Type'][1], data['Departments'][0]]
                try:
                    st.download_button(label=download_label, data=generateLabTables(uploaded_file, data['Departments'][0]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
                except ValueError as e:
                    st.warning(e)
                    
            elif department == data['Departments'][1]:
                words = [data['Timetable_Type'][1], data['Departments'][1]]
                try:
                    st.download_button(label=download_label, data=generateLabTables(uploaded_file, data['Departments'][1]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
                except ValueError as e:
                    st.warning(e)
            else:
                words = [data['Timetable_Type'][1], data['Departments'][3]]
                try:
                    st.download_button(label=download_label, data=generateLabTables(uploaded_file, department='all'), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
                except ValueError as e:
                    st.warning(e)
                
                   
        elif selected_type == data['Timetable_Type'][2]:
            st.markdown(f'<div class="rtl" style="font-size: 12pt; font-family: Cairo;">{upload_2nd_file_text}</div>', unsafe_allow_html=True)
            uploaded_SF24_file  = st.file_uploader('', type='csv', key=2)
            if uploaded_SF24_file is not None:
                st.markdown(f'<div class="rtl" style="font-size: 12pt; font-family: Cairo;">{department_text}</div>', unsafe_allow_html=True)
                department = st.selectbox("", [data['Departments'][0], data['Departments'][1], data['Departments'][3]])
                
                if department == data['Departments'][0]:
                    words = [data['Timetable_Type'][2], data['Departments'][0]]
                    try:
                        st.download_button(label=download_label, data=generateStudentTables(uploaded_file, uploaded_SF24_file, data['Departments'][0]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
                    except ValueError as e:
                        st.warning(e)
                        
                elif department == data['Departments'][1]:
                    words = [data['Timetable_Type'][2], data['Departments'][1]]
                    try:
                        st.download_button(label=download_label, data=generateStudentTables(uploaded_file, uploaded_SF24_file, data['Departments'][1]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
                    except ValueError as e:
                        st.warning(e)
                        
                else:
                    words = [data['Timetable_Type'][2], data['Departments'][3]]
                    try:
                        st.download_button(label=download_label, data=generateStudentTables(uploaded_file, uploaded_SF24_file, department='all'), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
                    except ValueError as e:
                        st.warning(e)
                
            
        elif selected_type == data['Timetable_Type'][3]:
            st.markdown(f'<div class="rtl" style="font-size: 12pt; font-family: Cairo;">{department_text}</div>', unsafe_allow_html=True)
            department = st.selectbox("", data['Departments'])
            if department == data['Departments'][0]:
                words = [data['Timetable_Type'][3], data['Departments'][0]]
                try:
                    st.download_button(label=download_label, data=generateWholeTables(uploaded_file, data['Departments'][0]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
                except ValueError as e:
                    st.warning(e)
                    
            elif department == data['Departments'][1]:
                words = [data['Timetable_Type'][3], data['Departments'][1]]
                try:
                    st.download_button(label=download_label, data=generateWholeTables(uploaded_file, data['Departments'][1]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
                except ValueError as e:
                    st.warning(e)
                    
            elif department == data['Departments'][2]:
                words = [data['Timetable_Type'][3], data['Departments'][2]]
                try:
                    st.download_button(label=download_label, data=generateWholeTables(uploaded_file, data['Departments'][2]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
                except ValueError as e:
                    st.warning(e)
                    
            else:
                words = [data['Timetable_Type'][3], data['Departments'][3]]
                try:
                    st.download_button(label=download_label, data=generateWholeTables(uploaded_file, department='all'), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
                except ValueError as e:
                    st.warning(e)
                
        elif selected_type == data['Timetable_Type'][4]:
            st.markdown(f'<div class="rtl" style="font-size: 12pt; font-family: Cairo;">{department_text}</div>', unsafe_allow_html=True)
            department = st.selectbox("", [data['Departments'][0], data['Departments'][1], data['Departments'][3]])
            if department == data['Departments'][0]:
                words = [data['Timetable_Type'][4], data['Departments'][0]]
                try:
                    st.download_button(label=download_label, data=generateWholeLabTables(uploaded_file, data['Departments'][0]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
                except ValueError as e:
                    st.warning(e)
                    
            elif department == data['Departments'][1]:
                words = [data['Timetable_Type'][4], data['Departments'][1]]
                try:
                    st.download_button(label=download_label, data=generateWholeLabTables(uploaded_file, data['Departments'][1]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
                except ValueError as e:
                    st.warning(e)
                    
            else:
                words = [data['Timetable_Type'][4], data['Departments'][3]]
                try:
                    st.download_button(label=download_label, data=generateWholeLabTables(uploaded_file, department='all'), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
                except ValueError as e:
                    st.warning(e)  
            
    
def find_slot():
    title = "البحث عن أوقات فراغ لشعبة محددة"
    st.markdown(f'<div class="rtl" style="font-size: 30pt; font-family: Cairo; margin-bottom: 100px; text-align: center;">{title}</div>', unsafe_allow_html=True)
    upload_text = "قم برفع ملف SS05 من تقارير رايات بتنسيق (CSV)"
    st.markdown(f'<div class="rtl" style="font-size: 12pt; font-family: Cairo;">{upload_text}</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("", type='csv', key=2)
    crn_input = st.text_input("Enter CRN (Course Reference Number):")
    
    if uploaded_file and crn_input:
        try:
            crn = int(crn_input)
            schedules = searchFreeSlot.get_students_schedule(uploaded_file, crn)
            if schedules:
                free_slots = searchFreeSlot.find_common_free_slots(schedules)
                st.write(free_slots)
            else:
                st.error("Invalid CRN or no data available for the provided CRN.")
        except ValueError:
            st.error("Please enter a valid CRN.")

def analyze_statistical_compliance():
    title = "احصائية الانتظام بالخطط التدريبية"
    st.markdown(f'<div class="rtl" style="font-size: 30pt; font-family: Cairo; margin-bottom: 100px; text-align: center;">{title}</div>', unsafe_allow_html=True)
    
    # File upload input
    upload_text = "قم برفع ملف SS05 بتنسيق (CSV)"
    st.markdown(f'<div class="rtl" style="font-size: 12pt; font-family: Cairo;">{upload_text}</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("", type='csv', key="compliance")
    
    # Dropdown for department selection
    department = st.selectbox("اختر القسم:", ["الحاسب وتقنية المعلومات", "التقنية الالكترونية"])
    
    # User inputs the current term
    current_term = st.text_input("ادخل الفصل التدريبي الحالي (مثال: 144520):", key="current_term")
    
    
    
    if uploaded_file and current_term and department:
        try:

            
            # Select course mapping based on department
            course_mapping = course_level_mapping_it if department == "الحاسب وتقنية المعلومات" else course_level_mapping_elec
            
            # Run analysis with the selected department's course mapping
            result_table = analyze_compliance(uploaded_file, current_term, course_mapping)
            st.write("جدول الاحصائيات:")
            st.dataframe(result_table)
            
            # Provide download option
            download_label = "تحميل جدول الاحصائية"
            file_name = "compliance_statistics.xlsx"
            with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
                result_table.to_excel(writer, index=False, sheet_name='Statistics')
            st.download_button(label=download_label, data=open(file_name, "rb").read(), file_name=file_name, mime='application/vnd.ms-excel')
        except Exception as e:
            st.warning(f"⚠️ الملف الذي رفعته قد يكون خاص بقسم آخر . الرجاء رفع الملف الصحيح")

# Main app logic
if choice == "تحميل الجداول التدريبية":
    timetables_download()
elif choice == "استعلامات الجداول":
    find_slot()
elif choice == "احصائية الانتظام بالخطط التدريبية":
    analyze_statistical_compliance()
                    
            










