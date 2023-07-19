import pandas as pd
import xlsxwriter
import streamlit as st
import io
from  utils.generateTables import run as generateTables
from utils.generateLabTables import run as generateLabTables
from utils.generateStudentTables import run as generateStudentTables
from utils.generateWholeTables import run as generateWholeTables
from utils.generateWholeLabTables import run as generateWholeLabTables

import json
import warnings

# Load the JSON file
with open('data\data.json', encoding='utf-8') as file:
    data = json.load(file)

css = """
<style>
    .rtl {
        direction: rtl;
        text-align: right;
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
            st.download_button(label=download_label, data=generateTables(uploaded_file, data['Departments'][0]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
    
        elif department == data['Departments'][1]:
            words = [data['Timetable_Type'][0], data['Departments'][1]]
            st.download_button(label=download_label, data=generateTables(uploaded_file, data['Departments'][1]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
    
        elif department == data['Departments'][2]:
            words = [data['Timetable_Type'][0], data['Departments'][2]]
            st.download_button(label=download_label, data=generateTables(uploaded_file, data['Departments'][2]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
    
        else:
            words = [data['Timetable_Type'][0], data['Departments'][3]]
            st.download_button(label=download_label, data=generateTables(uploaded_file, department='all'), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
            
    
        
    elif selected_type == data['Timetable_Type'][1]:
        st.markdown(f'<div class="rtl" style="font-size: 12pt; font-family: Cairo;">{department_text}</div>', unsafe_allow_html=True)
        department = st.selectbox("", [data['Departments'][0], data['Departments'][1], data['Departments'][3]])
        
        if department == data['Departments'][0]:
            words = [data['Timetable_Type'][1], data['Departments'][0]]
            st.download_button(label=download_label, data=generateLabTables(uploaded_file, data['Departments'][0]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
            
        elif department == data['Departments'][1]:
            words = [data['Timetable_Type'][1], data['Departments'][1]]
            st.download_button(label=download_label, data=generateLabTables(uploaded_file, data['Departments'][1]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
            
        else:
            words = [data['Timetable_Type'][1], data['Departments'][3]]
            st.download_button(label=download_label, data=generateLabTables(uploaded_file, department='all'), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
            
        
           
    elif selected_type == data['Timetable_Type'][2]:
        st.markdown(f'<div class="rtl" style="font-size: 12pt; font-family: Cairo;">{upload_2nd_file_text}</div>', unsafe_allow_html=True)
        uploaded_SF24_file  = st.file_uploader('', type='csv', key=2)
        if uploaded_SF24_file is not None:
            words = [data['Timetable_Type'][2]]
            st.download_button(label=download_label, data=generateStudentTables(uploaded_file, uploaded_SF24_file), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
            
        
    
    elif selected_type == data['Timetable_Type'][3]:
        st.markdown(f'<div class="rtl" style="font-size: 12pt; font-family: Cairo;">{department_text}</div>', unsafe_allow_html=True)
        department = st.selectbox("", data['Departments'])
        if department == data['Departments'][0]:
            words = [data['Timetable_Type'][3], data['Departments'][0]]
            st.download_button(label=download_label, data=generateWholeTables(uploaded_file, data['Departments'][0]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
    
        elif department == data['Departments'][1]:
            words = [data['Timetable_Type'][3], data['Departments'][1]]
            st.download_button(label=download_label, data=generateWholeTables(uploaded_file, data['Departments'][1]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
    
        elif department == data['Departments'][2]:
            words = [data['Timetable_Type'][3], data['Departments'][2]]
            st.download_button(label=download_label, data=generateWholeTables(uploaded_file, data['Departments'][2]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
    
        else:
            words = [data['Timetable_Type'][3], data['Departments'][3]]
            st.download_button(label=download_label, data=generateWholeTables(uploaded_file, department='all'), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
    
    
    elif selected_type == data['Timetable_Type'][4]:
        st.markdown(f'<div class="rtl" style="font-size: 12pt; font-family: Cairo;">{department_text}</div>', unsafe_allow_html=True)
        department = st.selectbox("", [data['Departments'][0], data['Departments'][1], data['Departments'][3]])
        if department == data['Departments'][0]:
            words = [data['Timetable_Type'][4], data['Departments'][0]]
            st.download_button(label=download_label, data=generateWholeLabTables(uploaded_file, data['Departments'][0]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
    
        elif department == data['Departments'][1]:
            words = [data['Timetable_Type'][4], data['Departments'][1]]
            st.download_button(label=download_label, data=generateWholeLabTables(uploaded_file, data['Departments'][1]), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
    
        else:
            words = [data['Timetable_Type'][4], data['Departments'][3]]
            st.download_button(label=download_label, data=generateWholeLabTables(uploaded_file, department='all'), file_name=f'{"_".join(words)}.xlsx', mime='application/vnd.ms-excel')
          
        
        










