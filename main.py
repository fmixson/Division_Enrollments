import openpyxl
from openpyxl import load_workbook
import lxml
from configparser import ConfigParser
import time
import pandas as pd
# import openpyxl
import os, sys



class DataframeWork:

    def __init__(self, enrollment_df):
        self.enrollment_df = enrollment_df



    # def sheet_integers(self):
    #     self.enrollment_df['Size'] = pd.to_numeric(self.enrollment_df['Size'], errors='coerce').fillna(0).astype('int')
    #     self.enrollment_df['Max'] = pd.to_numeric(self.enrollment_df['Max'], errors='coerce').fillna(0).astype('int')
    #     self.enrollment_df['Hours'] = pd.to_numeric(self.enrollment_df['Hours'], errors='coerce').fillna(0).astype(
    #         'int')
    #     return self.enrollment_df


    def lecture_only(self):
        science_labs = ['A&P 120', 'A&P 150', 'ASTR 105L', 'BIOL 120', 'CHEM 100', 'CHEM 110', 'CHEM 111', 'CHEM 112',
                        'CHEM 211',
                        'CHEM 212', 'CHEM 250l', 'ESCI 104L', 'GEOG 101L', 'GEOL 102L', 'MICR 200']
        # lecture_df = self.enrollment_df[self.enrollment_df['Type'] == 'Lecture'].reset_index()
        for i in range(len(self.enrollment_df)):
            if self.enrollment_df.loc[i, 'Course'] in science_labs:
                if self.enrollment_df.loc[i, 'Component'] == 'Laboratory':
                    self.enrollment_df.loc[i, 'Component'] = 'Sci Lab'
                if self.enrollment_df.loc[i, 'Component'] == 'Lecture':
                    self.enrollment_df.loc[i, 'Component'] = 'Sci-Lec'
        #         switch laboratory to Sci Lab
        #         switch lecture to laboratory
        # self.enrollment_df = self.enrollment_df[self.enrollment_df['Component'] != 'Laboratory']
        self.enrollment_df = self.enrollment_df.fillna(0)
        self.enrollment_df = self.enrollment_df[self.enrollment_df['Component'] != 0]
        cancelled_df = self.enrollment_df[self.enrollment_df['Status'] == 'Cancelled']
        self.enrollment_df = self.enrollment_df[self.enrollment_df['Status'] == 'Active']
        self.enrollment_df = self.enrollment_df.reset_index()
        self.enrollment_df.to_excel('Enrollment.xlsx')
        cancelled_df.to_excel('Cancelled_Test.xlsx')
        # lecture_enrollment_df = lecture_df

    def modalities(self):

        for i in range(len(self.enrollment_df)):
            print('i = ', i)
            if self.enrollment_df.loc[i, 'Room'] == 'ONLINE':
                self.enrollment_df.loc[i, 'Modality'] = 'ONLINE'
            elif self.enrollment_df.loc[i, 'Room'] == 'REMOTE':
                self.enrollment_df.loc[i, 'Modality'] = 'REMOTE'
            else:
                self.enrollment_df.loc[i, 'Modality'] = 'IN PERSON'
        for i in range(len(self.enrollment_df)):
            if self.enrollment_df.loc[i, 'Comment'] == '(HYBRID)':
                self.enrollment_df.loc[i, 'Modality'] = 'HYBRID'
            if self.enrollment_df.loc[i, 'Modality'] == '(HYFLEX)':
                self.enrollment_df.loc[i, 'Modality'] = 'HYFLEX'



        #unique combined courses
        #split into separate courses
        #create mini data frame of the combined courses
        #use the combined courses language to identify the course
        #combine section numbers
        #sum the total enrollment
        #create data frame that has only zeroes in combined column
        #
        # insert column into data frame
        # lecture_enrollment_df['Modality2']=lecture_enrollment_df.loc[lecture_enrollment_df['Modality'] == 'Hybrid Course', 'Modality2'] = 'HYBRID'
        # lecture_enrollment_df['Modality2'] = lecture_enrollment_df.loc[lecture_enrollment_df['Modality'] == 'Hyflex Course', 'Modality2'] = 'HYFLEX'
        # lecture_enrollment_df.loc[lecture_enrollment_df['Room'] == 'ONLINE', 'Modality2'] = 'Online'
        # lecture_enrollment_df.loc[lecture_enrollment_df['Room'] == 'REMOTE', 'Modality2'] = 'Remote'
        # lecture_enrollment_df.loc[lecture_enrollment_df['Modality'] == '(Honors Section) Hybrid Course', 'Modality2'] = 'Hybrid'
        #
        # rooms = ['AHS *','WHS *', 'LA105', 'LA106', 'LA109', 'LA201', 'SS207', 'LA213', 'LC218', 'LA110', 'SS211', 'SS225', 'LA103', 'SS224',
        #           'LC217','LA211', 'LA202', 'LA209', 'LA210', 'LA205', 'LA212', 'LA204', 'SS214', 'LA212', 'SS136', 'LC213',
        #          'LA203', 'FA134', 'LM20*', 'FA133', 'SS136', 'BELF*', 'MAYF*', 'AHS *', 'LC134', 'SPSM*', 'WHS *', 'NHS*',
        #          'STPI*', 'DOWN*', 'LA104', 'MP209', 'SS137', 'SS212', 'SS213', 'WARR*', 'AHS*', 'WHS*']
        # for room in rooms:
        #         lecture_enrollment_df.loc[lecture_enrollment_df['Room'] == room, 'Modality2'] = 'In Person'
        # lecture_enrollment_df.loc[lecture_enrollment_df['Modality'] == 'Hybrid Course', 'Modality2'] = 'Hybrid'

    def calculations(self):
        # self.enrollment_df['Fill Rate'] = self.enrollment_df['Size'] / self.enrollment_df['Max']
        # for i in range(len(self.enrollment_df)):
        #     if self.enrollment_df.loc[i, 'Session'] == 18:
        #         self.enrollment_df['Length Multiplier'] = 17.5
        #     else:
        #         print(len(self.enrollment_df.loc[i, 'Days']))
        #         self.enrollment_df.loc[i, 'Length Multiplier'] = len(self.enrollment_df.loc[i, 'Days'])

        # self.enrollment_df['FTES'] = self.enrollment_df['Size'] * (
        #         (self.enrollment_df['Hours'] / 18) / 525)
        self.enrollment_df['Potential FTES'] = self.enrollment_df['Capacity'] * (
                    (self.enrollment_df['Hours'] / 18) / 525)
        self.enrollment_df['FTEF'] = (self.enrollment_df['Hours'] / 18) / 15
        self.enrollment_df['Efficiency'] = self.enrollment_df['FTES'] / self.enrollment_df['FTEF']
        self.enrollment_df['Potential Efficiency'] = self.enrollment_df['Potential FTES'] / self.enrollment_df['FTEF']
        self.enrollment_df.reset_index()

    def clean_up_sessions(self):
        print('lecture df', self.enrollment_df)
        print('datatype', type(self.enrollment_df['Session']))
        for i in range(len(self.enrollment_df)):

            if 'Regular' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '18'
            if 'Apprenticeship' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '18'
            elif self.enrollment_df.loc[i, 'Session'] == 'Fifteen Week A ':
                self.enrollment_df.loc[i, 'Session'] = '15A'
            elif self.enrollment_df.loc[i, 'Session'] == 'Fifteen Week A1 ':
                self.enrollment_df.loc[i, 'Session'] = '15A'
            elif self.enrollment_df.loc[i, 'Session'] == 'Fifteen Week A2 ':
                self.enrollment_df.loc[i, 'Session'] = '15A'
            elif self.enrollment_df.loc[i, 'Session'] == 'Fifteen Week A3 ':
                self.enrollment_df.loc[i, 'Session'] = '15A'
            elif self.enrollment_df.loc[i, 'Session'] == 'Fifteen Week A6 ':
                self.enrollment_df.loc[i, 'Session'] = '15A'
            elif self.enrollment_df.loc[i, 'Session'] == 'Fifteen Week A7 ':
                self.enrollment_df.loc[i, 'Session'] = '15A'
            elif self.enrollment_df.loc[i, 'Session'] == 'Fifteen Week T, W, F Session':
                self.enrollment_df.loc[i, 'Session'] = '15A'

            elif 'Fifteen Week A M-F' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '15A'
            elif 'Fifteen Week A2 Tuesday' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '15A'
            elif 'Fifteen Week A2 ' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '15A'
            elif 'Fifteen Week A3 Wednesday' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '15A'
            elif 'Fifteen Week A6 M,W' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '15A'
            elif 'Fifteen Week A7 T,Th' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '15A'
            elif 'Fifteen Week A1 Monday' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '15A'


            elif 'Fifteen Week B M-F' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '15B'
            elif 'Fifteen Week B1 Monday' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '15B'
            elif 'Fifteen Week B6 M,W' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '15B'
            elif self.enrollment_df.loc[i, 'Session'] == 'Fifteen Week B7 T,Th':
                self.enrollment_df.loc[i, 'Session'] = '15B'
            elif self.enrollment_df.loc[i, 'Session'] == 'Fifteen Week B8 M-Th':
                self.enrollment_df.loc[i, 'Session'] = '15B'
            elif self.enrollment_df.loc[i, 'Session'] == 'Fifteen Week B3 Wednesday':
                self.enrollment_df.loc[i, 'Session'] = '15B'
            elif self.enrollment_df.loc[i, 'Session'] == 'Fifteen Week B4 Thursday':
                self.enrollment_df.loc[i, 'Session'] = '15B'
            elif self.enrollment_df.loc[i, 'Session'] == 'Fifteen Week B5 Friday':
                self.enrollment_df.loc[i, 'Session'] = '15B'
            elif self.enrollment_df.loc[i, 'Session'] == 'Fifteen Week B2 Tuesday':
                self.enrollment_df.loc[i, 'Session'] = '15B'
            elif self.enrollment_df.loc[i, 'Session'] == 'Fifteen Week B12 T,Th, F':
                self.enrollment_df.loc[i, 'Session'] = '15B'
            elif self.enrollment_df.loc[i, 'Session'] == 'Fifteen Week B13 T,W,Th':
                self.enrollment_df.loc[i, 'Session'] = '15B'
            elif self.enrollment_df.loc[i, 'Session'] == 'Fifteen Week B16 M,T,W':
                self.enrollment_df.loc[i, 'Session'] = '15B'
            elif self.enrollment_df.loc[i, 'Session'] == 'Fifteen Week B20':
                self.enrollment_df.loc[i, 'Session'] = '15B'
            elif self.enrollment_df.loc[i, 'Session'] == '15 week E15 T, W, Th':
                self.enrollment_df.loc[i, 'Session'] = '15B'
            elif self.enrollment_df.loc[i, 'Session'] == 'Fifteen Week 151 Th, F':
                self.enrollment_df.loc[i, 'Session'] = '15B'


            elif 'Nine Week A' \
                    in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9A'
            elif 'Nine Week A ' \
                    in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9A'
            elif 'Nine Week A1 ' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9A'
            elif 'Nine Week A2 ' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9A'

            elif 'Nine Week A3' \
                    in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9A'
            elif 'Nine Week A4' \
                    in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9A'
            elif 'Nine Week A5' \
                    in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9A'
            elif 'Nine Week A6 ' \
                    in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9A'
            elif 'Nine Week A7 ' \
                    in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9A'
            elif 'Nine Week A8' \
                    in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9A'
            elif 'Nine Week A10 ' \
                    in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9A'
            elif 'Nine Week A11 ' \
                    in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9A'
            elif 'Nine Week A12 ' \
                    in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9A'
            elif 'Nine Week A13 ' \
                    in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9A'
            elif 'Nine Week AJ ' \
                    in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9A'
            elif 'Nine Week AK ' \
                    in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9A'


            elif 'Nine Week B' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9B'
            elif 'Nine Week B ' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9B'
            elif 'Nine Week B1 ' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9B'
            elif 'Nine Week B5 ' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9B'
            elif 'Nine Week B2 ' \
                    in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9B'
            elif 'Nine Week B3 ' \
                    in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9B'
            elif 'Nine Week B4 ' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9B'
            elif 'Nine Week B6 ' \
                    in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9B'
            elif 'Nine Week B7 ' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9B'
            elif 'Nine Week B8 ' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9B'
            elif 'Nine Week B10' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9B'
            elif 'Nine Week B11 ' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9B'
            elif 'Nine Week B12 ' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9B'
            elif 'Nine Week B13 ' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9B'
            elif 'Nine Week BL ' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '9B'


            elif 'Sixteen' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '16'
            elif 'Twelve' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '12'
            elif 'Seven' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '7'
            elif 'Six Week 1 M-F' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '6A'
            elif 'Six Week 2 M-F' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '6B'
            elif 'Six Week 3 M-F' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '6C'
            elif 'Six Week D2 M-Th' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '6A'
            elif 'Six Week B6 M-Th' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '6B'
            elif 'Six Week A2 Tuesday' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '6B'
            elif 'Six Week B2 Tuesday' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '6C'
            elif 'Six Week C4 Thursday' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '6C'
            elif 'Six Week C13 M-Th' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '6B'
            elif 'Six Week B11 Saturday' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '6B'
            elif 'Six Week B4' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '6B'
            elif 'Six Week D2' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '6C'
            elif 'One Week A' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '1'
            elif 'One Week B' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '1'
            elif 'One Week C' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '1'

            elif 'Three Week 1' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '3'
            elif 'Three Week 3AA T, Th' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '3'
            elif 'Three Week 8' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '3'
            elif 'Three Week 3Y M-F' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = '3'

            elif 'Enrollment' in self.enrollment_df.loc[i, 'Session']:
                self.enrollment_df.loc[i, 'Session'] = 'Open'

        self.enrollment_df.to_excel('lab_test.xlsx')
        return self.enrollment_df



class CombineDateframes:

    def __init__(self, course_schedule_df, section_enrollment_df):
        self.course_schedule_df = course_schedule_df
        self.section_enrollment_def = section_enrollment_df

    def merge_dataframes(self):
        merged_df = pd.merge(self.course_schedule_df, self.section_enrollment_def, on=['Class#'])
        merged_df = merged_df[['Combined','Division', 'Dept','Course_x', 'Class#', 'Session', 'Modality', 'Component', 'Start', 'End',
                               'Days', 'Instructor', 'Room', 'Starting Enrollment', 'Current Enrollment','Capacity', 'Perc Drop', 'Graded Count', 'Success Count', 'Success Rate']]
        merged_df.to_excel('Merged_dataframes.xlsx')
        return merged_df

    def combined_course(self, merged_df):
        combined_df = merged_df[merged_df['Combined'] != '0']
        print(combined_df.dtypes)
        combined_list = combined_df['Combined'].unique()

        for i in combined_list:
            courses = i.split('/')
            total_enrollment = 0
            for course in courses:
                course_df = merged_df[merged_df['Course_x'] == course]
                section_list = course_df['Class#'].unique()
                for section in section_list:
                    section_df = course_df[course_df['Class#'] == section]
                    total_enrollment = total_enrollment + section_df.loc[:, 'Current Enrollment']

            print('courses',courses)

    def division_data(self, merged_df):
        division_list = merged_df['Division'].unique()
        print(division_list)
        for div in division_list:
            div_enrollment = merged_df[merged_df['Division'] == div]
            div_enrollment.to_excel(div + '_enrollment.xlsx')





course_schedule_df = pd.read_csv(
    'C:/Users/fmixson/Desktop/Fall 2024 Class Schedule.csv', encoding='latin-1')
pd.set_option('display.max_columns', None)

d = DataframeWork(enrollment_df=course_schedule_df)
# d.sheet_integers()
d.lecture_only()
d.modalities()

# d.division()
d.calculations()
course_schedule_df = d.clean_up_sessions()
course_schedule_df.to_csv('course_schedule.csv')
section_enrollment_df = pd.read_csv(
    'C:/Users/fmixson/PycharmProjects/Enrollment_by_Sections/Enrollment_by_Section.csv')
course_schedule_df = pd.read_csv(
    'C:/Users/fmixson/PycharmProjects/Division_Enrollments/course_schedule.csv'
)
print(section_enrollment_df)
# d.division_data()
e = CombineDateframes(course_schedule_df=course_schedule_df, section_enrollment_df=section_enrollment_df)
merged_df = e.merge_dataframes()
e.combined_course(merged_df=merged_df)
e.division_data(merged_df=merged_df)

