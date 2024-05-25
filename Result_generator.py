import statistics
import xlrd
import xlwt
from xlutils.copy import copy
import tkinter as tk
from tkinter import filedialog
from tkinter.ttk import *
from tkinter import *
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.pagesizes import letter, A4
from datetime import date
from webbrowser import open

select_subject = "Select Subject result File"
verdana_12 = "Verdana 12"
verdana_10 = "Verdana 10"
select_file = "Select File"
file_type = "Excel Files"
file_locations = [0]*12
no_of_files = 1
students_detail_wb =""
output_folder=""
univ_name = "DEVI AHILYA VISHWAVIDYALAYA, INDORE (M.P.)"
naac_acc = "(Accredited\"A+\" Grade by NAAC )"
add_davv = "Takshashila Campus, Khandwa Road, Indore"
dept_name = "SCHOOL OF INSTRUMENTATION"
stmy_grade =  "GRADE-SHEET"
course_name = ""
branch_name = ""
batch_year = ""

wb_subject = [] 
students_name = []
students_father_name = []
students_mother_name = []
comment_final = []
#for pevious sem details
student_sem_1,credit_sem_1,result_sem_1,attempt_sem_1 =[],[],[],[]
student_sem_2,credit_sem_2,result_sem_2,attempt_sem_2 =[],[],[],[]
student_sem_3,credit_sem_3,result_sem_3,attempt_sem_3 =[],[],[],[]
student_sem_4,credit_sem_4,result_sem_4,attempt_sem_4 =[],[],[],[]
student_sem_5,credit_sem_5,result_sem_5,attempt_sem_5 =[],[],[],[]
student_sem_6,credit_sem_6,result_sem_6,attempt_sem_6 =[],[],[],[]
student_sem_7,credit_sem_7,result_sem_7,attempt_sem_7 =[],[],[],[]
student_sem_8,credit_sem_8,result_sem_8,attempt_sem_8 =[],[],[],[]
student_sem_9,credit_sem_9,result_sem_9,attempt_sem_9 =[],[],[],[]
student_sem_10,credit_sem_10,result_sem_10,attempt_sem_10 =[],[],[],[]
ent_date = ""
#
students_roll_no = []
student_enroloment_no = []
#subjects_grades = [[0]*33]*12
subjects_grades = []
subject_code = []
sem_grade_avg_list=[]
overall_result = []
subject_name = []
course_credits = []
current_sem = ""
month_of_exam = ""
year_of_exam = ""
stud_count=0
sub_count =0
#@
no_of_students = 0
sem_grade_avg = 0
#@
def help():
    open("https://github.com/ArpitMourya/Result_generator")
def generate_result():
    global ent_date
    ent_date = date_text.get()
    global wb_subject
    global students_name; global students_father_name; global students_mother_name
    global comment_final
    #sem details
    global stud_count
    global sub_count
    global student_sem_1;global credit_sem_1;global result_sem_1; global attempt_sem_1
    global student_sem_2;global credit_sem_2;global result_sem_2; global attempt_sem_2
    global student_sem_3;  global credit_sem_3;  global result_sem_3;  global attempt_sem_3
    global student_sem_4;  global credit_sem_4;  global result_sem_4;  global attempt_sem_4
    global student_sem_5;  global credit_sem_5;  global result_sem_5;  global attempt_sem_5
    global student_sem_6;  global credit_sem_6;  global result_sem_6;  global attempt_sem_6
    global student_sem_7;  global credit_sem_7;  global result_sem_7;  global attempt_sem_7
    global student_sem_8;  global credit_sem_8;  global result_sem_8;  global attempt_sem_8
    global student_sem_9;  global credit_sem_9;  global result_sem_9;  global attempt_sem_9
    global student_sem_10; global credit_sem_10; global result_sem_10; global attempt_sem_10
    #
    global students_roll_no;  global student_enroloment_no
    #global subjects_grades
    global subject_code; global subject_name
    global course_name;  global branch_name
    global batch_year;   global current_sem;   global month_of_exam;  global year_of_exam
    global course_credits
    global output_folder
    global no_of_files
    global no_of_students
    wb_subject.clear()
    subject_code.clear()
    subject_name.clear()
    course_credits.clear()

    students_name.clear(),       students_roll_no.clear(),    student_enroloment_no.clear()
    students_father_name.clear(),students_mother_name.clear()
    comment_final.clear()
    student_sem_1.clear(),credit_sem_1.clear(),result_sem_1.clear(),attempt_sem_1.clear()
    student_sem_2.clear(),credit_sem_2.clear(),result_sem_2.clear(),attempt_sem_2.clear()
    student_sem_3.clear(),credit_sem_3.clear(),result_sem_3.clear(),attempt_sem_3.clear()
    student_sem_4.clear(),credit_sem_4.clear(),result_sem_4.clear(),attempt_sem_4.clear()
    student_sem_5.clear(),credit_sem_5.clear(),result_sem_5.clear(),attempt_sem_5.clear()
    student_sem_6.clear(),credit_sem_6.clear(),result_sem_6.clear(),attempt_sem_6.clear()
    student_sem_7.clear(),credit_sem_7.clear(),result_sem_7.clear(),attempt_sem_7.clear()
    student_sem_8.clear(),credit_sem_8.clear(),result_sem_8.clear(),attempt_sem_8.clear()
    student_sem_9.clear(),credit_sem_9.clear(),result_sem_9.clear(),attempt_sem_9.clear()
    student_sem_10.clear(),credit_sem_10.clear(),result_sem_10.clear(),attempt_sem_10.clear()
    try:
        output_folder = filedialog.askdirectory()
        n = 0
        # to get the data from excel sheets and generate the results
        wb_student_details = xlrd.open_workbook(students_detail_wb).sheet_by_index(0)
        #print(wb_student_details.nrows)
        count = 0
        for row in range(wb_student_details.nrows):
            cell_value = wb_student_details.cell_value(row,0)
            if cell_value != '':
                count += 1
        #@ Number of student will be equal to number of rowes in xl sheet - 4
        no_of_students = count-4
        #@
        course_name = wb_student_details.cell_value(0,1)
        branch_name = wb_student_details.cell_value(1,1)
        batch_year = wb_student_details.cell_value(2,1)
        # file locations consists location of subject marks
        # fetch some one time fields which is going to be same for all
        current_sem = xlrd.open_workbook(file_locations[0]).sheet_by_index(0).cell_value(3,3)
        month_of_exam = xlrd.open_workbook(file_locations[0]).sheet_by_index(0).cell_value(4,3)
        year_of_exam = int(xlrd.open_workbook(file_locations[0]).sheet_by_index(0).cell_value(5,3))

        for i in range (count-4):
            students_name.insert(i,wb_student_details.cell_value(i+4,1))
            students_roll_no.insert(i,wb_student_details.cell_value(i+4,2))
            student_enroloment_no.insert(i,wb_student_details.cell_value(i+4,3))
            students_father_name.insert(i,wb_student_details.cell_value(i+4,4))
            students_mother_name.insert(i,wb_student_details.cell_value(i+4,5))
            comment_final.insert(i,wb_student_details.cell_value(i+4,46))
            # To insert data from xl sheet a complete coloum ,in a list
            value_current_sem=check_current_sem(current_sem)
            if value_current_sem >1:
                student_sem_1.insert(i,wb_student_details.cell_value(i+4,6))
                credit_sem_1.insert(i,wb_student_details.cell_value(i+4,16))
                result_sem_1.insert(i,wb_student_details.cell_value(i+4,26))
                attempt_sem_1.insert(i,wb_student_details.cell_value(i+4,36))
            if value_current_sem >2:
                student_sem_2.insert(i,wb_student_details.cell_value(i+4,7))
                credit_sem_2.insert(i,wb_student_details.cell_value(i+4,17))
                result_sem_2.insert(i,wb_student_details.cell_value(i+4,27))
                attempt_sem_2.insert(i,wb_student_details.cell_value(i+4,37))
            if value_current_sem >3:
                student_sem_3.insert(i,wb_student_details.cell_value(i+4,8))
                credit_sem_3.insert(i,wb_student_details.cell_value(i+4,18))
                result_sem_3.insert(i,wb_student_details.cell_value(i+4,28))
                attempt_sem_3.insert(i,wb_student_details.cell_value(i+4,38))
            if value_current_sem >4:
                student_sem_4.insert(i,wb_student_details.cell_value(i+4,9))
                credit_sem_4.insert(i,wb_student_details.cell_value(i+4,19))
                result_sem_4.insert(i,wb_student_details.cell_value(i+4,29))
                attempt_sem_4.insert(i,wb_student_details.cell_value(i+4,39))
            if value_current_sem >5:
                student_sem_5.insert(i,wb_student_details.cell_value(i+4,10))
                credit_sem_5.insert(i,wb_student_details.cell_value(i+4,20))
                result_sem_5.insert(i,wb_student_details.cell_value(i+4,30))
                attempt_sem_5.insert(i,wb_student_details.cell_value(i+4,40))
            if value_current_sem >6:
                student_sem_6.insert(i,wb_student_details.cell_value(i+4,11))
                credit_sem_6.insert(i,wb_student_details.cell_value(i+4,21))
                result_sem_6.insert(i,wb_student_details.cell_value(i+4,31))
                attempt_sem_6.insert(i,wb_student_details.cell_value(i+4,41))
            if value_current_sem >7:
                student_sem_7.insert(i,wb_student_details.cell_value(i+4,12))
                credit_sem_7.insert(i,wb_student_details.cell_value(i+4,22))
                result_sem_7.insert(i,wb_student_details.cell_value(i+4,32))
                attempt_sem_7.insert(i,wb_student_details.cell_value(i+4,42))
            if value_current_sem >8:
                student_sem_8.insert(i,wb_student_details.cell_value(i+4,13))
                credit_sem_8.insert(i,wb_student_details.cell_value(i+4,23))
                result_sem_8.insert(i,wb_student_details.cell_value(i+4,33))
                attempt_sem_8.insert(i,wb_student_details.cell_value(i+4,43))
            if value_current_sem >9:
                student_sem_9.insert(i,wb_student_details.cell_value(i+4,14))
                credit_sem_9.insert(i,wb_student_details.cell_value(i+4,24))
                result_sem_9.insert(i,wb_student_details.cell_value(i+4,34))
                attempt_sem_9.insert(i,wb_student_details.cell_value(i+4,44))
            if value_current_sem >10:
                student_sem_10.insert(i,wb_student_details.cell_value(i+4,15))
                credit_sem_10.insert(i,wb_student_details.cell_value(i+4,25))
                result_sem_10.insert(i,wb_student_details.cell_value(i+4,35))
                attempt_sem_10.insert(i,wb_student_details.cell_value(i+4,45))
        #print(f"students_name{len(students_name)}")
        #print(f"students_roll_no{len(students_roll_no)}")
        #print(f"student_enroloment_no{len(student_enroloment_no)}")
        #print(f"students_father_name{len(students_father_name)}")
        #print(f"students_mother_name{len(students_mother_name)}")
        mod2 = [len(students_name),len(students_father_name),len(students_roll_no),len(student_enroloment_no),len(students_mother_name)]
        #if len(students_name)==len(students_father_name)==len(students_roll_no)==len(student_enroloment_no)==len(students_mother_name):
        stud_count = statistics.mode(mod2)
            #SEM
        for location in range(0,no_of_files):
            wb_subject.insert(n,xlrd.open_workbook(file_locations[n]).sheet_by_index(0))
            # print(wb_subject)
            # Logic for fetching data from and creating pdf
            # fetch the data
            subject_code.insert(n,wb_subject[n].cell_value(0,3))
            subject_name.insert(n,wb_subject[n].cell_value(1,3))
            course_credits.insert(n,wb_subject[n].cell_value(2,3))
            #stud_count = 0
            #for row in range(wb_subject[n].nrows):
            #    cell_value = wb_subject[n].cell_value(row-1,1)
            #    if cell_value != '':
            #        stud_count += 1
            # print("student count")
            grade = []
            for k in range(stud_count):
                grade.append(determine_grade(wb_subject[n].cell_value(k+8,9)))
                #print(f"{n} , { k} \n")
                #print(grade)
            subjects_grades.append(grade)
            n = n+1
        print(f"wb_subject {len(wb_subject)}")
        print(f"subject_code {len(subject_code)}")
        print(f"subject_name {len(subject_name)}")
        print(f"course_credits {len(course_credits)}")
        mod = [len(wb_subject),len(subject_code),len(subject_name),len(course_credits)]
        sub_count = statistics.mode(mod)
        print("student_count ",stud_count,end="\n\n")
        print(mod)
        #print(subjects_grades)
        createpdfs()
        sys_msg.configure(text="Results generated Successfully.", font=verdana_10)
    except Exception as e:
        sys_msg.configure(text=f"ERROR:- {e}", font=verdana_10)
        print(e)

def determine_grade(marks):
    if marks >= 90 and marks <=100:
        return "O"
    elif marks >= 80 and marks <90:
        return "A+"
    elif marks >= 70 and marks <80:
        return "A"
    elif marks >= 60 and marks <70:
        return "B+"
    elif marks >= 50 and marks <60:
        return "B"
    elif marks >= 40 and marks <50:
        return "C"
    elif marks >= 35 and marks <40:
        return "P"
    else:
        return "F"
def check_current_sem(current_sem):
    sem_str = str(current_sem)
    sem_str = sem_str.lower()
    sem = 0
    if "first" in sem_str:
        sem=1
    elif "second" in sem_str:
        sem=2
    elif "third" in sem_str:
        sem=3
    elif "fourth" in sem_str:
        sem=4
    elif "fifth" in sem_str:
        sem=5
    elif "sixth" in sem_str:
        sem=6
    elif "seventh" in sem_str:
        sem=7
    elif "eighth" in sem_str:
        sem=8
    elif "ninth" in sem_str:
        sem=9
    elif "tenth" in sem_str:
        sem=10
    return sem

def getGradeintocredit(cr,grade):
    credit = int(cr)
    if grade=="O":
        return (credit*10)
    elif grade=="A+":
        return (credit*9)
    elif grade=="A":
        return (credit*8)
    elif grade=="B+":
        return (credit*7)
    elif grade=="B":
        return (credit*6)
    elif grade=="C":
        return (credit*5)
    elif grade=="P":
        return (credit*4)
    else:
        return (0)
def createpdfs():
    inch = 72
    global ent_date
    global wb_subject
    is_ATKT_fail = ''
    fail_credits=0
    global students_roll_no
    global student_enroloment_no
    #global subjects_grades
    global sem_grade_avg_list
    global overall_result
    global subject_code
    global subject_name
    result_canvas = []
    result_index = 0
    global current_sem
    global month_of_exam
    global year_of_exam
    global course_credits
    global course_name
    global branch_name
    global batch_year
    global output_folder
    #writing in student_detail_file at result folder (XL)
    rd =xlrd.open_workbook(students_detail_wb)
    wb= copy(rd)
    wb.save(output_folder+"\\"+"student_detail_new.xls")
    #^relate with line number 654
    for student in students_name:

        result_canvas.insert(result_index,Canvas(output_folder+"\\"+student_enroloment_no[result_index]+".pdf",pagesize=A4))
        result_canvas[result_index].setTitle(student)

        '''
        These comments are for future templet of result , this is to print university name at top of the result.
        '''
        # result_canvas[result_index].drawImage(davv_logo,30,710,width=1.5*inch,height=1.5*inch,mask='none')
        #result_canvas[result_index].setFont("Helvetica-Bold",16)
        # result_canvas[result_index].drawCentredString(290,790,univ_name)
        # result_canvas[result_index].setFont("Helvetica-Bold",12)
        # result_canvas[result_index].drawCentredString(295,770,naac_acc)
        result_canvas[result_index].setFont("Helvetica-Bold",20)
        result_canvas[result_index].drawCentredString(295,725,dept_name)
        #result_canvas[result_index].setFont("Helvetica-Bold",12)
        # result_canvas[result_index].drawCentredString(295,735,add_davv)
        # result_canvas[result_index].setFont("Helvetica-Bold",14)
        # result_canvas[result_index].drawCentredString(295,710,stmy_grade)
        # adding 20 in y for making space for signatures at bottom
        result_canvas[result_index].setFont("Helvetica-Bold",14)
        result_canvas[result_index].drawCentredString(300,700,course_name+" "+branch_name)
        result_canvas[result_index].setFont("Helvetica-Bold",14)
        result_canvas[result_index].drawCentredString(300,680,"SEMESTER"+"-"+current_sem.upper()+","+" BATCH"+" "+batch_year)
        result_canvas[result_index].setFont("Helvetica",12)
        result_canvas[result_index].drawString(350,665,"Month of Exam:"+" "+month_of_exam+" "+str(year_of_exam))
        result_canvas[result_index].setFont("Helvetica",12)
        result_canvas[result_index].drawString(40,645,"NAME"+"  :   "+str(student).title())
        if len(students_father_name[result_index])>=30:
            result_canvas[result_index].setFont("Helvetica",12)
            result_canvas[result_index].drawString(300,645,"FATHER'S NAME"+"  :   ")
            result_canvas[result_index].setFont("Helvetica",10)
            result_canvas[result_index].drawString(415,645,str(students_father_name[result_index]).title())
        else:
            result_canvas[result_index].setFont("Helvetica",12)
            result_canvas[result_index].drawString(300,645,"FATHER'S NAME"+"  :   "+str(students_father_name[result_index]).title())

        result_canvas[result_index].setFont("Helvetica",12)
        result_canvas[result_index].drawString(40,605,"ROLL NO"+"  :   "+students_roll_no[result_index])
        result_canvas[result_index].setFont("Helvetica",12)
        if len(students_mother_name[result_index])>=30:

            result_canvas[result_index].drawString(300,625,"MOTHER'S NAME"+"  :   ")
            result_canvas[result_index].setFont("Helvetica",10)
            result_canvas[result_index].drawString(415,625,str(students_mother_name[result_index]).title())
        else:
            result_canvas[result_index].setFont("Helvetica",12)
            result_canvas[result_index].drawString(300,625,"MOTHER'S NAME"+"  :   "+str(students_mother_name[result_index]).title())

        result_canvas[result_index].setFont("Helvetica",12)
        result_canvas[result_index].drawString(300,605,"ENROLMENT NO."+"  :   "+student_enroloment_no[result_index])
        #sem
        result_canvas[result_index].setFont("Helvetica-Bold",10)
        result_canvas[result_index].drawString(270,260,str(comment_final[result_index]))
        result_canvas[result_index].setFont("Helvetica",12)
        result_canvas[result_index].setFont("Helvetica-Bold",10)
        result_canvas[result_index].rect(40, 550, 520, 40, stroke=1, fill=0)
        result_canvas[result_index].rect(40, 310, 60, 280, stroke=1, fill=0)
        
        result_canvas[result_index].rect(40, 310, 335, 280, stroke=1, fill=0)
        result_canvas[result_index].rect(40, 290, 395, 300, stroke=1, fill=0)
        result_canvas[result_index].rect(40, 290, 465, 300, stroke=1, fill=0)
        result_canvas[result_index].rect(40, 510, 520, 40, stroke=1, fill=0)
        result_canvas[result_index].rect(40, 510, 520, 20, stroke=1, fill=0)
        result_canvas[result_index].rect(40, 510, 520, 20, stroke=1, fill=0)
        result_canvas[result_index].rect(40, 490, 520, 20, stroke=1, fill=0)
        result_canvas[result_index].rect(40, 470, 520, 20, stroke=1, fill=0)
        result_canvas[result_index].rect(40, 450, 520, 20, stroke=1, fill=0)
        result_canvas[result_index].rect(40, 430, 520, 20, stroke=1, fill=0)
        result_canvas[result_index].rect(40, 410, 520, 20, stroke=1, fill=0)
        result_canvas[result_index].rect(40, 390, 520, 20, stroke=1, fill=0)
        result_canvas[result_index].rect(40, 370, 520, 20, stroke=1, fill=0)
        result_canvas[result_index].rect(40, 350, 520, 20, stroke=1, fill=0)
        result_canvas[result_index].rect(40, 330, 520, 20, stroke=1, fill=0)

        result_canvas[result_index].drawCentredString(70,575,"COURSE")
        result_canvas[result_index].drawCentredString(70,560,"CODE")
        result_canvas[result_index].drawCentredString(230,565,"SUBJECT NAME")
        result_canvas[result_index].drawCentredString(405,575,"COURSE")
        result_canvas[result_index].drawCentredString(405,560,"CREDITS")
        result_canvas[result_index].drawCentredString(470,575,"GRADE")
        result_canvas[result_index].drawCentredString(470,560,"OBTAINED")
        result_canvas[result_index].drawCentredString(535,575,"GRADE ")
        result_canvas[result_index].drawCentredString(535,565,"POINT")
        result_canvas[result_index].drawCentredString(535,555,"CREDITS")

        result_canvas[result_index].rect(40, 310, 520, 200, stroke=1, fill=0)
        result_canvas[result_index].rect(40, 290, 520, 20, stroke=1, fill=0)
        result_canvas[result_index].rect(40, 290, 335, 20, stroke=1, fill=0)
        result_canvas[result_index].rect(40, 270, 520, 20, stroke=1, fill=0)
        result_canvas[result_index].setFont("Helvetica-Bold",11)
        result_canvas[result_index].drawCentredString(325,295,"TOTAL CREDITS")
        #
        result_canvas[result_index].rect(40, 215, 520, 20, stroke=1, fill=0)
        result_canvas[result_index].setFont("Helvetica-Bold",10)
        result_canvas[result_index].drawCentredString(300,220,"RESULT SEMESTER-WISE")
        result_canvas[result_index].rect(40, 115, 520, 100, stroke=1, fill=0)
        result_canvas[result_index].setFont("Helvetica-Bold",10)
        result_canvas[result_index].drawCentredString(75,200,"SEMESTER")

        result_canvas[result_index].rect(40, 195, 520, 20, stroke=1, fill=0)
        result_canvas[result_index].setFont("Helvetica-Bold",10)
        result_canvas[result_index].drawCentredString(70,180,"CREDITS")
        result_canvas[result_index].rect(40, 175, 520, 20, stroke=1, fill=0)
        result_canvas[result_index].setFont("Helvetica-Bold",10)
        result_canvas[result_index].drawCentredString(63,160,"SGPA")
        result_canvas[result_index].rect(40, 155, 520, 20, stroke=1, fill=0)
        result_canvas[result_index].setFont("Helvetica-Bold",10)
        result_canvas[result_index].drawCentredString(73,140,"ATTEMPT")
        result_canvas[result_index].rect(40, 135, 520, 20, stroke=1, fill=0)
        result_canvas[result_index].setFont("Helvetica-Bold",10)
        result_canvas[result_index].drawCentredString(68,120,"RESULT")
        result_canvas[result_index].rect(40, 115, 520, 20, stroke=1, fill=0)
        #result_canvas[result_index].drawInlineImage(r"C:\Users\DELL\Desktop\New folder\WhatsApp Image 2024-03-08 at 19.40.38_84aba1f4.jpg", 40, 650, width=1.5 * inch, height=1.5 * inch)
        #@to check 5 years or 2 years
        course_branch = course_name+branch_name
        is_five = False
        if "iot" in course_branch.lower() or "internet of things" in course_branch.lower():
            is_five = True
            result_canvas[result_index].rect(160, 115, 40,100, stroke=1, fill=0)
            result_canvas[result_index].drawCentredString(175,200,"I")
            result_canvas[result_index].rect(200, 115, 40, 100, stroke=1, fill=0)
            result_canvas[result_index].drawCentredString(215,200,"II")
            result_canvas[result_index].rect(240, 115, 40, 100, stroke=1, fill=0)
            result_canvas[result_index].drawCentredString(255,200,"III")
            
            result_canvas[result_index].rect(280, 115, 40, 100, stroke=1, fill=0)
            result_canvas[result_index].drawCentredString(295,200,"IV")
            result_canvas[result_index].rect(320, 115, 40, 100, stroke=1, fill=0)
            result_canvas[result_index].drawCentredString(335,200,"V")
            result_canvas[result_index].rect(360, 115, 40, 100, stroke=1, fill=0)
            result_canvas[result_index].drawCentredString(375,200,"VI")
            result_canvas[result_index].rect(400, 115, 40, 100, stroke=1, fill=0)
            result_canvas[result_index].drawCentredString(415,200,"VII")
            result_canvas[result_index].rect(440, 115, 40, 100, stroke=1, fill=0)
            result_canvas[result_index].drawCentredString(455,200,"VIII")
            result_canvas[result_index].rect(480, 115, 40, 100, stroke=1, fill=0)
            result_canvas[result_index].drawCentredString(495,200,"IX")
            result_canvas[result_index].drawCentredString(535,200,"X")
        elif "executive" in course_branch.lower() or "instrumentation" in course_branch.lower():
            result_canvas[result_index].rect(160, 115, 100,100, stroke=1, fill=0)
            result_canvas[result_index].drawCentredString(210,200,"I")
            result_canvas[result_index].rect(260, 115, 100,100, stroke=1, fill=0)
            result_canvas[result_index].drawCentredString(310,200,"II")
            result_canvas[result_index].rect(360, 115, 100,100, stroke=1, fill=0)
            result_canvas[result_index].drawCentredString(410,200,"III")
            result_canvas[result_index].rect(460, 115, 100,100, stroke=1, fill=0)
            result_canvas[result_index].drawCentredString(510,200,"IV")
        #logic for fetching data from xl for previous sem (exclude current sem)
        if check_current_sem(current_sem) >1:
            if is_five:
                result_canvas[result_index].drawString(170,160,str(student_sem_1[result_index]))
                result_canvas[result_index].drawString(170,180,str(int(credit_sem_1[result_index])))
                result_canvas[result_index].drawString(170,120,str(result_sem_1[result_index]))
                result_canvas[result_index].drawString(170,140,str(int(attempt_sem_1[result_index])))
            else :
                result_canvas[result_index].drawString(205,160,str(student_sem_1[result_index]))
                result_canvas[result_index].drawString(205,180,str(int(credit_sem_1[result_index])))
                result_canvas[result_index].drawString(205,120,str(result_sem_1[result_index]))
                result_canvas[result_index].drawString(205,140,str(int(attempt_sem_1[result_index])))
        if check_current_sem(current_sem) >2:
            if is_five:
                result_canvas[result_index].drawString(210,160,str(student_sem_2[result_index]))
                result_canvas[result_index].drawString(210,180,str(int(credit_sem_2[result_index])))
                result_canvas[result_index].drawString(210,120,str(result_sem_2[result_index]))
                result_canvas[result_index].drawString(210,140,str(int(attempt_sem_2[result_index])))
            else:
                result_canvas[result_index].drawString(305,160,str(student_sem_2[result_index]))
                result_canvas[result_index].drawString(305,180,str(int(credit_sem_2[result_index])))
                result_canvas[result_index].drawString(305,120,str(result_sem_2[result_index]))
                result_canvas[result_index].drawString(305,140,str(int(attempt_sem_2[result_index])))
        if check_current_sem(current_sem) >3:
            if is_five:
                result_canvas[result_index].drawString(250,160,str(student_sem_3[result_index]))
                result_canvas[result_index].drawString(250,180,str(int(credit_sem_3[result_index])))
                result_canvas[result_index].drawString(250,120,str(result_sem_3[result_index]))
                result_canvas[result_index].drawString(250,140,str(int(attempt_sem_3[result_index])))
            else:
                result_canvas[result_index].drawString(405,160,str(student_sem_3[result_index]))
                result_canvas[result_index].drawString(405,180,str(int(credit_sem_3[result_index])))
                result_canvas[result_index].drawString(405,120,str(result_sem_3[result_index]))
                result_canvas[result_index].drawString(405,140,str(int(attempt_sem_3[result_index])))
        if check_current_sem(current_sem) >4:
            result_canvas[result_index].drawString(290,160,str(student_sem_4[result_index]))
            result_canvas[result_index].drawString(290,180,str(int(credit_sem_4[result_index])))
            result_canvas[result_index].drawString(290,120,str(result_sem_4[result_index]))
            result_canvas[result_index].drawString(290,140,str(int(attempt_sem_4[result_index])))
        if check_current_sem(current_sem) >5:
            result_canvas[result_index].drawString(330,160,str(student_sem_5[result_index]))
            result_canvas[result_index].drawString(330,180,str(int(credit_sem_5[result_index])))
            result_canvas[result_index].drawString(330,120,str(result_sem_5[result_index]))
            result_canvas[result_index].drawString(330,140,str(int(attempt_sem_5[result_index])))
        if check_current_sem(current_sem) >6:
            result_canvas[result_index].drawString(370,160,str(student_sem_6[result_index]))
            result_canvas[result_index].drawString(370,180,str(int(credit_sem_6[result_index])))
            result_canvas[result_index].drawString(370,120,str(result_sem_6[result_index]))
            result_canvas[result_index].drawString(370,140,str(int(attempt_sem_6[result_index])))
        if check_current_sem(current_sem) >7:
            result_canvas[result_index].drawString(410,160,str(student_sem_7[result_index]))
            result_canvas[result_index].drawString(410,180,str(int(credit_sem_7[result_index])))
            result_canvas[result_index].drawString(410,120,str(result_sem_7[result_index]))
            result_canvas[result_index].drawString(410,140,str(int(attempt_sem_7[result_index])))
        if check_current_sem(current_sem) >8:
            result_canvas[result_index].drawString(450,160,str(student_sem_8[result_index]))
            result_canvas[result_index].drawString(450,180,str(int(credit_sem_8[result_index])))
            result_canvas[result_index].drawString(450,120,str(result_sem_8[result_index]))
            result_canvas[result_index].drawString(450,140,str(int(attempt_sem_8[result_index])))
        if check_current_sem(current_sem) >9:
            result_canvas[result_index].drawString(490,160,str(student_sem_9[result_index]))
            result_canvas[result_index].drawString(490,180,str(int(credit_sem_9[result_index])))
            result_canvas[result_index].drawString(490,120,str(result_sem_9[result_index]))
            result_canvas[result_index].drawString(490,140,str(int(attempt_sem_9[result_index])))

        result_index = result_index + 1
    result_index = 0
    #to set co-ordinates to write(edit) sgpa in xl sheet
    strt_index_r,strt_index_c=0,0
    if check_current_sem(current_sem)==1:
        strt_index_r,strt_index_c=4,6
    elif check_current_sem(current_sem)==2:
        strt_index_r,strt_index_c=4,7
    elif check_current_sem(current_sem)==3:
        strt_index_r,strt_index_c=4,8
    elif check_current_sem(current_sem)==4:
        strt_index_r,strt_index_c=4,9
    elif check_current_sem(current_sem)==5:
        strt_index_r,strt_index_c=4,10
    elif check_current_sem(current_sem)==6:
        strt_index_r,strt_index_c=4,11
    elif check_current_sem(current_sem)==7:
        strt_index_r,strt_index_c=4,12
    elif check_current_sem(current_sem)==8:
        strt_index_r,strt_index_c=4,13
    elif check_current_sem(current_sem)==9:
        strt_index_r,strt_index_c=4,14
    elif check_current_sem(current_sem)==10:
        strt_index_r,strt_index_c=4,15
    print("stud_count",stud_count)
    print('sub_count',sub_count)
    for student in range(stud_count):
        is_ATKT_fail = 'PASS'
        fail_credits = 0
        result_canvas[result_index].setFont("Helvetica",11)
        start_x = 70
        start_y = 535
        for sub_code in subject_code:
            #print(sub_code)
            result_canvas[result_index].drawCentredString(start_x,start_y,sub_code)
            start_y = start_y-20

        start_x = 105
        start_y = 535
        for subj in subject_name:
            #print(sub_name)
            result_canvas[result_index].drawString(start_x,start_y,subj)
            start_y = start_y-20

        start_x = 405
        start_y = 535
        for credit in course_credits:
            result_canvas[result_index].drawString(start_x,start_y,str(int(credit)))
            start_y = start_y-20

        result_canvas[result_index].setFont("Helvetica",12)
        result_canvas[result_index].drawCentredString(405,295,str(int(sum(course_credits))))

        start_x = 470
        start_y = 535
        for i in range(sub_count):
            result_canvas[result_index].drawString(start_x,start_y,(subjects_grades[i][result_index]))
            if 'F' in subjects_grades[i][result_index]:
                is_ATKT_fail = 'ATKT'
                fail_credits += course_credits[i]
            start_y = start_y-20
        if fail_credits >=12:
            is_ATKT_fail = "FAIL"
            #else:
            #    is_ATKT_fail = "PASS"


        start_x = 525
        start_y = 535
        grade_credit_sum = 0
        for i in range(sub_count):
            grade_credit = getGradeintocredit(course_credits[i],subjects_grades[i][result_index])
            grade_credit_sum = grade_credit_sum + grade_credit
            result_canvas[result_index].drawString(start_x,start_y,(str)(grade_credit))
            start_y = start_y-20

        result_canvas[result_index].setFont("Helvetica",12)
        result_canvas[result_index].drawCentredString(535,295,str(int(grade_credit_sum)))

        sem_grade_avg = grade_credit_sum/sum(course_credits)
        result_canvas[result_index].drawString(50,275,"Semester Grade Point Average(SGPA) = "+str(round(sem_grade_avg,3)))
        # FOR ATKT/BACKLOG
        #

        #
        #LIST FOR PRINTING SGPA AND CREDITS OF CURRENT SEM at pdf
        if is_five:
            co_ordinate = [0,170,210,250,290,330,370,410,450,490,530]
        else:
            co_ordinate = [0,205,305,405,390,430,470,510,550,590,630]
        result_canvas[result_index].setFont("Helvetica-Bold",10)
        result_canvas[result_index].drawString(co_ordinate[check_current_sem(current_sem)],160,str(round(sem_grade_avg,3)))
        sem_grade_avg_list.append(round(sem_grade_avg,3))
        overall_result.append(is_ATKT_fail)
        #credits
        result_canvas[result_index].drawString(co_ordinate[check_current_sem(current_sem)],120,str(is_ATKT_fail))
        #if "PASS" in is_ATKT_fail.upper():
        result_canvas[result_index].drawString(co_ordinate[check_current_sem(current_sem)],140,str(1))
        result_canvas[result_index].drawString(co_ordinate[check_current_sem(current_sem)],180,str(int(sum(course_credits))))

        today = date.today()
        data_of_issue = today.strftime("%d %B %Y")
        if (ent_date != ""):
            data_of_issue = ent_date

        print(data_of_issue)
        result_canvas[result_index].drawString(50,260,"*Grade in repeated Examination")
        result_canvas[result_index].drawString(50,100,"DATE OF RESULT: "+data_of_issue)
        result_canvas[result_index].drawString(50,60,"CHECKED BY")
        result_canvas[result_index].drawString(280,60,"RESULT")
        result_canvas[result_index].drawString(260,50,"CO-ORDINATOR")
        result_canvas[result_index].drawString(470,60,"HEAD")
        result_canvas[result_index].save()
        result_index = result_index + 1
        #writing in student_detail_file (XL)

        rd1 =xlrd.open_workbook(output_folder+"\\"+"student_detail_new.xls")
        wb1= copy(rd1)
        w_sheet = wb1.get_sheet(0)
        w_sheet.write(strt_index_r,strt_index_c,float(round(sem_grade_avg,3)))
        wb1.save(output_folder+"\\"+"student_detail_new.xls")
        w_sheet.write(strt_index_r,strt_index_c+10,int(sum(course_credits)))
        wb1.save(output_folder+"\\"+"student_detail_new.xls")
        w_sheet.write(strt_index_r,strt_index_c+20,str(is_ATKT_fail))
        wb1.save(output_folder+"\\"+"student_detail_new.xls")
        w_sheet.write(strt_index_r,strt_index_c+30,1)
        wb1.save(output_folder+"\\"+"student_detail_new.xls")
        strt_index_r+=1
    creat_master_xlsheet()
def creat_master_xlsheet():
    wrte = xlwt.Workbook()
    ws =wrte.add_sheet("master_sheet")
    ws.write(1,0,"Sr No."),ws.write(1,1,"Student Name"),ws.write(1,2,"Roll Number"),ws.write(2,3,"Credits :-"),ws.write(1,3,"Enrolment Number")
    current_colum = 3
    for sub in subject_name:
        ws.write(0,current_colum+1,subject_code[current_colum-3])
        ws.write(1,current_colum+1,sub)
        ws.write(2,current_colum+1,course_credits[current_colum-3])
        current_colum +=1
    ws.write(1,current_colum+1,"SGPA")
    ws.write(1,current_colum+2,"Result")
    sem_count_iterator = 1
    for i in range(current_colum+2,current_colum+1+check_current_sem(current_sem)):
        ws.write(1,i+1,f"Sem-{sem_count_iterator}")
        if sem_count_iterator ==1:
            for j in range(stud_count):
                ws.write(3+j,current_colum+2+1,student_sem_1[j])
        if sem_count_iterator ==2:
            for j in range(stud_count):
                ws.write(3+j,current_colum+3+1,student_sem_2[j])
        if sem_count_iterator ==3:
            for j in range(stud_count):
                ws.write(3+j,current_colum+4+1,student_sem_3[j])
        if sem_count_iterator ==4:
            for j in range(stud_count):
                ws.write(3+j,current_colum+5+1,student_sem_4[j])
        if sem_count_iterator ==5:
            for j in range(stud_count):
                ws.write(3+j,current_colum+6+1,student_sem_5[j])
        if sem_count_iterator ==6:
            for j in range(stud_count):
                ws.write(3+j,current_colum+7+1,student_sem_6[j])
        if sem_count_iterator ==7:
            for j in range(stud_count):
                ws.write(3+j,current_colum+8+1,student_sem_7[j])
        if sem_count_iterator ==8:
            for j in range(stud_count):
                ws.write(3+j,current_colum+9+1,student_sem_8[j])
        if sem_count_iterator ==9:
            for j in range(stud_count):
                ws.write(3+j,current_colum+10+1,student_sem_9[j])
        sem_count_iterator+=1
    row_num =1
    for name in students_name:
        ws.write(row_num+2,0,row_num)
        ws.write(row_num+2,1,name)
        ws.write(row_num+2,2,students_roll_no[row_num-1])
        ws.write(row_num+2,3,student_enroloment_no[row_num-1])
        for sub_no in range(0,len(subject_name)):
            ws.write(row_num+2,sub_no + 4,subjects_grades[sub_no][row_num-1])
        row_num+=1
    row_for_sgpa =1
    for sgpa in sem_grade_avg_list:
        ws.write(row_for_sgpa+2,4+len(subject_name),sgpa)
        ws.write(row_for_sgpa+2,5+len(subject_name),overall_result[row_for_sgpa-1])
        #ws.write(row_num,1,name)
        row_for_sgpa+=1
    wrte.save( output_folder + "//"+ 'master sheet.xls')
def browse_file(i):
    global file_locations
    global no_of_files
    #print(file_locations)
    global students_detail_wb
    filename = filedialog.askopenfilename(
        parent=root, title=select_file, filetypes=((file_type, "*.xls*"), ("All Files", "*.*")))
    # change label content
    if i == 0:
        student_info_select.configure(
            text=filename, fg="black", bg="white", font=verdana_10, width=70)
        students_detail_wb = filename
    elif i == 1:
        subject_1_select.configure(
            text=filename, fg="black", bg="white", font=verdana_10, width=70)
        file_locations[i-1] = filename
        no_of_files = i
    elif i == 2:
        subject_2_select.configure(
            text=filename, fg="black", bg="white", font=verdana_10, width=70)
        file_locations[i-1] = filename
        no_of_files = i
    elif i == 3:
        subject_3_select.configure(
            text=filename, fg="black", bg="white", font=verdana_10, width=70)
        file_locations[i-1] = filename
        no_of_files = i
    elif i == 4:
        subject_4_select.configure(
            text=filename, fg="black", bg="white", font=verdana_10, width=70)
        file_locations[i-1] = filename
        no_of_files = i
    elif i == 5:
        subject_5_select.configure(
            text=filename, fg="black", bg="white", font=verdana_10, width=70)
        file_locations[i-1] = filename
        no_of_files = i
    elif i == 6:
        subject_6_select.configure(
            text=filename, fg="black", bg="white", font=verdana_10, width=70)
        file_locations[i-1] = filename
        no_of_files = i
    elif i == 7:
        subject_7_select.configure(
            text=filename, fg="black", bg="white", font=verdana_10, width=70)
        file_locations[i-1] = filename
        no_of_files = i
    elif i == 8:
        subject_8_select.configure(
            text=filename, fg="black", bg="white", font=verdana_10, width=70)
        file_locations[i-1] = filename
        no_of_files = i
    elif i == 9:
        subject_9_select.configure(
            text=filename, fg="black", bg="white", font=verdana_10, width=70)
        file_locations[i-1] = filename
        no_of_files = i
    elif i == 10:
        subject_10_select.configure(
            text=filename, fg="black", bg="white", font=verdana_10, width=70)
        file_locations[i-1] = filename
        no_of_files = i
    elif i == 11:
        subject_11_select.configure(
            text=filename, fg="black", bg="white", font=verdana_10, width=70)
        file_locations[i-1] = filename
        no_of_files = i
    elif i == 12:
        subject_12_select.configure(
            text=filename, fg="black", bg="white", font=verdana_10, width=70)
        file_locations[i-1] = filename
        no_of_files = i
# create the root
root = tk.Tk()
root.geometry("1000x600")
root.resizable(False, False)
root.title("SOI DAVV")
# create the widgets
department = tk.Label(
    root, text="School of Instrumentation, DAVV, Indore", fg="green", font="Verdana 20")
tool = tk.Label(root, text="PDF Generator Software",
                fg="black", font="Verdana 16")
note = tk.Label(
    root, text="Note : Please select the excel files of subject results and student_info and click on generate results.", font=verdana_12)
#--------------------------------
today1 = date.today()
curr_date = today1.strftime("%d %B %Y")
date_text = tk.Entry(font=verdana_10)
date_sel = tk.Label(root, text='DATE  \n'+f"{curr_date}",
                fg="black", font="Verdana 10")

# __________________________________________________________________________________________________________________________________________
student_info = tk.Label(root, text="Students Info :", font=verdana_10)
student_info_select = tk.Label(
    root, text="Select Student Info File", fg="gray", bg="white", font=verdana_10, width=70)
button_explore = tk.Button(root, text=select_file,
                           command=lambda: browse_file(0))
# __________________________________________________________________________________________________________________________________________
subject_1 = tk.Label(root, text="Subject 1 :", font=verdana_10)
subject_1_select = tk.Label(
    root, text=select_subject, fg="gray", bg="white", font=verdana_10, width=70)
subject_1_explore = tk.Button(root, text=select_file,
                              command=lambda: browse_file(1))
# __________________________________________________________________________________________________________________________________________
subject_2 = tk.Label(root, text="Subject 2 :", font=verdana_10)
subject_2_select = tk.Label(
    root, text=select_subject, fg="gray", bg="white", font=verdana_10, width=70)
subject_2_explore = tk.Button(root, text=select_file,
                              command=lambda: browse_file(2))
# __________________________________________________________________________________________________________________________________________
subject_3 = tk.Label(root, text="Subject 3 :", font=verdana_10)
subject_3_select = tk.Label(
    root, text=select_subject, fg="gray", bg="white", font=verdana_10, width=70)
subject_3_explore = tk.Button(root, text=select_file,
                              command=lambda: browse_file(3))
# __________________________________________________________________________________________________________________________________________
subject_4 = tk.Label(root, text="Subject 4 :", font=verdana_10)
subject_4_select = tk.Label(
    root, text=select_subject, fg="gray", bg="white", font=verdana_10, width=70)
subject_4_explore = tk.Button(root, text=select_file,
                              command=lambda: browse_file(4))
# __________________________________________________________________________________________________________________________________________
subject_5 = tk.Label(root, text="Subject 5 :", font=verdana_10)
subject_5_select = tk.Label(
    root, text=select_subject, fg="gray", bg="white", font=verdana_10, width=70)
subject_5_explore = tk.Button(root, text=select_file,
                              command=lambda: browse_file(5))
# __________________________________________________________________________________________________________________________________________
subject_6 = tk.Label(root, text="Subject 6 :", font=verdana_10)
subject_6_select = tk.Label(
    root, text=select_subject, fg="gray", bg="white", font=verdana_10, width=70)
subject_6_explore = tk.Button(root, text=select_file,
                              command=lambda: browse_file(6))
# __________________________________________________________________________________________________________________________________________
subject_7 = tk.Label(root, text="Subject 7 :", font=verdana_10)
subject_7_select = tk.Label(
    root, text=select_subject, fg="gray", bg="white", font=verdana_10, width=70)
subject_7_explore = tk.Button(root, text=select_file,
                              command=lambda: browse_file(7))
# __________________________________________________________________________________________________________________________________________
subject_8 = tk.Label(root, text="Subject 8 :", font=verdana_10)
subject_8_select = tk.Label(
    root, text=select_subject, fg="gray", bg="white", font=verdana_10, width=70)
subject_8_explore = tk.Button(root, text=select_file,
                              command=lambda: browse_file(8))
# __________________________________________________________________________________________________________________________________________
subject_9 = tk.Label(root, text="Subject 9 :", font=verdana_10)
subject_9_select = tk.Label(
    root, text=select_subject, fg="gray", bg="white", font=verdana_10, width=70)
subject_9_explore = tk.Button(root, text=select_file,
                              command=lambda: browse_file(9))
# __________________________________________________________________________________________________________________________________________
subject_10 = tk.Label(root, text="Subject 10 :", font=verdana_10)
subject_10_select = tk.Label(
    root, text=select_subject, fg="gray", bg="white", font=verdana_10, width=70)
subject_10_explore = tk.Button(root, text=select_file,
                               command=lambda: browse_file(10))
# __________________________________________________________________________________________________________________________________________

subject_11 = tk.Label(root, text="Subject 11 :", font=verdana_10)
subject_11_select = tk.Label(
    root, text=select_subject, fg="gray", bg="white", font=verdana_10, width=70)
subject_11_explore = tk.Button(root, text=select_file,
                               command=lambda: browse_file(11))
# __________________________________________________________________________________________________________________________________________
subject_12 = tk.Label(root, text="Subject 12 :", font=verdana_10)
subject_12_select = tk.Label(
    root, text=select_subject, fg="gray", bg="white", font=verdana_10, width=70)
subject_12_explore = tk.Button(root, text=select_file,
                               command=lambda: browse_file(12))
# __________________________________________________________________________________________________________________________________________

gen_result = tk.Button(root, text="Generate Results",font=verdana_12,
                               command=generate_result)
get_help = tk.Button(root, text="Help",font=verdana_10,
                               command=help)
sys_msg = tk.Label(root, text="Click on 'Generate Results' to get result in PDF format.", font=verdana_10)

# pack the widgets
department.grid(row=0, column=0, columnspan=3)
tool.grid(row=1, column=0, columnspan=3)
note.grid(row=2, column=0, columnspan=3,padx=20,pady=20)

date_text.grid(row=17,column=2,columnspan=2,pady=5)
date_sel.grid(row=16,column=2,columnspan=2,pady=5)


student_info.grid(row=3, column=0, sticky=W,padx=20)
student_info_select.grid(row=3, column=1)
button_explore.grid(row=3, column=2)

subject_1.grid(row=4, column=0, sticky=W,padx=20)
subject_1_select.grid(row=4, column=1)
subject_1_explore.grid(row=4, column=2)

subject_2.grid(row=5, column=0, sticky=W,padx=20)
subject_2_select.grid(row=5, column=1)
subject_2_explore.grid(row=5, column=2)

subject_3.grid(row=6, column=0, sticky=W,padx=20)
subject_3_select.grid(row=6, column=1)
subject_3_explore.grid(row=6, column=2)

subject_4.grid(row=7, column=0, sticky=W,padx=20)
subject_4_select.grid(row=7, column=1)
subject_4_explore.grid(row=7, column=2)

subject_5.grid(row=8, column=0, sticky=W,padx=20)
subject_5_select.grid(row=8, column=1)
subject_5_explore.grid(row=8, column=2)

subject_6.grid(row=9, column=0, sticky=W,padx=20)
subject_6_select.grid(row=9, column=1)
subject_6_explore.grid(row=9, column=2)

subject_7.grid(row=10, column=0, sticky=W,padx=20)
subject_7_select.grid(row=10, column=1)
subject_7_explore.grid(row=10, column=2)

subject_8.grid(row=11, column=0, sticky=W,padx=20)
subject_8_select.grid(row=11, column=1)
subject_8_explore.grid(row=11, column=2)

subject_9.grid(row=12, column=0, sticky=W,padx=20)
subject_9_select.grid(row=12, column=1)
subject_9_explore.grid(row=12, column=2)

subject_10.grid(row=13, column=0, sticky=W,padx=20)
subject_10_select.grid(row=13, column=1)
subject_10_explore.grid(row=13, column=2)

subject_11.grid(row=14, column=0, sticky=W,padx=20)
subject_11_select.grid(row=14, column=1)
subject_11_explore.grid(row=14, column=2)

subject_12.grid(row=15, column=0, sticky=W,padx=20)
subject_12_select.grid(row=15, column=1)
subject_12_explore.grid(row=15, column=2)

gen_result.grid(row=16,column=0,columnspan=3,pady=10)
get_help.grid(row=17,column=0,columnspan=1,pady=11)
sys_msg.grid(row=17,column=0,columnspan=3,pady=10)
# calling the mainloop
root.mainloop()
