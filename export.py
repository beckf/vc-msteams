import veracross_api
import os
import csv
from subprocess import check_output
import config
import mailer
import logging
import datetime

"""
Step 1
Create a config.py that contains the following information in same dir as export.py
- Target Domain is a filter to only export VC people with an email address that include that domain.

config = {'vcuser': 'apiuser',
          'vcpass': 'apipass',
          'vcurl': 'https://api.veracross.com/XX/v2/',
          'logdir': 'C:\\Logs\\vc-msteams',
          'smtp_server': 'mailrelay.domain.org',
          'mail_from': 'mailuser@domain.org',
          'mail_to': 'maillist@domain.org',
          'target_domain': '@domain.org'
          }
"""

"""
Step 2
Install a very specific version of AZCopy in C:\Program Files (x86)\Microsoft SDKs\Azure\AzCopy\AzCopy.exe
The version linked on SDS site is too old and v10 doesn't work with SDS.
Install AzCopy 8.1 x86
Download URL: https://aka.ms/downloadazcopy
"""

"""
Step 3
Install SDS Toolkit from
https://docs.microsoft.com/en-us/schooldatasync/install-the-school-data-sync-toolkit#BK_Install
"""

"""
Step 4
Edit sds_sync.ps1 to reflect your username, profile, path to csv, and path to log dir
"""

"""
Step 5 
Run sds_sync.ps1 once in powershell to save credentials for global admin in SDS.
- Do this as the user that will run the vc-msteams script.  Credentials will be stored in the users credential store.
"""

"""
Step 6 (Optional)
Create a scheduled task to run #>python export.py
"""

# Some Globals #
current_datetime = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
logfile = config.config["logdir"] + "\\vc-msteams_" + str(current_datetime) + ".log"

# Strings to remove from Team Names
bad_naming_text = [';', ',', '!', "*", "/", "(", ")", "&", "Global Online:", "\\",
                   "~", "#", "%", "&", "{", "}", "+", ":", "<", ">", "?", "|", "'"]

# Connect to VC
vc = veracross_api.Veracross(config.config)


# Some Common Functions #
def log(data):
    """
    Logs data to the log file and prints to stdout
    :param data:
    :return:
    """
    if not os.path.exists(config.config["logdir"]):
        os.makedirs(config.config["logdir"])
    logging.basicConfig(filename=logfile, level=logging.INFO)
    logging.info(data)
    print(data)


def match_school_level(school_level):
    # Match school level to teams code
    if 'Lower School' in school_level:
        return '1'
    elif 'Middle School' in school_level:
        return '2'
    elif 'Upper School' in school_level:
        return '3'
    elif 'Preschool' in school_level:
        return '4'
    elif 'All School' in school_level:
        return '3'


def count_csv_lines():
    csv_total_length = 0
    csv_files = os.listdir('csv')
    for file in csv_files:
        if ".csv" in file:
            f = open('./csv/' + file)
            r = csv.reader(f)
            line_count = len(list(r))
            csv_total_length = csv_total_length + line_count
    return csv_total_length


# Gather Information on Last Sync #
previous_csv_total_length = count_csv_lines()
log("Found {} total lines in previous CSV files.".format(previous_csv_total_length))

# Sections #
# Build the sections.csv file

# Create a place holder for class_ids
class_id_set = set()

sections_parameters = {'course_type': '1,2,5,10'}
sections = vc.pull('classes', parameters=sections_parameters)

with open('csv/section.csv', mode='w') as section_file:
    section_writer = csv.writer(section_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
    # Make a header row
    section_writer.writerow(['SIS ID', 'School SIS ID', 'Section Name'])

    for section in sections:
        # Build a set of classes
        class_id_set.add(section["class_pk"])

        course_name = section['class_id'] + " " + section['description']
        # Strip bad characters
        for text in bad_naming_text:
            course_name = course_name.replace(text, '')
            course_name = course_name.replace("  ", ' ')

        # Write section to file
        section_writer.writerow([section['class_pk'], match_school_level(section['school_level']), course_name[0:49]])

# Students #
# Create a place holder for student_ids
student_id_set = set()

students_parameters = {'option': '2'}
students = vc.pull('students', parameters=students_parameters)

with open('csv/student.csv', mode='w') as student_file:
    student_writer = csv.writer(student_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
    # Make a header row
    student_writer.writerow(['SIS ID', 'School SIS ID', 'Username'])

    for student in students:
        if student['school_level']:
            if student['email_1']:
                if config.config['target_domain'] in student['email_1']:
                    # Build a set of student_ids
                    student_id_set.add(student['person_pk'])

                    # Write student to file
                    student_writer.writerow([student['person_pk'], match_school_level(student['school_level']),
                                             student['email_1']])
                else:
                    log("Student: Student {} email address not {}.".format(
                        student['person_pk'], config.config['target_domain']))
            else:
                log("Student: Student {} email address missing.".format(
                    student['person_pk']))
        else:
            log("Student: Student {} does not have a school_level assigned.".format(
                student['person_pk']))

# Teachers #
# Create a place holder for teacher_ids
teacher_id_set = set()

teachers_parameters = {'roles': '1,2'}
teachers = vc.pull('facstaff', parameters=teachers_parameters)

with open('csv/teacher.csv', mode='w') as teacher_file:
    teacher_writer = csv.writer(teacher_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
    # Make a header row
    teacher_writer.writerow(['SIS ID', 'School SIS ID', 'Username'])

    for teacher in teachers:
        # Build a set of teacher_ids
        teacher_id_set.add(teacher['person_pk'])

        if teacher['school_level']:
            if teacher['email_1']:
                if config.config['target_domain'] in teacher['email_1']:
                    # Write teacher to file
                    teacher_writer.writerow([teacher['person_pk'], match_school_level(teacher['school_level']),
                                             teacher['email_1']])
                else:
                    log("Teacher: Teacher {} email address not {}.".format(
                        teacher['person_pk'], config.config['target_domain']))
            else:
                log("Teacher: Teacher {} email address missing.".format(
                    teacher['person_pk']))
        else:
            log("Teacher: Teacher {} does not have a school_level assigned.".format(
                teacher['person_pk']))

# Student Enrollments #
student_enrollments = vc.pull('enrollments')

# Create a dict of classes and count enrollments for stats
student_enrollment_count = {}
for c in class_id_set:
    student_enrollment_count[c] = 0

with open('csv/studentenrollment.csv', mode='w') as student_enrollment_file:
    student_enrollment_writer = csv.writer(student_enrollment_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
    # Make a header row
    student_enrollment_writer.writerow(['SIS ID', 'Section SIS ID'])

    for student_enrollment in student_enrollments:
        if student_enrollment['student_fk'] in student_id_set:
            if student_enrollment['class_fk'] in class_id_set:
                student_enrollment_writer.writerow([student_enrollment['student_fk'], student_enrollment['class_fk']])
                student_enrollment_count[student_enrollment['class_fk']] += 1
            else:
                log("Student Enrollment: Class {} is not in section.csv.".format(student_enrollment['class_fk']))
        else:
            log("Student Enrollment: Student {} is not in student.csv.".format(student_enrollment['student_fk']))

# Teacher Enrollments #

# Create a dict of classes and count enrollments for stats
teacher_enrollment_count = {}
for c in class_id_set:
    teacher_enrollment_count[c] = 0

with open('csv/teacherroster.csv', mode='w') as teacher_roster_file:
    teacher_roster_writer = csv.writer(teacher_roster_file, delimiter=',', quotechar='"',
                                           quoting=csv.QUOTE_MINIMAL)
    # Make a header row
    teacher_roster_writer.writerow(['SIS ID', 'Section SIS ID'])

    for section in sections:
        for teacher in section['teachers']:
            if teacher['person_fk'] in teacher_id_set:
                teacher_roster_writer.writerow([teacher['person_fk'], section['class_pk']])
            else:
                log("Teacher Enrollment: Teacher {} is not in teachers.csv.".format(teacher['person_fk']))

# Stats #
# Match class_pk to class_id
classes = dict()
for s in sections:
    classes[s['class_pk']] = s['class_id']

for k in student_enrollment_count:
    if student_enrollment_count[k] == 0:
        log("Class {} has no student enrollments".format(classes[k]))

# SDS Sync #
current_csv_total_length = count_csv_lines()
log("Found {} total lines in current CSV files.".format(current_csv_total_length))

if current_csv_total_length != previous_csv_total_length:
    log("Sync with SDS required.")
    file_diff = min(previous_csv_total_length,current_csv_total_length) / max(previous_csv_total_length,
                                                                              current_csv_total_length)
    if file_diff > .5:
        log("Diff between previous and current less than 50%. Safe to sync.")
        check_output("C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe .\\sds_sync.ps1", shell=True)

        with open(logfile, 'r') as file:
            message_text = file.read()
        mailer.send_mail_notification(send_from=config.config["mail_from"],
                                      send_to=config.config["mail_to"],
                                      subject="VC-MSTeams Sync Updated SDS",
                                      text=message_text,
                                      server=config.config["smtp_server"])
    else:
        log("Diff between previous and current greater than 50%. Not safe to sync! Exiting. ")
else:
    log("Nothing new to update. SDS will not sync.")

