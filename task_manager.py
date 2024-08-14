import datetime
import random
import webbrowser
import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

names = ['Rohitha', 'Chinnodu', 'Junia', 'Swathi']

courses = {
    'AWS': {'classes': 18, 'link': 'https://learn.internshipstudio.com/learn/home/aws/section/307634/lesson/1875569'},
    'Data Scientist': {'classes': 300, 'link': None},
    'DevOps': {'classes': 25, 'link': 'https://learn.internshipstudio.com/learn/home/DevOps/section/471668/lesson/2960749'},
    'Data Structures': {'classes': 28, 'link': None},
    'Problem of the Day': {'classes': None, 'link': 'https://leetcode.com/problemset/'},
    'Implementation "projects"': {'classes': None, 'link': 'C:\\Users\\prave\\Desktop\\project_space'}
}

working_day_timetable = {
    "05:30": "AWS / Data Scientist",
    "18:10": "DevOps / Data Structures",
    "20:00": "Problem of the Day",
    "20:35": "Implementation \"projects\"",
    "23:00": "End of Day"
}

sunday_timetable = {
    "05:30": "AWS / Data Scientist",
    "07:20": "DevOps / Data Structures",
    "09:30": "Recollecting all the week work",
    "20:00": "Problem of the Day",
    "20:35": "Implementation \"projects\"",
    "23:00": "End of Day"
}

task_data = []

def get_greeting():
    current_hour = datetime.datetime.now().hour
    if 5 <= current_hour < 12:
        return "Good morning"
    elif 12 <= current_hour < 17:
        return "Good afternoon"
    else:
        return "Good evening"

def get_current_task(timetable):
    current_time = datetime.datetime.now().time()
    scheduled_task = None
    next_task_time = None

    for time_str, task in timetable.items():
        task_time = datetime.datetime.strptime(time_str, "%H:%M").time()
        if current_time >= task_time:
            scheduled_task = task
        else:
            next_task_time = task_time
            break

    if scheduled_task is None:
        scheduled_task = list(timetable.values())[-1]

    return scheduled_task, next_task_time

def clock_in():
    print("Clocking in...")
    return datetime.datetime.now()

def clock_out(start_time):
    print("Clocking out...")
    end_time = datetime.datetime.now()
    duration = end_time - start_time
    return duration

def format_duration(duration):
    total_seconds = int(duration.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

def open_course(course):
    if course in courses:
        if courses[course]['link']:
            print(f"Opening {course} materials...")
            if course == 'Implementation "projects"':
                os.system(f"code {courses[course]['link']}")
            else:
                webbrowser.open(courses[course]['link'])
        else:
            print(f"No direct link available for {course}. Please navigate to the course materials manually.")
    else:
        print(f"Course '{course}' not found in the system.")

def navigate_courses(start_time):
    while True:
        print("\nHere are the available courses:")
        for i, (course, info) in enumerate(courses.items(), 1):
            classes = f"({info['classes']} classes)" if info['classes'] else ""
            print(f"{i}. {course} {classes}")
        print(f"{len(courses) + 1}. Exit")

        choice = input("\nEnter the number of the course you'd like to work on: ").strip()
        if choice.isdigit():
            choice = int(choice)
            if 1 <= choice <= len(courses):
                selected_course = list(courses.keys())[choice - 1]
                course_start_time = datetime.datetime.now()
                print(f"\nStarting {selected_course} now. Timer is running.")
                open_course(selected_course)
                input("Press Enter when you're done with this course...")
                course_end_time = datetime.datetime.now()
                duration = course_end_time - course_start_time
                task_data.append({
                    'Date': course_start_time.date(),
                    'Course': selected_course,
                    'Start Time': course_start_time.strftime('%H:%M:%S'),
                    'End Time': course_end_time.strftime('%H:%M:%S'),
                    'Duration': format_duration(duration)
                })
            elif choice == len(courses) + 1:
                break
            else:
                print("Invalid choice. Please try again.")
        else:
            print("Invalid input. Please enter a number.")

    duration = clock_out(start_time)
    print(f"You worked for {format_duration(duration)}")

def save_to_excel():
    df = pd.DataFrame(task_data)
    filename = f"task_data_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx"
    df.to_excel(filename, index=False)
    return filename

def send_email(attachment_path):
    sender_email = os.environ.get('GMAIL_USER')
    sender_password = os.environ.get('GMAIL_PASSWORD')
    receiver_email = "praveenmacha777@gmail.com"

    if not sender_email or not sender_password:
        print("Email credentials not set. Please set GMAIL_USER and GMAIL_PASSWORD environment variables.")
        return

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = f"Task Data - {datetime.datetime.now().strftime('%Y-%m-%d')}"

    body = "Please find attached the task data for today."
    msg.attach(MIMEText(body, 'plain'))

    with open(attachment_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {os.path.basename(attachment_path)}",
    )
    msg.attach(part)

    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email. Error: {str(e)}")
        print("Please check your email credentials and internet connection.")

def main():
    start_time = clock_in()
    
    greeting = get_greeting()
    name = random.choice(names)
    print(f"{greeting}, {name}!")

    if datetime.datetime.now().weekday() == 6:
        timetable = sunday_timetable
        print("It's Sunday. Using the Sunday timetable.")
    else:
        timetable = working_day_timetable
        print("It's a working day. Using the regular timetable.")

    current_task, next_task_time = get_current_task(timetable)

    print(f"Current time: {datetime.datetime.now().strftime('%H:%M')}")
    print(f"{name}, you have {current_task} scheduled now.")
    
    if next_task_time:
        print(f"Your next task starts at {next_task_time.strftime('%H:%M')}.")

    if current_task != "End of Day":
        print("Would you like to proceed with the scheduled task? (yes/no)")
        response = input().lower()

        if response in ['yes', 'y']:
            print(f"Starting {current_task} now. Timer is running.")
            if "/" in current_task:
                print("This time slot has multiple courses. Please choose one:")
                courses_for_slot = [c.strip() for c in current_task.split('/')]
                for i, course in enumerate(courses_for_slot, 1):
                    print(f"{i}. {course}")
                choice = int(input("Enter the number of your choice: "))
                chosen_course = courses_for_slot[choice - 1]
                course_start_time = datetime.datetime.now()
                open_course(chosen_course)
                input("Press Enter when you're done with this task...")
                course_end_time = datetime.datetime.now()
                duration = course_end_time - course_start_time
                task_data.append({
                    'Date': course_start_time.date(),
                    'Course': chosen_course,
                    'Start Time': course_start_time.strftime('%H:%M:%S'),
                    'End Time': course_end_time.strftime('%H:%M:%S'),
                    'Duration': format_duration(duration)
                })
            else:
                course_start_time = datetime.datetime.now()
                open_course(current_task)
                input("Press Enter when you're done with this task...")
                course_end_time = datetime.datetime.now()
                duration = course_end_time - course_start_time
                task_data.append({
                    'Date': course_start_time.date(),
                    'Course': current_task,
                    'Start Time': course_start_time.strftime('%H:%M:%S'),
                    'End Time': course_end_time.strftime('%H:%M:%S'),
                    'Duration': format_duration(duration)
                })
        else:
            print(f"No worries, {name}, you can navigate yourself as per your needs.")
    else:
        print("It's the end of your scheduled day. Feel free to rest or choose a task if you want to continue working.")

    navigate_courses(start_time)

    # Save data to Excel and send email
    excel_file = save_to_excel()
    send_email(excel_file)

if __name__ == "__main__":
    main()
