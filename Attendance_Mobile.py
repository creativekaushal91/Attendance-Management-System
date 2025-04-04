from csv import excel
import pandas as pd

#Reading Input Files
df_attendance_dump = pd.read_csv("D://Automation//Attendance Management//input//Attendance_dump.csv")
df_user_data = pd.read_csv("D://Automation//Attendance Management//input//Employee.csv", encoding="ISO-8859-1")
df_leave_data = pd.read_csv("D://Automation//Attendance Management//input//Leave_Data.csv", encoding="ISO-8859-1")


#Cleaning Input Files
df_attendance_dump = df_attendance_dump.dropna(subset = ['olm_id'])
df_attendance_dump = df_attendance_dump[df_attendance_dump['attendance_status'] != 'REJECTED']
df_attendance_dump['olm_id'] = df_attendance_dump['olm_id'].str.upper()
df_attendance_dump['created_date_time'] = pd.to_datetime(df_attendance_dump['created_date_time'], errors = 'coerce')
df_attendance_dump['Day'] = df_attendance_dump['created_date_time'].dt.day_name()
df_attendance_dump['created_date_time'] = df_attendance_dump['created_date_time'].dt.date
df_attendance_dump_mon_fri = df_attendance_dump[df_attendance_dump['Day'].isin(['Monday','Tuesday','Wednesday','Thursday','Friday'])]

df_user_data['olm_id'] = df_user_data['olm_id'].str.upper()
df_user_data = df_user_data.drop_duplicates(subset='olm_id', keep='last')
# Count unique created_date_time per olm_id
att_day_count = df_attendance_dump_mon_fri.groupby('Mob no')['created_date_time'].nunique().reset_index()
att_day_count.columns = ['Mob no', 'Total Present Days(Mon-Fri)']

df_user_data['Mob no'] = df_user_data['Mob no'].astype(str)
att_day_count['Mob no'] = att_day_count['Mob no'].astype(str)


df_user_data = df_user_data.merge(att_day_count[['Mob no', 'Total Present Days(Mon-Fri)']], on='Mob no', how='left')
df_user_data['Total Present Days(Mon-Fri)'] = df_user_data['Total Present Days(Mon-Fri)'].fillna(0)


df_weekends = df_attendance_dump[df_attendance_dump['Day'].isin(['Saturday'])]
weekend_day_count = df_weekends.groupby('Mob no')['created_date_time'].nunique().reset_index()
weekend_day_count.columns = ['Mob no', 'Sat marked']



weekend_day_count['Mob no'] = weekend_day_count['Mob no'].astype(str)

df_user_data = df_user_data.merge(weekend_day_count[['Mob no', 'Sat marked']], on='Mob no', how='left')
df_user_data['Sat marked'] = df_user_data['Sat marked'].fillna(0)
df_user_data['Actual Present Days'] = df_user_data ['Total Present Days(Mon-Fri)'] + df_user_data ['Sat marked']


df_weekends = df_attendance_dump[df_attendance_dump['Day'].isin(['Sunday'])]
weekend_day_count = df_weekends.groupby('Mob no')['created_date_time'].nunique().reset_index()
weekend_day_count.columns = ['Mob no', 'Sun marked']
weekend_day_count['Mob no'] = weekend_day_count['Mob no'].astype(str)
df_user_data = df_user_data.merge(weekend_day_count[['Mob no', 'Sun marked']], on='Mob no', how='left')
df_user_data['Sun marked'] = df_user_data['Sun marked'].fillna(0)
#df_user_data['Actual Present Days'] = df_user_data ['Total Present Days'] + df_user_data ['Sat marked']



# start_date = df_attendance_dump["created_date_time"].min()
# end_date = df_attendance_dump["created_date_time"].max()
# calender = pd.DataFrame({"date": pd.date_range(start=start_date, end=end_date)})
# calender['Day'] = calender['date'].dt.day_name()
# calender = calender[calender['Day'].isin(['Saturday', 'Sunday'])]
# sat_sun_count = calender['date'].nunique()
# #df_user_data ['Default Sat/Sun'] = sat_sun_count



df_leave_data['olm_id'] = df_leave_data['olm_id'].str.upper()

df_leave_data['Leave Taken - CL'] = df_leave_data['Leave Taken - CL'].fillna(0)
df_leave_data['Leave Taken - EL'] = df_leave_data['Leave Taken - EL'].fillna(0)

df_leave_data['Leaves'] = df_leave_data ['Leave Taken - CL'] + df_leave_data ['Leave Taken - EL']


df_leave_data['Mob no'] = df_leave_data['Mob no'].astype(str)
df_user_data = df_user_data.merge(df_leave_data[['Mob no', 'Leaves']], on='Mob no', how='left')
df_user_data['Leaves'] = df_user_data['Leaves'].fillna(0)

# Prompt the user for input
holiday_count = int(input("Enter the number of national/Festivel holidays in the month: "))
df_user_data['Holidays'] = holiday_count


# Apply the condition: if 'Actual Present Days' is 0, set other columns to 0
df_user_data.loc[df_user_data['Actual Present Days'] == 0, [ 'Leaves', 'Holidays']] = 0


df_user_data ['Total Attendance'] = df_user_data['Actual Present Days']  + df_user_data ['Leaves'] + df_user_data['Holidays']


# Define a function to calculate Attendance Status
def attendance_status(row):
    if row['Total Attendance'] >=25 :
        return 'Full Attendance'
    elif 0 < row['Total Attendance'] < 25:
        return 'Partial Attendance'
    elif row['Total Attendance'] == 0:
        return 'Nil Attendance'

# Apply the function to create a new column 'Attendance Status'
df_user_data['Attendance Status'] = df_user_data.apply(attendance_status, axis=1)
#df_user_data = df_user_data.drop_duplicates(subset='olm_id', keep='first')


output_file_path = "D:\\Automation\\Attendance Management\\Attendance report_Mobile.xlsx"
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:    
    df_user_data.to_excel(writer, sheet_name="Employees_Attendance", index=False)
    #df_attendance_dump.to_excel(writer, sheet_name="Attendance Dump", index=False)