# TODO: create tool to help admins collect time
# [x] pull all timesheets from folder
# [x]  create array for timesheet data
# [x]  iterate through array of timesheets to start creating employee objects
# [x]  check for timesheets without finished time and notify
# [x]  check for overlapping time and notify
# []  find out if alert on more than one ts per employee/ts from other date or use
# []  after iteration, populate excel sheet with total hours and info for all timesheets
# []  or just add to excel sheet as you iterate

import os
from datetime import datetime
import openpyxl
from tkinter import messagebox

og_timesheets = [f for f in os.listdir("daily_timesheets") if f != ".DS_Store"]

# warn if there are no timesheets in the proper directory
if not len(og_timesheets):
    title = "No Timesheets"
    message = "There are no timesheets in the 'daily_timesheets' folder"
    messagebox.showerror(title, message)


timesheets = []

# handles error message for user
def handle_error(message, file_path):
    title = "Invalid Hours"
    message = message
    answer = messagebox.askquestion(title, message)
    if answer == "yes":
        os.system(f"open -a 'Microsoft Excel' '{os.getcwd()}/{file_path}'")


# iterate over jobs rows
def iter_jobs(start, end, wksht, file_path):
    arr = []
    i = start
    date = wksht["H4"].value
    last_stop_time = None

    for x in range((end + 1) - start):
        job_num = wksht[f"B{i}"].value
        job_desc = wksht[f"C{i}"].value
        tot_hours = wksht[f"K{i}"].value
        start_time = wksht[f"H{i}"].value
        stop_time = wksht[f"I{i}"].value
        indv_tot_hours = wksht[f"K{i}"].value
        cells_check = [
            job_num,
            job_desc,
            tot_hours,
            start_time,
            stop_time,
            indv_tot_hours,
        ]
        hasBustedHours = isinstance(start_time, datetime) or isinstance(
            stop_time, datetime
        )

        # check for incomplete job cells
        if any(cells_check) and any(not d for d in cells_check):
            title = f"Missing {'job number' if not job_num else 'job description' if not job_desc else 'start time' if not start_time else 'stop time' if not stop_time else 'individiual hours total' if not indv_tot_hours else 'hours total'}"
            message = f"A timesheet is missing a {'job number' if not job_num else 'job description' if not job_desc else 'start time' if not start_time else 'stop time' if not stop_time else 'individiual hours total' if not indv_tot_hours else 'hours total'}. Would you like to open the file now?"
            answer = messagebox.askquestion(title, message)
            if answer == "yes":
                os.system(f"open -a 'Microsoft Excel' '{os.getcwd()}/{file_path}'")
                break

        # check for start times that are greater than stop times
        if hasBustedHours and tot_hours <= 0:
            handle_error(
                f"A timesheet for {wksht['h3'].value} has a start time that is greater than the stop time on {job_num}. Would you like to open the file now?",
                file_path,
            )
            break
        elif (
            not hasBustedHours
            and start_time
            and datetime.combine(date, start_time) > datetime.combine(date, stop_time)
        ):
            handle_error(
                f"A timesheet for {wksht['h3'].value} has a start time that is greater than the stop time on {job_num}. Would you like to open the file now?",
                file_path,
            )
            break

        # check for overlapping hours
        if last_stop_time and start_time and last_stop_time > start_time:
            handle_error(
                f"A timesheet for {wksht['h3'].value} has overlapping hours on {job_num}. Would you like to open the file now?",
                file_path,
            )
            break

        # handle hours
        if job_num:
            job = {
                f"job_{x+1}_number": job_num,
                f"job_{x+1}_hours": round(tot_hours, 2),
            }
            arr.append(job)

        last_stop_time = (
            stop_time.time() if isinstance(stop_time, datetime) else stop_time
        )
        i += 1
    return arr


# iterate over contractors rows
def iter_contrs(start, end, wksht, file_path):
    arr = []
    i = start
    date = wksht["H4"].value
    contractor_name = wksht["C18"].value
    contractor_number = wksht["H18"].value

    # check for missing contractor name or number
    if (not contractor_name or not contractor_number) and wksht[f"B{i}"].value:
        handle_error(
            f"A timesheet for {wksht['h3'].value} is missing a contractor {'name' if not contractor_name else 'number'}. Would you like to open the file now?",
            file_path,
        )

    for x in range((end + 1) - start):
        empl_name = wksht[f"B{i}"].value
        empl_role = wksht[f"E{i}"].value
        start_time = wksht[f"I{i}"].value
        stop_time = wksht[f"J{i}"].value
        tot_hours = wksht[f"L{i}"].value
        cells_check = [empl_name, empl_role, start_time, stop_time, tot_hours]
        hasBustedHours = isinstance(start_time, datetime) or isinstance(
            stop_time, datetime
        )

        # check for incomplete contractor cells
        if any(cells_check) and any(not d for d in cells_check):
            title = f"Missing {'contractor employee name' if not empl_name else 'contractor employee role' if not empl_role else 'contractor employee start time' if not start_time else 'contractor employee stop time' if not stop_time else 'contractor hours total'}"
            message = f"A timesheet is missing a {'contractor employee name' if not empl_name else 'contractor employee role' if not empl_role else 'contractor employee start time' if not start_time else 'contractor employee stop time' if not stop_time else 'contractor hours total'}. Would you like to open the file now?"
            answer = messagebox.askquestion(title, message)
            if answer == "yes":
                os.system(f"open -a 'Microsoft Excel' '{os.getcwd()}/{file_path}'")
                break

        # check for start times that are greater than stop times
        if hasBustedHours and tot_hours <= 0:
            handle_error(
                f"A timesheet for {wksht['h3'].value} has a contractor start time that is greater than the stop time on {empl_name}. Would you like to open the file now?",
                file_path,
            )
            break
        elif (
            not hasBustedHours
            and start_time
            and datetime.combine(date, start_time) > datetime.combine(date, stop_time)
        ):
            handle_error(
                f"A timesheet for {wksht['h3'].value} has a contractor start time that is greater than the stop time on {empl_name}. Would you like to open the file now?",
                file_path,
            )
            break

        # handle contractor hours
        if empl_name:
            minutes = str(tot_hours)[3:5]
            minutes_converted = (
                "25"
                if minutes == "15"
                else "50"
                if minutes == "30"
                else "75"
                if minutes == "45"
                else "00"
            )
            contractor = {
                f"contractor_{x+1}_name": empl_name,
                f"contractor_{x+1}_role": empl_role,
                f"contractor_{x+1}_hours": float(
                    (str(tot_hours)[0:3] + minutes_converted).replace(":", ".")
                ),
            }
            arr.append(contractor)

        i += 1
    return arr


# handle equipment
def get_equip(wksht, file_path):
    arr = []
    i = 33
    while i <= 34:
        item = wksht[f"B{i}"].value
        job_num = wksht[f"G{i}"].value
        quant = wksht[f"I{i}"].value
        tot = wksht[f"K{i}"].value
        cells_check = [item, job_num, quant, tot]

        # check for incomplete equipment cells
        if any(cells_check) and any(not d for d in cells_check):
            title = f"Missing {'equipment item name' if not item else 'equipment job number' if not job_num else 'equipment quantity' if not quant else 'equipment total'}"
            message = f"A timesheet is missing an {'equipment item name' if not item else 'equipment job number' if not job_num else 'equipment quantity' if not quant else 'equipment total'}. Would you like to open the file now?"
            answer = messagebox.askquestion(title, message)
            if answer == "yes":
                os.system(f"open -a 'Microsoft Excel' '{os.getcwd()}/{file_path}'")
                break

        if item:
            equip = {
                f"equip_item_{1 if i == 33 else 2}": item,
                f"equip_job_num_{1 if i == 33 else 2}": job_num,
                f"equip_quant_{1 if i == 33 else 2}": quant,
                f"equip_tot_{1 if i == 33 else 2}": tot,
            }
            arr.append(equip)

        i += 1
    return arr


# iterate over sample rows
def iter_samples(start, end, wksht, file_path):
    arr = []
    i = start
    for x in range((end + 1) - start):
        job_num = wksht[f"B{i}"].value
        job_desc = wksht[f"D{i}"].value
        samp_type = wksht[f"H{i}"].value
        quant = wksht[f"K{i}"].value
        cells_check = [job_num, job_desc, samp_type, quant]

        # check for incomplete sample cells
        if any(cells_check) and any(not d for d in cells_check):
            title = f"Missing {'sample job number' if not job_num else 'sample job description' if not job_desc else 'sample type' if not samp_type else 'sample quantity'}"
            message = f"A timesheet is missing a {'sample job number' if not job_num else 'sample job description' if not job_desc else 'sample type' if not samp_type else 'sample quantity'}. Would you like to open the file now?"
            answer = messagebox.askquestion(title, message)
            if answer == "yes":
                os.system(f"open -a 'Microsoft Excel' '{os.getcwd()}/{file_path}'")
                break

        # handle sample datas
        if job_num:
            sample = {
                f"sample_{x+1}_job_num": job_num,
                f"sample_{x+1}_job_desc": job_desc,
                f"sample_{x+1}_type": samp_type,
                f"sample_{x+1}_quant": quant,
            }
            arr.append(sample)

        i += 1
    return arr


# iterate over timesheets and handle data
def handle_timesheets():
    dups_warned = False
    dates_warned = False
    dates = []
    for ts in og_timesheets:
        # load excel file
        file_path = f"daily_timesheets/{ts}"
        wb = openpyxl.load_workbook(file_path, data_only=True)

        # load timesheet
        ws1 = wb["Daily Worksheet"]

        # check for name and date
        name_cell = ws1["H3"].value
        date_cell = ws1["H4"].value
        if not name_cell or not date_cell:
            title = f"Missing {'name' if not name_cell else 'date'}"
            message = f"A timesheet is missing a {'name' if not name_cell else 'date'}. Would you like to open the file now?"
            answer = messagebox.askquestion(title, message)
            if answer == "yes":
                os.system(f"open -a 'Microsoft Excel' '{os.getcwd()}/{file_path}'")
                break
            else:
                break

        # check for problems with total hours
        tot_hours = ws1["k16"].value
        invalid_hours = False
        if type(tot_hours).__name__ == "str" or round(tot_hours, 2) <= 0:
            invalid_hours = True
            title = "Invalid Hours"
            message = f"A timesheet for {ws1['h3'].value} has invalid total hours. Would you like to open the file now?"
            answer = messagebox.askquestion(title, message)
            if answer == "yes":
                os.system(f"open -a 'Microsoft Excel' '{os.getcwd()}/{file_path}'")
                break
            else:
                break

        # get basic timesheet information
        empl_name = ws1["h3"].value
        date = ws1["h4"].value.strftime("%m/%d/%Y")
        jobs = iter_jobs(9, 15, ws1, file_path)
        contractors = iter_contrs(21, 29, ws1, file_path)
        equip = get_equip(ws1, file_path)
        samples = iter_samples(38, 41, ws1, file_path)
        timesheet = {
            "empl_name": empl_name,
            "date": date,
            "jobs": jobs,
            "tot_hours": "invalid hours" if invalid_hours else round(tot_hours, 2),
            "contractors": contractors,
            "equip": equip,
            "samples": samples,
        }

        # warn about multiple timesheets for one person
        if any(
            str.lower(d["empl_name"]) == str.lower(ws1["h3"].value) and not dups_warned
            for d in timesheets
        ):
            title = "Multiple Timesheets"
            message = f"A timesheet for {ws1['h3'].value} has already been added. Would you like to allow all timesheets with duplicate names to be added the final document?"
            answer = messagebox.askquestion(title, message)
            if answer == "yes":
                timesheets.append(timesheet)
                dups_warned = True
                if not date in dates and not len(dates):
                    dates.append(date)
                #  warn about multiple dates
                elif not date in dates and not dates_warned:
                    title = "Multiple Timesheets"
                    message = f"A timesheet for {ws1['h3'].value} has a date of {date}. Would you like to allow timesheets with different dates to be added the final document?"
                    answer = messagebox.askquestion(title, message)
                    if answer == "yes":
                        timesheets.append(timesheet)
                        dates_warned = True
                    else:
                        return
            else:
                return
        else:
            #  warn about multiple dates
            if not date in dates and not len(dates):
                dates.append(date)
                timesheets.append(timesheet)
            elif not date in dates and not dates_warned:
                title = "Multiple Timesheets"
                message = f"A timesheet for {ws1['h3'].value} has a date of {date}. Would you like to allow timesheets with different dates to be added the final document?"
                answer = messagebox.askquestion(title, message)
                if answer == "yes":
                    timesheets.append(timesheet)
                    dates_warned = True
                else:
                    return
            else:
                timesheets.append(timesheet)


handle_timesheets()

print(len(timesheets))
