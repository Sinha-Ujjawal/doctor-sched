import pandas as pd
from pandas import ExcelWriter
from doctor_schedule import generate_schedule, generate_month_dates, weeks, all_shifts
from doctors import *


def main():
    year = 2024
    month = 11
    all_doctors = [
        DR_MINAKSHI_MISHRA,
        DR_ABHILASHA_MISHRA,
        DR_AMIT_TRIPATHI,
        DR_SUVARNA_KUMAR,
        DR_RASHMI_SHARMA,
        DR_SHARMILA,
        DR_SAUMYA_SHUKLA,
        DR_MADHURI_TRIPATHI,
    ]
    dates = generate_month_dates(year, month)
    fixed_shifts = {}
    unavailable_shifts = {}
    leaves = [
        (DR_MADHURI_TRIPATHI, 1),
        (DR_MADHURI_TRIPATHI, 7),
        (DR_MADHURI_TRIPATHI, 8),
        (DR_MADHURI_TRIPATHI, 9),
        (DR_MADHURI_TRIPATHI, 23),
        (DR_MADHURI_TRIPATHI, 24),
        (DR_MADHURI_TRIPATHI, 25),

        (DR_AMIT_TRIPATHI,    1),
        (DR_AMIT_TRIPATHI,    7),
        (DR_AMIT_TRIPATHI,    8),
        (DR_AMIT_TRIPATHI,    9),
        (DR_AMIT_TRIPATHI,    23),
        (DR_AMIT_TRIPATHI,    24),
        (DR_AMIT_TRIPATHI,    25),

        (DR_SUVARNA_KUMAR,    9),
        (DR_SUVARNA_KUMAR,    10),
        (DR_SUVARNA_KUMAR,    11),
        (DR_SUVARNA_KUMAR,    12),
        (DR_SUVARNA_KUMAR,    13),
        (DR_SUVARNA_KUMAR,    14),
        (DR_SUVARNA_KUMAR,    15),
        (DR_SUVARNA_KUMAR,    16),

        (DR_SAUMYA_SHUKLA,    7),

        (DR_SHARMILA,         8),
        (DR_SHARMILA,         9),
        (DR_SHARMILA,         10),

        (DR_RASHMI_SHARMA,    23),
        (DR_RASHMI_SHARMA,    24),
        (DR_RASHMI_SHARMA,    25),
        (DR_RASHMI_SHARMA,    26),
        (DR_RASHMI_SHARMA,    27),
    ]
    for dt in dates:
        day = dt.day
        unavailable_shifts[(DR_RASHMI_SHARMA, day)]    = ["ot_duty"]
        if dt.weekday() == weeks.index("Sun"):
            unavailable_shifts[(DR_MADHURI_TRIPATHI, day)] = ["ot_duty", "night"]
        else:
            unavailable_shifts[(DR_MADHURI_TRIPATHI, day)] = ["ot_duty", "morning", "night"]
        unavailable_shifts[(DR_SHARMILA, day)]         = ["ot_duty"]
        unavailable_shifts[(DR_MINAKSHI_MISHRA, day)]  = ["ot_duty"]
        if dt.weekday() == weeks.index("Sun"):
            unavailable_shifts[(DR_AMIT_TRIPATHI, day)] = ["morning"]
        if dt.weekday() == weeks.index("Wed"):
            unavailable_shifts[(DR_ABHILASHA_MISHRA, day)] = ["night"]
        #if dt.weekday() == weeks.index("Sun"):
        #    unavailable_shifts[(DR_MADHURI_TRIPATHI, day)] = all_shifts
    for doctor, day in leaves:
        unavailable_shifts[(doctor, day)] = all_shifts
    #df_prefilled = pd.read_excel("./schedule_2024_11_prefilled.xlsx").fillna("")
    #for record in df_prefilled.to_dict(orient="records"):
    #    dt = record["date"]
    #    shifts_by_doctor = {}
    #    for shift in all_shifts:
    #        doctor = record[shift]
    #        if doctor:
    #            if doctor not in shifts_by_doctor:
    #                shifts_by_doctor[doctor] = [shift]
    #            else:
    #                shifts_by_doctor[doctor].append(shift)
    #    for doctor, shifts in shifts_by_doctor.items():
    #        fixed_shifts[(doctor, dt.day)] = shifts
    for dt in dates:
        day = dt.day
        week = dt.weekday()

        if week == weeks.index("Mon"):
            fixed_shifts[(DR_ABHILASHA_MISHRA, day)] = ["ot_duty"]
        if week == weeks.index("Tue"):
            fixed_shifts[(DR_AMIT_TRIPATHI, day)] = ["ot_duty"]
        if week == weeks.index("Wed"):
            fixed_shifts[(DR_AMIT_TRIPATHI, day)] = ["ot_duty"]
        if week == weeks.index("Thu"):
            fixed_shifts[(DR_SAUMYA_SHUKLA, day)] = ["ot_duty"]
        if week == weeks.index("Fri") and day != 8:
            fixed_shifts[(DR_SUVARNA_KUMAR, day)] = ["ot_duty"]

        if week == weeks.index("Tue"):
            fixed_shifts[(DR_SUVARNA_KUMAR, day)] = ["night"]
        if week != weeks.index("Sun"):
            fixed_shifts[(DR_MADHURI_TRIPATHI, day)] = ["evening"]

        #if day == 4:
        #    fixed_shifts[(DR_SAUMYA_SHUKLA, day)] = ["night"]
        #if day == 6:
        #    fixed_shifts[(DR_ABHILASHA_MISHRA, day)] = ["night"]
        #if day == 8:
        #    fixed_shifts[(DR_SHARMILA, day)] = ["night"]
        #if day == 9:
        #    fixed_shifts[(DR_MINAKSHI_MISHRA, day)] = ["night"]
        #if day == 10:
        #    fixed_shifts[(DR_RASHMI_SHARMA, day)] = ["night"]


    max_night_shifts = {
        DR_MINAKSHI_MISHRA  : 4,
        DR_ABHILASHA_MISHRA : 5,
        DR_AMIT_TRIPATHI    : 5,
        DR_SUVARNA_KUMAR    : 5,
        DR_RASHMI_SHARMA    : 5,
        DR_SHARMILA         : 4,
        DR_SAUMYA_SHUKLA    : 4,
        DR_MADHURI_TRIPATHI : 0,
    }
    max_morning_shifts = {
        DR_MINAKSHI_MISHRA  : 8,
        DR_ABHILASHA_MISHRA : 0,
        DR_AMIT_TRIPATHI    : 7,
        DR_SUVARNA_KUMAR    : 0,
        DR_RASHMI_SHARMA    : 8,
        DR_SHARMILA         : 8,
        DR_SAUMYA_SHUKLA    : 0,
        DR_MADHURI_TRIPATHI : 1,
    }
    first_night_off = DR_ABHILASHA_MISHRA
    print("Generating schedule...")
    solution_maybe = generate_schedule(
        doctors=all_doctors,
        year=year,
        month=month,
        fixed_shifts=fixed_shifts,
        unavailable_shifts=unavailable_shifts,
        first_night_off=first_night_off,
        max_night_shifts=max_night_shifts,
        max_morning_shifts=max_morning_shifts,
        sat_ot_duty_rotation_size=4,
    )
    if solution_maybe is not None:
        df_schedule, df_stats = solution_maybe
        excel_output = f"schedule_{year}_{month}.xlsx"
        print(f"One solution found!, writing it to {excel_output}")
        with ExcelWriter(excel_output) as xlw:
            df_schedule.to_excel(index=False, excel_writer=xlw, sheet_name="schedule")
            df_stats.to_excel(index=False, excel_writer=xlw, sheet_name="stats")
    else:
        print("No solutions found. Try adjusting the constraints")


if __name__ == "__main__":
    main()
