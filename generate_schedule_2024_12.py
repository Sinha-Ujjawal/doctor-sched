import pandas as pd
from pandas import ExcelWriter
from doctor_schedule import generate_schedule, generate_month_dates, weeks, all_shifts
from doctors import *


def main():
    year = 2024
    month = 12
    all_doctors = [
        DR_MINAKSHI_MISHRA,
        DR_ABHILASHA_MISHRA,
        DR_AMIT_TRIPATHI,
        DR_RASHMI_SHARMA,
        DR_SHARMILA,
        DR_SAUMYA_SHUKLA,
        DR_MADHURI_TRIPATHI,
    ]
    dates = generate_month_dates(year, month)
    fixed_shifts = {}
    unavailable_shifts = {}
    leaves = [
        (DR_MINAKSHI_MISHRA , 1),
        (DR_MINAKSHI_MISHRA , 2),
        (DR_MINAKSHI_MISHRA , 3),
        (DR_MINAKSHI_MISHRA , 31),

        (DR_RASHMI_SHARMA   , 25),
        (DR_RASHMI_SHARMA   , 26),
        (DR_RASHMI_SHARMA   , 27),
        (DR_RASHMI_SHARMA   , 28),
        (DR_RASHMI_SHARMA   , 29),
        (DR_RASHMI_SHARMA   , 30),

        (DR_MADHURI_TRIPATHI, 13),
        (DR_MADHURI_TRIPATHI, 14),
        (DR_MADHURI_TRIPATHI, 15),

        (DR_AMIT_TRIPATHI   , 13),
        (DR_AMIT_TRIPATHI   , 14),
        (DR_AMIT_TRIPATHI   , 15),

        (DR_SAUMYA_SHUKLA   , 29),
        (DR_SAUMYA_SHUKLA   , 30),
        (DR_SAUMYA_SHUKLA   , 31),
    ]
    for dt in dates:
        day = dt.day
        unavailable_shifts[(DR_RASHMI_SHARMA  , day)] = ["ot_duty"]
        unavailable_shifts[(DR_SHARMILA       , day)] = ["ot_duty"]
        unavailable_shifts[(DR_MINAKSHI_MISHRA, day)] = ["ot_duty"]
        if dt.weekday() == weeks.index("Sun"):
            unavailable_shifts[(DR_MADHURI_TRIPATHI, day)] = ["ot_duty", "night"]
            unavailable_shifts[(DR_AMIT_TRIPATHI   , day)] = ["morning", "evening"]
            unavailable_shifts[(DR_ABHILASHA_MISHRA, day)] = ["morning", "evening"]
            unavailable_shifts[(DR_SAUMYA_SHUKLA   , day)] = ["morning", "evening"]
        else:
            unavailable_shifts[(DR_MADHURI_TRIPATHI, day)] = ["ot_duty", "morning", "night"]
        if dt.weekday() == weeks.index("Wed"):
            unavailable_shifts[(DR_ABHILASHA_MISHRA, day)] = ["night"]
    for doctor, day in leaves:
        unavailable_shifts[(doctor, day)] = all_shifts
    for dt in dates:
        day = dt.day
        week = dt.weekday()
        if week in (weeks.index("Mon"), weeks.index("Thu")):
            fixed_shifts[(DR_ABHILASHA_MISHRA, day)] = ["ot_duty"]
        if week in (weeks.index("Tue"), weeks.index("Sat")):
            fixed_shifts[(DR_AMIT_TRIPATHI, day)] = ["ot_duty"]
        if week in (weeks.index("Wed"), weeks.index("Fri")):
            fixed_shifts[(DR_SAUMYA_SHUKLA, day)] = ["ot_duty"]
        if week != weeks.index("Sun"):
            fixed_shifts[(DR_MADHURI_TRIPATHI, day)] = ["evening"]

    max_night_shifts = {
        DR_MINAKSHI_MISHRA  : 5,
        DR_ABHILASHA_MISHRA : 5,
        DR_AMIT_TRIPATHI    : 5,
        DR_RASHMI_SHARMA    : 6,
        DR_SHARMILA         : 5,
        DR_SAUMYA_SHUKLA    : 5,
        DR_MADHURI_TRIPATHI : 0,
    }
    max_morning_shifts = {
        # DR_ABHILASHA_MISHRA : 0,
        # DR_SAUMYA_SHUKLA    : 1,
        # DR_MINAKSHI_MISHRA  : 9,
        # DR_AMIT_TRIPATHI    : 9,
        # DR_RASHMI_SHARMA    : 9,
        # DR_SHARMILA         : 9,
        # DR_MADHURI_TRIPATHI : 9,
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
        sat_ot_duty_rotation_size=None,
        sun_ot_duty_rotation_size=3,
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
