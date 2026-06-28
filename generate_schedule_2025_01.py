import pandas as pd
from pandas import ExcelWriter
from doctor_schedule import generate_schedule, generate_month_dates, weeks, all_shifts
from doctors import *


def main():
    year = 2025
    month = 1
    all_doctors = [
        DR_MINAKSHI_MISHRA,
        DR_ABHILASHA_MISHRA,
        DR_AMIT_TRIPATHI,
        DR_RASHMI_SHARMA,
        DR_SAUMYA_SHUKLA,
        DR_MADHURI_TRIPATHI,
    ]
    dates = generate_month_dates(year, month)
    fixed_shifts = {}
    unavailable_shifts = {}
    leaves = [
        (DR_MINAKSHI_MISHRA, 1),
        (DR_MINAKSHI_MISHRA, 28),
        (DR_MINAKSHI_MISHRA, 29),
        (DR_MINAKSHI_MISHRA, 30),
    ]
    for dt in dates:
        day = dt.day
        unavailable_shifts[(DR_RASHMI_SHARMA  , day)] = ["ot_duty"]
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
    amit_tripathi_4 = (DR_AMIT_TRIPATHI, 4)
    if amit_tripathi_4 not in unavailable_shifts:
        unavailable_shifts[amit_tripathi_4] = ["ot_duty"]
    else:
        unavailable_shifts[amit_tripathi_4] = unavailable_shifts[amit_tripathi_4].append("ot_duty")
    for doctor, day in leaves:
        unavailable_shifts[(doctor, day)] = all_shifts
    for dt in dates:
        day = dt.day
        week = dt.weekday()
        if week in (weeks.index("Mon"), weeks.index("Thu")):
            fixed_shifts[(DR_ABHILASHA_MISHRA, day)] = ["ot_duty"]
        if week in (weeks.index("Tue"),):
            fixed_shifts[(DR_AMIT_TRIPATHI, day)] = ["ot_duty"]
        if week in (weeks.index("Fri"),):
            fixed_shifts[(DR_SAUMYA_SHUKLA, day)] = ["ot_duty"]
        if week != weeks.index("Sun"):
            fixed_shifts[(DR_MADHURI_TRIPATHI, day)] = ["evening"]
        if day == 1:
            fixed_shifts[(DR_RASHMI_SHARMA, day)] = ["night"]
    # for day in [1, 8, 15, 24]:
    #     amit_day = (DR_AMIT_TRIPATHI, day)
    #     if amit_day not in fixed_shifts:
    #         fixed_shifts[amit_day] = ["ot_duty"]
    #     else:
    #         fixed_shifts[amit_day].append("ot_duty")

    max_night_shifts = {
        DR_MINAKSHI_MISHRA  : 7,
        DR_ABHILASHA_MISHRA : 6,
        DR_AMIT_TRIPATHI    : 6,
        DR_RASHMI_SHARMA    : 6,
        DR_SAUMYA_SHUKLA    : 6,
        DR_MADHURI_TRIPATHI : 0,
    }
    max_morning_shifts = {
        DR_ABHILASHA_MISHRA : 0,
        DR_SAUMYA_SHUKLA    : 0,
        DR_MADHURI_TRIPATHI : 1,
        DR_MINAKSHI_MISHRA  : 12,
        # DR_AMIT_TRIPATHI    : 8,
        DR_RASHMI_SHARMA    : 13,
    }
    minmax_ot_duty_shifts = {
        DR_ABHILASHA_MISHRA : (10, 11),
        DR_AMIT_TRIPATHI    : (10, 11),
        DR_SAUMYA_SHUKLA    : (10, 10),
    }
    avoid_shift_collision = []
    for dt in dates:
        day = dt.day
        week = dt.weekday()
        if week == weeks.index("Sun"):
            avoid_shift_collision.extend([
                (
                    DR_AMIT_TRIPATHI,
                    "ot_duty",
                    day,
                    DR_MADHURI_TRIPATHI,
                    "morning",
                    day,
                ),
                (
                    DR_AMIT_TRIPATHI,
                    "ot_duty",
                    day,
                    DR_MADHURI_TRIPATHI,
                    "evening",
                    day,
                ),
                (
                    DR_AMIT_TRIPATHI,
                    "ot_duty",
                    day,
                    DR_MADHURI_TRIPATHI,
                    "night",
                    day,
                ),
            ])
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
        minmax_ot_duty_shifts=minmax_ot_duty_shifts,
        wed_ot_duty_rotation_size=None,
        sat_ot_duty_rotation_size=2,
        sun_ot_duty_rotation_size=3,
        same_sat_and_sun_ot_duty=True,
        avoid_shift_collision=avoid_shift_collision,
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
