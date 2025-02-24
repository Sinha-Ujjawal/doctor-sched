import pandas as pd
from pandas import ExcelWriter
from doctor_schedule import generate_schedule, generate_month_dates, weeks, all_shifts
from doctors import *


def main():
    year = 2025
    month = 3
    all_doctors = [
        DR_MINAKSHI_MISHRA,
        DR_ABHILASHA_MISHRA,
        DR_AMIT_TRIPATHI,
        DR_RASHMI_SHARMA,
        DR_SAUMYA_SHUKLA,
        DR_MADHURI_TRIPATHI,
        DR_KRITIKA_PRASAD,
        DR_ASHOK_KUMAR,
    ]
    dates = generate_month_dates(year, month)
    fixed_shifts = {}
    unavailable_shifts = {}
    leaves = [
        (DR_ASHOK_KUMAR, list(range(1, 12))),
        (DR_SAUMYA_SHUKLA, (6, 7, 8, 9, 12, 13, 14, 15, 16)),
        (DR_MINAKSHI_MISHRA, (3, 4, 5, 6, 7, 16, 17)),
        (DR_RASHMI_SHARMA, (5, 13, 14, 15, 16)),
    ]
    for dt in dates:
        day = dt.day
        unavailable_shifts[(DR_RASHMI_SHARMA    , day)] = ["ot_duty"]
        unavailable_shifts[(DR_MINAKSHI_MISHRA  , day)] = ["ot_duty"]
        unavailable_shifts[(DR_KRITIKA_PRASAD   , day)] = ["ot_duty"]
        unavailable_shifts[(DR_MADHURI_TRIPATHI , day)] = ["ot_duty"]
        if dt.weekday() == weeks.index("Mon"):
            unavailable_shifts[(DR_ABHILASHA_MISHRA, day)] = ["ot_duty"]
        if dt.weekday() == weeks.index("Sun"):
            unavailable_shifts[(DR_MADHURI_TRIPATHI, day)] = ["ot_duty", "night"]
            unavailable_shifts[(DR_AMIT_TRIPATHI   , day)] = ["morning", "evening"]
            unavailable_shifts[(DR_ABHILASHA_MISHRA, day)] = ["morning", "evening"]
            unavailable_shifts[(DR_SAUMYA_SHUKLA   , day)] = ["morning", "evening"]
            unavailable_shifts[(DR_ASHOK_KUMAR     , day)] = ["morning", "evening"]
        # else:
        #     unavailable_shifts[(DR_MADHURI_TRIPATHI, day)] = ["ot_duty", "morning", "night"]
        if dt.weekday() == weeks.index("Wed"):
            unavailable_shifts[(DR_ABHILASHA_MISHRA, day)] = ["night"]
    for doctor, days in leaves:
        for day in days:
            unavailable_shifts[(doctor, day)] = all_shifts
    for dt in dates:
        day = dt.day
        week = dt.weekday()
        if week in (weeks.index("Thu"),):
            fixed_shifts[(DR_ABHILASHA_MISHRA, day)] = ["ot_duty"]
        if week in (weeks.index("Tue"),):
            fixed_shifts[(DR_AMIT_TRIPATHI, day)] = ["ot_duty"]
        if week in (weeks.index("Fri"),):
            fixed_shifts[(DR_SAUMYA_SHUKLA, day)] = ["ot_duty"]
        if week != weeks.index("Sun"):
            fixed_shifts[(DR_MADHURI_TRIPATHI, day)] = ["evening"]
        if day == 14:
            fixed_shifts[(DR_ABHILASHA_MISHRA, day)] = ["ot_duty"]
        # if day in (15, 16, 17, 24, 31):
        #     fixed_shifts[(DR_ASHOK_KUMAR, day)] = ["ot_duty"]

    def my_custom_constraints(model, shift_vars):
        # DR_ABHILASHA_MISHRA will do ot-duty with night on same day
        # DR_MADHURI will do both morning and evening on 14th Mar 2025
        DR_ABHILASHA_MISHRA_INDEX = all_doctors.index(DR_ABHILASHA_MISHRA)
        DR_MADHURI_TRIPATHI_INDEX = all_doctors.index(DR_MADHURI_TRIPATHI)
        # model.Add(sum(
        #     shift_vars[(DR_ABHILASHA_MISHRA_INDEX, dt.day, "night")]
        #     for dt in dates
        #     if dt.weekday() == weeks.index("Thu")
        # ) >= 1)
        for dt in dates:
            week = dt.weekday()
            day = dt.day
            if week in (weeks.index("Tue"), weeks.index("Fri")):
                model.Add(shift_vars[(DR_ABHILASHA_MISHRA_INDEX, dt.day, "night")] == 1).OnlyEnforceIf(
                    shift_vars[(DR_ABHILASHA_MISHRA_INDEX, dt.day, "ot_duty")]
                )
                model.Add(shift_vars[(DR_ABHILASHA_MISHRA_INDEX, dt.day, "morning")] == 0).OnlyEnforceIf(
                    shift_vars[(DR_ABHILASHA_MISHRA_INDEX, dt.day, "ot_duty")]
                )
                model.Add(shift_vars[(DR_ABHILASHA_MISHRA_INDEX, dt.day, "evening")] == 0).OnlyEnforceIf(
                    shift_vars[(DR_ABHILASHA_MISHRA_INDEX, dt.day, "ot_duty")]
                )
            elif week != weeks.index("Sun"):
                model.Add(sum(shift_vars[(DR_ABHILASHA_MISHRA_INDEX, dt.day, shift)] for shift in all_shifts) <= 1)
            if dt.day == 14:
                model.Add(shift_vars[DR_MADHURI_TRIPATHI_INDEX, dt.day, "morning"] == 1)
                model.Add(shift_vars[DR_MADHURI_TRIPATHI_INDEX, dt.day, "evening"] == 1)
            elif week != weeks.index("Sun"):
                model.Add(sum(shift_vars[(DR_MADHURI_TRIPATHI_INDEX, dt.day, shift)] for shift in all_shifts) <= 1)

    max_night_shifts = {
        DR_MINAKSHI_MISHRA  : 5,
        DR_ABHILASHA_MISHRA : 7,
        DR_AMIT_TRIPATHI    : 5,
        DR_RASHMI_SHARMA    : 3,
        DR_SAUMYA_SHUKLA    : 5,
        DR_KRITIKA_PRASAD   : 5,
        DR_ASHOK_KUMAR      : 3,
        DR_MADHURI_TRIPATHI : 0,
    }
    max_morning_shifts = {
        DR_MINAKSHI_MISHRA  : 10,
        DR_RASHMI_SHARMA    : 10,
        DR_KRITIKA_PRASAD   : 10,
        DR_MADHURI_TRIPATHI : 2,
        DR_AMIT_TRIPATHI    : 0,
        DR_SAUMYA_SHUKLA    : 0,
        DR_ABHILASHA_MISHRA : 0,
        DR_ASHOK_KUMAR      : 0,
    }
    minmax_ot_duty_shifts = {
        # DR_ASHOK_KUMAR      : (6, 6),
        DR_ABHILASHA_MISHRA : (9, 9),
        DR_AMIT_TRIPATHI    : (9, 9),
        DR_SAUMYA_SHUKLA    : (8, 8),
    }
    avoid_shift_collision = []
    for dt in dates:
        day = dt.day
        week = dt.weekday()
        avoid_shift_collision.extend([
            (
                DR_SAUMYA_SHUKLA,
                "ot_duty",
                day,
                DR_ABHILASHA_MISHRA,
                "morning",
                day,
            ),
            (
                DR_SAUMYA_SHUKLA,
                "ot_duty",
                day,
                DR_ABHILASHA_MISHRA,
                "evening",
                day,
            ),
            (
                DR_SAUMYA_SHUKLA,
                "ot_duty",
                day,
                DR_ABHILASHA_MISHRA,
                "night",
                day,
            ),
            (
                DR_ABHILASHA_MISHRA,
                "ot_duty",
                day,
                DR_SAUMYA_SHUKLA,
                "morning",
                day,
            ),
            (
                DR_ABHILASHA_MISHRA,
                "ot_duty",
                day,
                DR_SAUMYA_SHUKLA,
                "evening",
                day,
            ),
            (
                DR_ABHILASHA_MISHRA,
                "ot_duty",
                day,
                DR_SAUMYA_SHUKLA,
                "night",
                day,
            ),
        ])
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
    first_night_off = DR_AMIT_TRIPATHI
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
        wed_ot_duty_rotation_size=4,
        sat_ot_duty_rotation_size=None,
        sun_ot_duty_rotation_size=4,
        same_sat_and_sun_ot_duty=True,
        avoid_shift_collision=avoid_shift_collision,
        custom_constraints=my_custom_constraints,
        doctors_who_wants_do_more_shifts_per_day=[DR_ABHILASHA_MISHRA, DR_MADHURI_TRIPATHI],
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
