import pandas as pd
from pandas import ExcelWriter
from doctor_schedule import generate_schedule, generate_month_dates, weeks, all_shifts
from doctors import *


def main():
    year = 2025
    month = 2
    all_doctors = [
        DR_MINAKSHI_MISHRA,
        DR_ABHILASHA_MISHRA,
        DR_AMIT_TRIPATHI,
        DR_RASHMI_SHARMA,
        DR_SAUMYA_SHUKLA,
        DR_MADHURI_TRIPATHI,
        DR_KRITIKA_PRASAD,
    ]
    dates = generate_month_dates(year, month)
    fixed_shifts = {}
    unavailable_shifts = {}
    leaves = [
        (DR_RASHMI_SHARMA,    (2, 3, 6, 7, 8, 9, 10, 14, 15, 16)),
        (DR_SAUMYA_SHUKLA,    (20,)),
        # (DR_AMIT_TRIPATHI,    (3, 13, 14, 15)),
        (DR_AMIT_TRIPATHI,    (3, 13, 14)),
        (DR_MADHURI_TRIPATHI, (3, 13, 14, 15, 16, 17, 18)),
    ]
    for dt in dates:
        day = dt.day
        unavailable_shifts[(DR_RASHMI_SHARMA  , day)] = ["ot_duty"]
        unavailable_shifts[(DR_MINAKSHI_MISHRA, day)] = ["ot_duty"]
        unavailable_shifts[(DR_KRITIKA_PRASAD   , day)] = ["ot_duty"]
        if dt.weekday() == weeks.index("Sun"):
            unavailable_shifts[(DR_MADHURI_TRIPATHI, day)] = ["ot_duty", "night"]
            unavailable_shifts[(DR_AMIT_TRIPATHI   , day)] = ["morning", "evening"]
            unavailable_shifts[(DR_ABHILASHA_MISHRA, day)] = ["morning", "evening"]
            unavailable_shifts[(DR_SAUMYA_SHUKLA   , day)] = ["morning", "evening"]
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
        if week in (weeks.index("Mon"),):
            fixed_shifts[(DR_ABHILASHA_MISHRA, day)] = ["ot_duty"]
        if week in (weeks.index("Thu"),) and day != 27:
            fixed_shifts[(DR_ABHILASHA_MISHRA, day)] = ["ot_duty"]
        if week in (weeks.index("Tue"),):
            fixed_shifts[(DR_AMIT_TRIPATHI, day)] = ["ot_duty"]
        if week in (weeks.index("Fri"),):
            fixed_shifts[(DR_SAUMYA_SHUKLA, day)] = ["ot_duty"]
        # if week != weeks.index("Sun"):
        #     fixed_shifts[(DR_MADHURI_TRIPATHI, day)] = ["evening"]
        # if day in (4, 12):
        #     fixed_shifts[(DR_RASHMI_SHARMA, day)] = ["night"]
        if day in (22, 23):
            fixed_shifts[(DR_AMIT_TRIPATHI, day)] = ["ot_duty"]

    def my_custom_constraints(model, shift_vars):
        # For days != Sunday
        # DR_MADHURI_TRIPATHI will do morning emergency with DR_AMIT_TRIPATHI (ot_duty)
        # Other days, she will only do evening
        DR_MADHURI_TRIPATHI_INDEX = all_doctors.index(DR_MADHURI_TRIPATHI)
        DR_AMIT_TRIPATHI_INDEX = all_doctors.index(DR_AMIT_TRIPATHI)
        for dt in dates:
            day = dt.weekday()
            if day == weeks.index("Sun"):
                continue
            if (DR_MADHURI_TRIPATHI, dt.day) in unavailable_shifts:
                continue
            # If DR_AMIT_TRIPATHI doing ot_duty
            model.Add(shift_vars[(DR_MADHURI_TRIPATHI_INDEX, dt.day, "morning")] == 1).OnlyEnforceIf(
                shift_vars[(DR_AMIT_TRIPATHI_INDEX, dt.day, "ot_duty")]
            )
            model.Add(shift_vars[(DR_MADHURI_TRIPATHI_INDEX, dt.day, "evening")] == 0).OnlyEnforceIf(
                shift_vars[(DR_AMIT_TRIPATHI_INDEX, dt.day, "ot_duty")]
            )
            model.Add(shift_vars[(DR_MADHURI_TRIPATHI_INDEX, dt.day, "night")] == 0).OnlyEnforceIf(
                shift_vars[(DR_AMIT_TRIPATHI_INDEX, dt.day, "ot_duty")]
            )
            # otherwise,
            model.Add(shift_vars[(DR_MADHURI_TRIPATHI_INDEX, dt.day, "ot_duty")] == 0).OnlyEnforceIf(
                ~shift_vars[(DR_AMIT_TRIPATHI_INDEX, dt.day, "ot_duty")]
            )
            model.Add(shift_vars[(DR_MADHURI_TRIPATHI_INDEX, dt.day, "morning")] == 0).OnlyEnforceIf(
                ~shift_vars[(DR_AMIT_TRIPATHI_INDEX, dt.day, "ot_duty")]
            )
            model.Add(shift_vars[(DR_MADHURI_TRIPATHI_INDEX, dt.day, "evening")] == 1).OnlyEnforceIf(
                ~shift_vars[(DR_AMIT_TRIPATHI_INDEX, dt.day, "ot_duty")]
            )
            model.Add(shift_vars[(DR_MADHURI_TRIPATHI_INDEX, dt.day, "night")] == 0).OnlyEnforceIf(
                ~shift_vars[(DR_AMIT_TRIPATHI_INDEX, dt.day, "ot_duty")]
            )

    max_night_shifts = {
        DR_MINAKSHI_MISHRA  : 5,
        DR_ABHILASHA_MISHRA : 5,
        DR_AMIT_TRIPATHI    : 5,
        DR_RASHMI_SHARMA    : 5,
        DR_SAUMYA_SHUKLA    : 4,
        DR_KRITIKA_PRASAD     : 4,
        DR_MADHURI_TRIPATHI : 0,
    }
    max_morning_shifts = {
        DR_MINAKSHI_MISHRA  : 7,
        DR_KRITIKA_PRASAD     : 8,
        DR_SAUMYA_SHUKLA    : 0,
        DR_ABHILASHA_MISHRA : 0,
    }
    minmax_ot_duty_shifts = {
        DR_ABHILASHA_MISHRA : (9, 9),
        DR_AMIT_TRIPATHI    : (9, 10),
        DR_SAUMYA_SHUKLA    : (9, 9),
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
        sat_ot_duty_rotation_size=None,
        sun_ot_duty_rotation_size=3,
        same_sat_and_sun_ot_duty=True,
        avoid_shift_collision=avoid_shift_collision,
        custom_constraints=my_custom_constraints,
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
