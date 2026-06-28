from pandas import ExcelWriter
from doctor_schedule import generate_schedule, generate_month_dates, weeks
from doctors import *


def main():
    year = 2024
    month = 10
    doctors = [
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
        if week == weeks.index("Fri"):
            fixed_shifts[(DR_SUVARNA_KUMAR, day)] = ["ot_duty"]

        if week == weeks.index("Tue") and day != 29:
            fixed_shifts[(DR_SUVARNA_KUMAR, day)] = ["night"]
        if week != weeks.index("Sun"):
            fixed_shifts[(DR_MADHURI_TRIPATHI, day)] = ["evening"]

        unavailable_shifts[(DR_RASHMI_SHARMA, day)] = ["ot_duty"]
        unavailable_shifts[(DR_MADHURI_TRIPATHI, day)] = ["ot_duty"]
        unavailable_shifts[(DR_SHARMILA, day)] = ["ot_duty"]
        unavailable_shifts[(DR_MINAKSHI_MISHRA, day)] = ["ot_duty"]
        unavailable_shifts[(DR_MADHURI_TRIPATHI, day)] = ["night"]
    max_night_shifts = {
        DR_MINAKSHI_MISHRA: 5,
        DR_ABHILASHA_MISHRA: 4,
        DR_AMIT_TRIPATHI: 4,
        DR_SUVARNA_KUMAR: 5,
        DR_RASHMI_SHARMA: 4,
        DR_SHARMILA: 4,
        DR_SAUMYA_SHUKLA: 5,
        DR_MADHURI_TRIPATHI: 0,
    }
    first_night_off = DR_RASHMI_SHARMA
    print("Generating schedule...")
    solution_maybe = generate_schedule(
        doctors=doctors,
        year=year,
        month=month,
        fixed_shifts=fixed_shifts,
        unavailable_shifts=unavailable_shifts,
        max_night_shifts=max_night_shifts,
        first_night_off=first_night_off,
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
