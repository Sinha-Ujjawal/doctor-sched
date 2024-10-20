from pandas import ExcelWriter
from doctor_schedule import generate_schedule, generate_month_dates, weeks


def main():
    year = 2024
    month = 10
    doctors = [
        "Dr. Minakshi Mishra",
        "Dr. Abhilasha Mishra",
        "Dr. Amit Tripathi",
        "Dr. Suvarna Kumar",
        "Dr. Rashmi Sharma",
        "Dr. Sharmila",
        "Dr. Saumya Shukla",
        "Dr. Madhuri Tripathi",
    ]
    dates = generate_month_dates(year, month)
    fixed_shifts = {}
    unavailable_shifts = {}
    for dt in dates:
        day = dt.day
        week = dt.weekday()

        if week == weeks.index("Mon"):
            fixed_shifts[("Dr. Abhilasha Mishra", day)] = "ot_duty"
        if week == weeks.index("Tue"):
            fixed_shifts[("Dr. Amit Tripathi", day)] = "ot_duty"
        if week == weeks.index("Wed"):
            fixed_shifts[("Dr. Amit Tripathi", day)] = "ot_duty"
        if week == weeks.index("Thu"):
            fixed_shifts[("Dr. Saumya Shukla", day)] = "ot_duty"
        if week == weeks.index("Fri"):
            fixed_shifts[("Dr. Suvarna Kumar", day)] = "ot_duty"

        if week == weeks.index("Tue") and day != 29:
            fixed_shifts[("Dr. Suvarna Kumar", day)] = "night"
        if week != weeks.index("Sun"):
            fixed_shifts[("Dr. Madhuri Tripathi", day)] = "evening"

        unavailable_shifts[("Dr. Rashmi Sharma", day)] = "ot_duty"
        unavailable_shifts[("Dr. Madhuri Tripathi", day)] = "ot_duty"
        unavailable_shifts[("Dr. Sharmila", day)] = "ot_duty"
        unavailable_shifts[("Dr. Minakshi Mishra", day)] = "ot_duty"
        unavailable_shifts[("Dr. Madhuri Tripathi", day)] = "night"
    max_night_shifts = {
        "Dr. Minakshi Mishra": 5,
        "Dr. Abhilasha Mishra": 4,
        "Dr. Amit Tripathi": 4,
        "Dr. Suvarna Kumar": 5,
        "Dr. Rashmi Sharma": 4,
        "Dr. Sharmila": 4,
        "Dr. Saumya Shukla": 5,
        "Dr. Madhuri Tripathi": 0,
    }
    first_night_off = "Dr. Rashmi Sharma"
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
