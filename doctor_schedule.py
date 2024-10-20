from typing import List, Optional, Dict, Tuple
from datetime import date, timedelta

from ortools.sat.python import cp_model
import pandas as pd


Doctor = str
Shift = str
Day = int


def generate_month_dates(year: int, month: int) -> List[date]:
    """Generate a list of all dates in a given month in yyyy-mm-dd format"""
    next_month = month+1 if month < 12 else 1
    next_month_year = year if month < 12 else year+1
    num_dates = (date(next_month_year, next_month, 1) - timedelta(days=1)).day
    return [date(year=year, month=month, day=day) for day in range(1, num_dates + 1)]


weeks = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
shifts = ["ot_duty", "morning", "evening", "night"]


def generate_schedule(
    *,
    doctors: List[Doctor],
    year: int,
    month: int,
    fixed_shifts: Dict[Tuple[Doctor, Day], Shift],
    unavailable_shifts: Dict[Tuple[Doctor, Day], Shift],
    max_night_shifts: Dict[Doctor, int],
    first_night_off: Doctor,
) -> Optional[Tuple[pd.DataFrame, pd.DataFrame]]:
    num_doctors = len(doctors)
    dates = generate_month_dates(year, month)

    # Create the model
    model = cp_model.CpModel()

    # Variables
    shift_vars: Dict[Tuple[int, int, Shift], cp_model.IntVar] = {}

    # Create variables for each doctor, each day, and each shift
    for d in range(num_doctors):
        for dt in dates:
            day = dt.day
            for shift in shifts:
                shift_vars[(d, day, shift)] = model.NewBoolVar(
                    f"doctor_{d}_day_{day}_{shift}"
                )

    # Constraints
    # First night off
    fd = doctors.index(first_night_off)
    model.Add(sum(shift_vars[(fd, 1, shift)] for shift in shifts) < 1)

    monday_doctors = {
        doctors.index(doctor)
        for (doctor, day), _ in fixed_shifts.items()
        if dates[day - 1].weekday() == weeks.index("Mon")
    }

    # Each doctor can only work one shift per day, except for Sunday shift constraints
    for dt in dates:
        day = dt.day
        if dt.weekday() == weeks.index("Sun"):
            for d in range(num_doctors):
                if d not in monday_doctors:
                    model.Add(
                        shift_vars[(d, day, "ot_duty")] == shift_vars[(d, day, "night")]
                    ).OnlyEnforceIf(shift_vars[(d, day, "ot_duty")])
                model.Add(
                    shift_vars[(d, day, "morning")] == shift_vars[(d, day, "evening")]
                ).OnlyEnforceIf(shift_vars[(d, day, "morning")])
                model.Add(sum(shift_vars[(d, day, shift)] for shift in shifts) <= 2)
        else:
            for d in range(num_doctors):
                model.Add(sum(shift_vars[(d, day, shift)] for shift in shifts) <= 1)

    # All shifts must have exactly one doctor assigned
    for dt in dates:
        day = dt.day
        for shift in shifts:
            model.Add(sum(shift_vars[(d, day, shift)] for d in range(num_doctors)) >= 1)

    # Fixed shifts
    for (doctor, day), shift in fixed_shifts.items():
        d = doctors.index(doctor)
        model.Add(shift_vars[(d, day, shift)] == 1)

    # Unavailable shifts
    for (doctor, day), shift in unavailable_shifts.items():
        d = doctors.index(doctor)
        model.Add(shift_vars[(d, day, shift)] == 0)

    # Night shift constraints
    night_shifts_count = {}
    for doctor, max_night_shift in max_night_shifts.items():
        d = doctors.index(doctor)
        night_shifts_count[d] = model.NewIntVar(
            0, max_night_shift, f"count_night_shifts_{d}"
        )
        # Sum the night shifts assigned to this doctor
        model.Add(
            night_shifts_count[d]
            == sum(shift_vars[(d, dt.day, "night")] for dt in dates)
        )

    # Night off logic: if a doctor works night shift, they cannot work the next day
    for d in range(num_doctors):
        for dt in dates[:-1]:
            day = dt.day
            model.Add(
                sum(shift_vars[(d, day + 1, shift)] for shift in shifts) == 0
            ).OnlyEnforceIf(shift_vars[(d, day, "night")])

    # if OT no emergency next day: if a doctor works ot duty, then no morning emergency next day
    for d in range(num_doctors):
        for dt in dates[:-1]:
            day = dt.day
            model.Add(shift_vars[(d, day + 1, "morning")] == 0).OnlyEnforceIf(
                shift_vars[(d, day, "ot_duty")]
            )

    # no multiple sundays for a given doctor
    for d in range(num_doctors):
        model.Add(
            sum(
                shift_vars[(d, dt.day, "ot_duty")]
                for dt in dates
                if dt.weekday() == weeks.index("Sun")
            )
            <= 1
        )
        model.Add(
            sum(
                shift_vars[(d, dt.day, "morning")]
                for dt in dates
                if dt.weekday() == weeks.index("Sun")
            )
            <= 1
        )

    # OT Duty and night for Saturday on rotation
    for d in range(num_doctors):
        model.Add(
            sum(
                shift_vars[(d, dt.day, "ot_duty")]
                for dt in dates
                if dt.weekday() == weeks.index("Sat")
            )
            <= 1
        )
        model.Add(
            sum(
                shift_vars[(d, dt.day, "night")]
                for dt in dates
                if dt.weekday() == weeks.index("Sat")
            )
            <= 1
        )

    # Solve the model
    solver = cp_model.CpSolver()
    status = solver.Solve(model)

    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):

        def find_doctor(day, shift):
            for d in range(num_doctors):
                if solver.Value(shift_vars[(d, day, shift)]) >= 1:
                    return doctors[d]
            return None

        schedule = []
        for dt in dates:
            day = dt.day
            week_num = day % 7
            week_day = dt.weekday()
            schedule.append(
                [
                    day,
                    dt,
                    week_num,
                    weeks[week_day],
                    *[find_doctor(day, shift) for shift in shifts],
                ]
            )
        df_schedule = pd.DataFrame(
            schedule, columns=["day", "date", "week_num", "week", *shifts]
        )
        df_schedule.sort_values(["day"], inplace=True)
        df_schedule["night_off"] = df_schedule["night"].shift(1)
        df_schedule.loc[df_schedule["day"] == 0, "night_off"] = first_night_off
        df_schedule.sort_values(["week_num", "date"], inplace=True)
        df_schedule.drop(columns=["week_num", "day"], inplace=True)
        df_stats = pd.DataFrame(
            [
                [doctor, *[(df_schedule[shift] == doctor).sum() for shift in shifts]]
                for doctor in doctors
            ],
            columns=["doctor", *shifts],
        )
        return df_schedule, df_stats
    return None
