from typing import List, Optional, Dict, Tuple, Callable
from datetime import date, timedelta
from collections import deque
from itertools import islice

from ortools.sat.python import cp_model
import pandas as pd


Doctor = str
Shift = str
Day = int


def clamp(x, lb, ub):
    "Clamps a value (x) between [lb..ub]"
    return min(max(x, lb), ub)


def sliding_window(iterable, n):
    "Collect data into overlapping fixed-length chunks or blocks."
    # sliding_window('ABCDEFG', 4) → ABCD BCDE CDEF DEFG
    iterator = iter(iterable)
    window = deque(islice(iterator, n - 1), maxlen=n)
    for x in iterator:
        window.append(x)
        yield tuple(window)


def generate_month_dates(year: int, month: int) -> List[date]:
    "Generate a list of all dates in a given month in yyyy-mm-dd format"
    next_month = month + 1 if month < 12 else 1
    next_month_year = year if month < 12 else year + 1
    num_dates = (date(next_month_year, next_month, 1) - timedelta(days=1)).day
    return [date(year=year, month=month, day=day) for day in range(1, num_dates + 1)]


weeks = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
all_shifts = ["ot_duty", "morning", "evening", "night"]


def generate_schedule(
    *,
    doctors: List[Doctor],
    year: int,
    month: int,
    fixed_shifts: Dict[Tuple[Doctor, Day], List[Shift]],
    unavailable_shifts: Dict[Tuple[Doctor, Day], List[Shift]],
    first_night_off: Optional[Doctor],
    max_night_shifts: Dict[Doctor, int] = {},
    max_evening_shifts: Dict[Doctor, int] = {},
    max_morning_shifts: Dict[Doctor, int] = {},
    minmax_ot_duty_shifts: Dict[Doctor, Tuple[int, int]] = {},
    sat_ot_duty_rotation_size: Optional[int] = -1,
    sun_ot_duty_rotation_size: Optional[int] = -1,
    sun_morning_evening_duty_rotation_size: Optional[int] = None,
    wed_ot_duty_rotation_size: Optional[int] = -1,
    same_sat_and_sun_ot_duty: bool = False,
    avoid_shift_collision: List[Tuple[Doctor, Shift, Day, Doctor, Shift, Day]] = [],
    custom_constraints: Optional[Callable] = None,
    doctors_who_wants_do_more_shifts_per_day: Optional[List[Doctor]] = None,
    seed: int = 0,
) -> Optional[Tuple[pd.DataFrame, pd.DataFrame]]:
    # removing unavailable shifts from fixed_shifts
    fixed_shifts = {**fixed_shifts}
    for (doctor, day), shifts in unavailable_shifts.items():
        if (doctor, day) in fixed_shifts:
            updated_shifts = list(set(fixed_shifts[(doctor, day)]) - set(shifts))
            if updated_shifts:
                fixed_shifts[(doctor, day)] = updated_shifts
            else:
                del fixed_shifts[(doctor, day)]

    num_doctors = len(doctors)
    dates = generate_month_dates(year, month)

    week_to_days = {
        week: [dt.day for dt in dates if dt.weekday() == week_num]
        for week_num, week in enumerate(weeks)
    }
    if sat_ot_duty_rotation_size is not None:
        if sat_ot_duty_rotation_size == -1:
            sat_ot_duty_rotation_size = len(week_to_days["Sat"])
        sat_ot_duty_rotation_size = clamp(
            sat_ot_duty_rotation_size, 2, len(week_to_days["Sat"])
        )
    if sun_ot_duty_rotation_size is not None:
        if sun_ot_duty_rotation_size == -1:
            sun_ot_duty_rotation_size = len(week_to_days["Sun"])
        sun_ot_duty_rotation_size = clamp(
            sun_ot_duty_rotation_size, 2, len(week_to_days["Sun"])
        )
    if wed_ot_duty_rotation_size is not None:
        if wed_ot_duty_rotation_size == -1:
            wed_ot_duty_rotation_size = len(week_to_days["Wed"])
        wed_ot_duty_rotation_size = clamp(
            wed_ot_duty_rotation_size, 2, len(week_to_days["Wed"])
        )
    if sun_morning_evening_duty_rotation_size is not None:
        if sun_morning_evening_duty_rotation_size == -1:
            sun_morning_evening_duty_rotation_size = len(week_to_days["Sun"])
        sun_morning_evening_duty_rotation_size = clamp(
            sun_morning_evening_duty_rotation_size, 2, len(week_to_days["Sun"])
        )

    # Create the model
    model = cp_model.CpModel()

    # Variables
    shift_vars: Dict[Tuple[int, int, Shift], cp_model.IntVar] = {}

    # Create variables for each doctor, each day, and each shift
    for d in range(num_doctors):
        for dt in dates:
            day = dt.day
            for shift in all_shifts:
                shift_vars[(d, day, shift)] = model.NewBoolVar(
                    f"doctor_{d}_day_{day}_{shift}"
                )

    # Constraints
    # First night off
    if first_night_off is not None:
        fd = doctors.index(first_night_off)
        model.Add(sum(shift_vars[(fd, 1, shift)] for shift in all_shifts) < 1)

    monday_fixed_shifts = {
        (doctors.index(doctor), day)
        for (doctor, day), _ in fixed_shifts.items()
        if day in week_to_days["Mon"]
    }

    # Each doctor, except for those who choses can only work one shift per day (doctors_who_wants_do_more_shifts_per_day), except for Sunday shift constraints
    for dt in dates:
        day = dt.day
        next_day = day + 1
        if dt.weekday() == weeks.index("Sun"):
            for d in range(num_doctors):
                if (d, next_day) not in monday_fixed_shifts:
                    model.Add(
                        shift_vars[(d, day, "ot_duty")] == shift_vars[(d, day, "night")]
                    ).OnlyEnforceIf(shift_vars[(d, day, "ot_duty")])
                model.Add(
                    shift_vars[(d, day, "morning")] == shift_vars[(d, day, "evening")]
                ).OnlyEnforceIf(shift_vars[(d, day, "morning")])
                model.Add(sum(shift_vars[(d, day, shift)] for shift in all_shifts) <= 2)
        else:
            for d in range(num_doctors):
                if doctors[d] not in (doctors_who_wants_do_more_shifts_per_day or []):
                    model.Add(sum(shift_vars[(d, day, shift)] for shift in all_shifts) <= 1)

    # All shifts must have exactly one doctor assigned
    for dt in dates:
        day = dt.day
        for shift in all_shifts:
            model.Add(sum(shift_vars[(d, day, shift)] for d in range(num_doctors)) >= 1)

    # Fixed shifts
    for (doctor, day), shifts in fixed_shifts.items():
        d = doctors.index(doctor)
        for shift in shifts:
            model.Add(shift_vars[(d, day, shift)] == 1)

    # Unavailable shifts
    for (doctor, day), shifts in unavailable_shifts.items():
        d = doctors.index(doctor)
        for shift in shifts:
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

    # Otduty shift constraints
    for doctor, (min_ot_duty_shift, max_ot_duty_shift) in minmax_ot_duty_shifts.items():
        d = doctors.index(doctor)
        # Sum the ot_duty shifts assigned to this doctor
        sum_ot_duties = sum(shift_vars[(d, dt.day, "ot_duty")] for dt in dates)
        model.Add(
            sum_ot_duties >= min_ot_duty_shift
        )
        model.Add(
            sum_ot_duties <= max_ot_duty_shift
        )

    # Evening shift constraints
    evening_shifts_count = {}
    for doctor, max_evening_shift in max_evening_shifts.items():
        d = doctors.index(doctor)
        evening_shifts_count[d] = model.NewIntVar(
            0, max_evening_shift, f"count_evening_shifts_{d}"
        )
        # Sum the evening shifts assigned to this doctor
        model.Add(
            evening_shifts_count[d]
            == sum(shift_vars[(d, dt.day, "evening")] for dt in dates)
        )

    # Morning shift constraints
    morning_shifts_count = {}
    for doctor, max_morning_shift in max_morning_shifts.items():
        d = doctors.index(doctor)
        morning_shifts_count[d] = model.NewIntVar(
            0, max_morning_shift, f"count_morning_shifts_{d}"
        )
        # Sum the morning shifts assigned to this doctor
        model.Add(
            morning_shifts_count[d]
            == sum(shift_vars[(d, dt.day, "morning")] for dt in dates)
        )

    # Night off logic: if a doctor works night shift, they cannot work the next day
    for d in range(num_doctors):
        for dt in dates[:-1]:
            day = dt.day
            model.Add(
                sum(shift_vars[(d, day + 1, shift)] for shift in all_shifts) == 0
            ).OnlyEnforceIf(shift_vars[(d, day, "night")])

    # if OT no emergency next day: if a doctor works ot duty, then no morning emergency next day
    for d in range(num_doctors):
        for dt in dates[:-1]:
            day = dt.day
            model.Add(shift_vars[(d, day + 1, "morning")] == 0).OnlyEnforceIf(
                shift_vars[(d, day, "ot_duty")]
            )

    # OT Duty on Saturday and Sundays on rotation
    for d in range(num_doctors):
        if wed_ot_duty_rotation_size is not None:
            for window in sliding_window(week_to_days["Wed"], wed_ot_duty_rotation_size):
                model.Add(sum(shift_vars[(d, day, "ot_duty")] for day in window) <= 1)
        if sat_ot_duty_rotation_size is not None:
            for window in sliding_window(week_to_days["Sat"], sat_ot_duty_rotation_size):
                model.Add(sum(shift_vars[(d, day, "morning")] for day in window) <= 1)
                model.Add(sum(shift_vars[(d, day, shift)]
                              for day in window
                              for shift in ["ot_duty", "night"]) <= 1)
        if sun_ot_duty_rotation_size is not None:
            for window in sliding_window(week_to_days["Sun"], sun_ot_duty_rotation_size):
                model.Add(sum(shift_vars[(d, day, "ot_duty")] for day in window) <= 1)
                model.Add(sum(shift_vars[(d, day, "morning")] for day in window) <= 1)
                model.Add(sum(shift_vars[(d, day, "evening")] for day in window) <= 1)
                model.Add(sum(shift_vars[(d, day, "night")]   for day in window) <= 1)
    
    if same_sat_and_sun_ot_duty:
        for d in range(num_doctors):
            for sat in week_to_days["Sat"]:
                sun = sat + 1
                if sun not in week_to_days["Sun"]: continue
                model.Add(shift_vars[(d, sat, "ot_duty")] == 1).OnlyEnforceIf(
                    shift_vars[(d, sun, "ot_duty")]
                )
                model.Add(shift_vars[(d, sun, "ot_duty")] == 1).OnlyEnforceIf(
                    shift_vars[(d, sat, "ot_duty")]
                )

    # Morning and Evening duty on Sundays on rotation
    for d in range(num_doctors):
        if sun_morning_evening_duty_rotation_size is not None:
            for window in sliding_window(week_to_days["Sun"], sun_morning_evening_duty_rotation_size):
                model.Add(sum(shift_vars[(d, day, "morning")] for day in window) <= 1)
                model.Add(sum(shift_vars[(d, day, "evening")] for day in window) <= 1)

    for avoid_shift_spec in avoid_shift_collision:
        # print(f"Avoiding shift collision for {avoid_shift_spec}")
        doc1, shift1, day1, doc2, shift2, day2 = avoid_shift_spec
        d1 = doctors.index(doc1)
        d2 = doctors.index(doc2)
        model.Add(shift_vars[(d2, day2, shift2)] == 0).OnlyEnforceIf(
            shift_vars[(d1, day1, shift1)]
        )

    if custom_constraints:
        custom_constraints(model=model, shift_vars=shift_vars)

    # Solve the model
    solver = cp_model.CpSolver()
    solver.parameters.random_seed = seed
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
            week_num = (day - 1) % 7
            week_day = dt.weekday()
            schedule.append(
                [
                    day,
                    dt,
                    week_num,
                    weeks[week_day],
                    *[find_doctor(day, shift) for shift in all_shifts],
                ]
            )
        df_schedule = pd.DataFrame(
            schedule, columns=["day", "date", "week_num", "week", *all_shifts]
        )
        df_schedule.sort_values(["day"], inplace=True)
        df_schedule["night_off"] = df_schedule["night"].shift(1)
        df_schedule.loc[df_schedule["day"] == 1, "night_off"] = first_night_off
        # df_schedule.sort_values(["week_num", "date"], inplace=True)
        df_schedule.sort_values(["date"], inplace=True)
        df_schedule.drop(columns=["week_num", "day"], inplace=True)
        df_stats = pd.DataFrame(
            [
                [
                    doctor,
                    *[(df_schedule[shift] == doctor).sum() for shift in all_shifts],
                    "|".join(map(str, sorted({
                        (dt.day, weeks[dt.weekday()], shift)
                        for shift in all_shifts
                        for dt in df_schedule[df_schedule[shift] == doctor]["date"]
                    })))
                ]
                for doctor in doctors
            ],
            columns=["doctor", *all_shifts, "working_days"],
        )
        df_stats["total_shifts"] = sum(df_stats[shift] for shift in all_shifts)
        return df_schedule, df_stats
    return None
