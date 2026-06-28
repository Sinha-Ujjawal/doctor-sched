## Doctor Schedule 🩺📅
This is an python script to help me generate schedule for doctors at [DWH Gonda](https://gonda.nic.in/public-utility/district-womens-hospital/). The script uses [Google's OR-Tools SAT Solver](https://developers.google.com/optimization/cp/cp_solver) for generating the schedule.

## Prerequisites

1. Install [python>=3.8](https://www.python.org/downloads/release/python-380/)
2. Install dependencies present in [requirements.txt](./requirements.txt)
``` console
python -m pip install -r requirements.txt
```

## Getting started

1. Copy the example script [generate_schedule_2024_10.py](./generate_schedule_2024_10.py) and change the `year`, `month` and other constraints `fixed_shifts` for doctors who have fixed their shifts on a particular date, and `unavailable_shifts` for doctors who are either on leave, or don't want to work a particular shift.
2. Run the script
```console
python <your-script.py>

```
3. This should generate an excel file containing the `schedule` and the `stats` sheet. Modify it to your content and use.

## Copyrights

Licensed under [@MIT](./LICENSE)

