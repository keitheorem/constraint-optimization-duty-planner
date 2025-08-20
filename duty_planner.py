import pandas as pd
from datetime import datetime
import calendar
from ortools.sat.python import cp_model
import sys
import math
from collections import defaultdict

# === ACCEPT INPUTS ===
# Ask user for Public Holiday dates which will also mark the eve in the subsequent code
ph_input = input("Enter Public Holiday dates ( / separated, e.g. 5/19): ").strip()
public_holidays = set()
if ph_input:
    try:
        public_holidays = {int(day.strip()) for day in ph_input.split('/') if day.strip().isdigit()}
    except ValueError:
        print("Invalid PH input. Ignoring Public Holidays.")

# User-tweakable scale for decimal precision
# scale = 1000 -> preserves up to 3 decimal places (e.g. 1.5 -> 1500)
SCALE = 1000

# Load data from template Excel: Name | On Leave/Course | Current Score
staff_df = pd.read_excel("Template.xlsx", sheet_name='Sheet1')  # Update excel or sheet name if needed

# Detect frozen names from "On Leave/Course" column
frozen_names = set(
    staff_df.loc[
        staff_df["On Leave/Course"].astype(str).str.strip().str.lower() == "frozen",
        "Name"
    ].str.strip().str.lower()
)

# === ASSIGN POINTS TO DAYS ===
# get current date
now = datetime.now()
start_date = datetime(now.year, now.month, 1)
last_day = calendar.monthrange(now.year, now.month)[1]
end_date = datetime(now.year, now.month, last_day)
days = pd.date_range(start=start_date, end=end_date)

# Ask if last day of month is a Public Holiday Eve -> note it will overwrite if the last day is a weekend
ph_eve_last_day_input = input(f"Is the last day of the month ({last_day}) a Public Holiday Eve? (y/n): ").strip().lower()
last_day_is_ph_eve = (ph_eve_last_day_input == 'y')

# creates a list of format [Full Date/Time,
#  Day (0 = Monday),
#  Points (float before scaling)]
# - Note weekdays Mon-Thu = 1 point, Fri = 1.5, weekends = 2 points
duty_days = []
for d in days:
    wd = d.weekday()  # Monday=0 ... Sunday=6
    if wd >= 5:
        point = 2.0
    elif wd == 4:  # Friday
        point = 1.5
    else:
        point = 1.0

    # Override for Public Holidays
    if d.day in public_holidays:
        point = 2.0  # PH itself
    elif (d + pd.Timedelta(days=1)).day in public_holidays and point < 2.0: # overwrite if weekend is PH eve
        point = 1.5  # PH eve

    # Special override for last day PH eve, only if not weekend
    if d.day == last_day and last_day_is_ph_eve and point < 2.0:
        point = 1.5

    # store float point (will be converted to scaled integer later)
    duty_days.append((d, wd, point))

# Precompute scaled integer points for each day to keep CP-SAT integer-friendly i.e. scale up points to remain integer
duty_days_scaled = []
for (d, wd, p) in duty_days:
    int_p = int(round(p * SCALE))
    duty_days_scaled.append((d, wd, p, int_p))  # keep both float p and int_p for reference

# initialise the OR-Tools model -> note that this is what runs the iterations
model = cp_model.CpModel()
assignments = {} # Note that this is a Boolean Variable e.g. day_1_staff_1
day_names = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']

# === GENERATE TRUTH TABLE ===
# IMPORTANT NOTE INDEX I = DAYS INDEX J = STAFF
# create a "truth table" considering constraints for every possible permutation, and store it into assignments
for i, (day, weekday, point) in enumerate(duty_days):
    for j, staff in staff_df.iterrows():

        # Check if this person is frozen
        if str(staff["Name"]).strip().lower() in frozen_names:
            continue

        hard_constraints = str(staff.get("On Leave/Course", "")).strip().lower()
        hard_days = set()

        allowed = True
        preferred_days = []

        if hard_constraints: # this deals with all the days a person is unable to do duty due to leave or on course
            try:
                tokens = [int(t.strip()) for t in hard_constraints.replace(',', ' ').split() if t.strip().isdigit()]
                hard_days = set(tokens)
            except Exception as e:
                print(f"Error parsing hard constraints for {staff.get('Name')}: {e}", file=sys.stderr)

            # If current day-of-month is in hard constraint days, block it
            if day.day in hard_days:
                allowed = False

        if allowed:
            var = model.NewBoolVar(f"day_{i}_staff_{j}")
            assignments[(i, j)] = var

# === ASSIGN CONSTRAINTS ===
# Iterate through every day and add a new constraint to the model "AddExactlyOne"
for i in range(len(duty_days)):
    # ensure we only add the constraint if at least one staff is allowed that day
    allowed_vars = [assignments[(i, j)] for j in range(len(staff_df)) if (i, j) in assignments]
    if allowed_vars:
        model.AddExactlyOne(allowed_vars)
    else:
        # No one is available that day (all blocked by hard constraints)
        print(f"Warning: No available staff for date {duty_days[i][0].strftime('%Y-%m-%d')}", file=sys.stderr)

# Constraint #1: No more than one duty per week per person
for j in range(len(staff_df)):
    week_groups = {}
    for i, (day, _, _) in enumerate(duty_days):
        week_num = day.isocalendar()[1]  # ISO week number
        week_groups.setdefault(week_num, []).append(i)

    for week_num, days_in_week in week_groups.items():
        vars_in_week = [assignments[(i, j)] for i in days_in_week if (i, j) in assignments] # e.g. week 34, when can Keith do duty?
        if vars_in_week:  # Only add if staff is eligible for that week
            model.Add(sum(vars_in_week) <= 1)

# Constraint #2: 4 day gap for the same person
for j in range(len(staff_df)):  # Loop over staff
    if staff_df.loc[j, "Name"].strip().lower() in [n.lower() for n in frozen_names]: # don't have to deal with frozen names
      continue
    for i, (date_i, _, _) in enumerate(duty_days):
        for k, (date_k, _, _) in enumerate(duty_days):
            if abs((date_k - date_i).days) < 4 and i != k:
                if (i, j) in assignments and (k, j) in assignments:
                    model.Add(assignments[(i, j)] + assignments[(k, j)] <= 1)


# Calculate and balance scores
# Read current scores and scale them ("supports floats" after scaling)
current_scores_scaled = []
for j in range(len(staff_df)):
    raw = staff_df.loc[j, "Current Score"]
    try:
        val = float(raw)
    except Exception:
        val = 0.0
    current_scores_scaled.append(int(round(val * SCALE)))

# Precompute maximum possible month points (scaled)
max_month_points = sum(int_p for (_, _, _, int_p) in duty_days_scaled)

staff_scores = []

# === SOLVE OPTIMAL SOLUTION WITH OBJECTIVE FUNCTION ===
for j in range(len(staff_df)):
    # Sum of all scaled points assigned this month for staff j
    assigned_points_expr = sum(assignments[(i, j)] * duty_days_scaled[i][3]
                               for i in range(len(duty_days_scaled)) if (i, j) in assignments)

    # Add the current score (constant, scaled)
    # compute reasonable bounds for the total_score variable
    lower_bound = min(current_scores_scaled)  # at least the minimum current score
    upper_bound = max(current_scores_scaled) + max_month_points

    total_score = model.NewIntVar(lower_bound - max_month_points, upper_bound + max_month_points, f"score_{j}")
    model.Add(total_score == assigned_points_expr + current_scores_scaled[j])

    staff_scores.append(total_score)

# max and min score across staff (scaled)
# set bounds reasonably
global_min_bound = min(current_scores_scaled) - max_month_points
global_max_bound = max(current_scores_scaled) + max_month_points

# Compute a fixed target score (scaled)
# Average current score
avg_current_score = sum(current_scores_scaled) / len(current_scores_scaled)
# Average monthly points per person
avg_month_points = int(sum(int_p for (_, _, _, int_p) in duty_days_scaled) / (len(staff_df) - len(frozen_names)))
target_score_scaled = int(round(avg_current_score + avg_month_points))

# Create deviation variables
deviations = []
for j, score_var in enumerate(staff_scores):
    dev = model.NewIntVar(0, max_month_points, f"dev_{j}")
    model.Add(dev >= score_var - target_score_scaled)
    model.Add(dev >= target_score_scaled - score_var)
    deviations.append(dev)

# Objective: minimize total deviation
model.Minimize(sum(deviations))

# Solve the model
solver = cp_model.CpSolver()
solver.parameters.max_time_in_seconds = 10.0  # limit to 15 seconds (otherwise it will take about 1-4 mins to solve with no improved performance or may loop infinitely)
status = solver.Solve(model)

# Output results
if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
    schedule = []
    final_scores_scaled = [current_scores_scaled[j] for j in range(len(staff_df))]

    for i, (date, _, _, point_scaled) in enumerate(duty_days_scaled):
        for j in range(len(staff_df)):
            if (i, j) in assignments and solver.Value(assignments[(i, j)]):
                name = staff_df.loc[j, "Name"]
                # convert scaled points back to float for output convenience
                pts_float = point_scaled / SCALE
                schedule.append({"Date": date.strftime('%Y-%m-%d'), "Assigned To": name, "Points": pts_float})
                final_scores_scaled[j] += point_scaled

    schedule_df = pd.DataFrame(schedule)

    # Build mapping of actual duty days per person
    actual_duties = defaultdict(list)
    for _, row in schedule_df.iterrows():
        actual_duties[row["Assigned To"].strip().lower()].append(
            datetime.strptime(row["Date"], "%Y-%m-%d")
        )

    # Prepare for standby assignment
    staff_list = list(staff_df["Name"].str.strip())
    # Exclude frozen staff from standby eligibility
    staff_list = [
        name.strip()
        for name in staff_df["Name"]
        if name.strip().lower() not in frozen_names
    ]
    standby_schedule = []
    standby_counts = {name: 0 for name in staff_list}  # Track how many standbys each person has

    # Assign standby evenly with ≥4-day gap rule
    for i, (duty_date, _, _, _) in enumerate(duty_days_scaled):
        # Sort staff by current standby count (lowest first) for balancing
        sorted_candidates = sorted(staff_list, key=lambda n: standby_counts[n])

        assigned = False
        for candidate in sorted_candidates:
            cand_lower = candidate.lower()

            # Check the 4-day gap from actual duties
            too_close = any(abs((duty_date - ad).days) < 4 for ad in actual_duties[cand_lower])
            if too_close:
                continue

            # Assign standby
            standby_schedule.append({
                "Date": duty_date.strftime("%Y-%m-%d"),
                "Standby": candidate
            })
            standby_counts[candidate] += 1
            assigned = True
            break

        if not assigned:
            standby_schedule.append({
                "Date": duty_date.strftime("%Y-%m-%d"),
                "Standby": "No eligible staff"
            })

    # Create standby DataFrame
    standby_df = pd.DataFrame(standby_schedule)

    # convert final scaled scores back to floats with up to 3 decimal places
    final_scores = [round(s / SCALE, 3) for s in final_scores_scaled]
    score_df = pd.DataFrame({
        "Name": staff_df["Name"],
        "Score After Planning": final_scores
    })

    # Minus off the average points for the next month planning use
    score_df["Next Score to Use"] = score_df["Score After Planning"] - (avg_month_points / 1000)

    # If frozen person exists,  overwrite the above and keep their Next Score same as their current score
    for frozen_name in frozen_names:

      score_df.loc[score_df["Name"].str.lower() == frozen_name,
                  "Next Score to Use"] = staff_df.loc[
                      staff_df["Name"].str.lower() == frozen_name,
                      "Current Score"
                  ].values

      # Add (Frozen) next to names that were frozen for this month
      score_df.loc[score_df["Name"].str.lower() == frozen_name, "Name"] = str(frozen_name).upper() + " (Frozen)"

    # Update scaled down average month points
    score_df.loc[0, "Average Duty Score"] = avg_month_points / 1000

    # Merge standby names into schedule_df
    schedule_df["Standby"] = [entry["Standby"] for entry in standby_schedule]

    with pd.ExcelWriter("Duty_Planner_Combined.xlsx", engine='openpyxl') as writer:
      schedule_df.to_excel(writer, sheet_name="Duty Schedule", index=False)
      score_df.to_excel(writer, sheet_name="Updated Scores", index=False)

    print("Exported: Duty_Planner_Combined.xlsx")

else:
    print("No feasible solution found.")