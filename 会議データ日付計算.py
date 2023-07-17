import datetime

def calculate_date():

    today = datetime.date.today()
    year = today.year
    month = today.month
    day = today.day

    first_day_of_previous_month = today.replace(day=1) - datetime.timedelta(days=1)
    if today.month == 1:
        first_day_of_previous_month = today.replace(year=today.year - 1, month=12, day=1)
        last_day_of_previous_month = first_day_of_previous_month.replace(month = 12, day=31)
    else:
        last_day_of_previous_month = today.replace(day=1) - datetime.timedelta(days=1)

    y1 = first_day_of_previous_month.year
    m1 = first_day_of_previous_month.month
    d1 = 1

    y2 = last_day_of_previous_month.year
    m2 = last_day_of_previous_month.month
    d2 = last_day_of_previous_month.day

    current_year = today.year
    current_month = today.month

    if current_month <= 8:
        y1_mid = current_year - 1
        m1_mid = 8
        d1_mid = 1

        last_day_of_previous_month_mid = today.replace(day=1) - datetime.timedelta(days=1)
        y2_mid = last_day_of_previous_month_mid.year
        m2_mid = last_day_of_previous_month_mid.month
        d2_mid = last_day_of_previous_month_mid.day
    else:
        y1_mid = current_year
        m1_mid = 8
        d1_mid = 1

        last_day_of_previous_month_mid = today.replace(day=1) - datetime.timedelta(days=1)
        y2_mid = last_day_of_previous_month_mid.year
        m2_mid = last_day_of_previous_month_mid.month
        d2_mid = last_day_of_previous_month_mid.day

    print("直近月末")
    print(f"y1 = {y1}, m1 = {m1}, d1 = {d1}")
    print(f"y2 = {y2}, m2 = {m2}, d2 = {d2}")
    print("\n期中")
    print(f"y1 = {y1_mid}, m1 = {m1_mid}, d1 = {d1_mid}")
    print(f"y2 = {y2_mid}, m2 = {m2_mid}, d2 = {d2_mid}")



calculate_date()
