from datetime import date, timedelta


def workdays_month_plus_10(year, month):
    def is_workday(d):
        fixed = {(1,1),(1,6),(4,25),(5,1),(6,2),(8,15),(11,1),(12,8),(12,25),(12,26)}
        return d.weekday() < 5 and (d.month, d.day) not in fixed
    out = []
    d = date(year, month, 1)
    while d.month == month:
        if is_workday(d):
            out.append(d.isoformat())
        d += timedelta(days=1)
    extra = 0
    while extra < 10:
        if is_workday(d):
            out.append(d.isoformat())
            extra += 1
        d += timedelta(days=1)
    return out


if __name__ == "__main__":
    days = workdays_month_plus_10(2026, 7)
    assert days[0] == "2026-07-01", days[0]
    assert days[-1] == "2026-08-14", days[-1]
    assert len([d for d in days if d.startswith("2026-08")]) == 10
    print("OK luglio 2026:", days[0], "->", days[-1], "giorni", len(days))
