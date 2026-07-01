from datetime import date, timedelta


def workdays_month_plus_10(year, month):
    def is_workday(d):
        fixed = {(1, 1), (1, 6), (4, 25), (5, 1), (6, 2), (8, 15), (11, 1), (12, 8), (12, 25), (12, 26)}
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


def first_workday(year, month):
    return next(d for d in workdays_month_plus_10(year, month) if d.startswith(f"{year}-{month:02d}"))


def assert_month_rule(year, month):
    days = workdays_month_plus_10(year, month)
    month_prefix = f"{year}-{month:02d}"
    month_days = [d for d in days if d.startswith(month_prefix)]
    extra_days = [d for d in days if not d.startswith(month_prefix)]

    assert month_days, f"Nessun giorno lavorativo nel mese {month_prefix}"
    assert days[0] == first_workday(year, month), (year, month, days[0])
    assert len(extra_days) == 10, (year, month, len(extra_days), extra_days)
    assert not any(d.endswith('-30') and month == 7 and d.startswith(f"{year}-06") for d in days), days

    return days[0], days[-1], len(days)


if __name__ == "__main__":
    for year, month in [(2026, 1), (2026, 2), (2026, 7), (2026, 8), (2026, 9), (2026, 12)]:
        start, end, total = assert_month_rule(year, month)
        print(f"OK {year}-{month:02d}: {start} -> {end}, giorni {total}")
