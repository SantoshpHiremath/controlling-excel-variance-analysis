"""
Generates a synthetic but realistic controlling dataset for a mid-size
family business (bakery-chain-shaped: production, retail/stores,
logistics, admin cost centers) over 24 months, split into a budget plan
and monthly actuals with deliberate, realistic variance patterns.

Deterministic given a seed, so downstream pivot/report generation and
tests are reproducible.
"""
import random
from dataclasses import dataclass


COST_CENTERS = [
    ("1000", "Produktion Backwaren", "Herstellkosten"),
    ("1100", "Produktion Konditorei", "Herstellkosten"),
    ("2000", "Filialen München Zentrum", "Vertrieb"),
    ("2100", "Filialen München Umland", "Vertrieb"),
    ("2200", "Filialen Franchise", "Vertrieb"),
    ("3000", "Logistik & Fuhrpark", "Logistik"),
    ("4000", "Einkauf & Wareneinsatz", "Material"),
    ("5000", "Verwaltung", "Verwaltung"),
    ("6000", "Marketing", "Verwaltung"),
]

ACCOUNTS = [
    ("60000", "Personalkosten"),
    ("61000", "Materialaufwand"),
    ("62000", "Energiekosten"),
    ("63000", "Miete & Nebenkosten"),
    ("64000", "Fuhrpark & Logistik"),
    ("65000", "Marketing & Werbung"),
    ("66000", "Sonstige betriebliche Aufwendungen"),
]

MONTHS = [f"{y}-{m:02d}" for y in (2025, 2026) for m in range(1, 13)]


@dataclass
class Row:
    period: str
    cost_center_id: str
    cost_center_name: str
    division: str
    account_id: str
    account_name: str
    budget_eur: float
    actual_eur: float


def _base_monthly_budget(cc_id: str, account_id: str, rng: random.Random) -> float:
    """A stable per (cost-center, account) base budget, with realistic
    scale differences (Personalkosten >> Energiekosten, production cost
    centers >> admin, etc.)."""
    cc_scale = {
        "1000": 3.2, "1100": 1.4, "2000": 2.6, "2100": 1.8,
        "2200": 1.1, "3000": 1.6, "4000": 2.2, "5000": 1.0, "6000": 0.6,
    }[cc_id]
    account_base = {
        "60000": 42000, "61000": 38000, "62000": 6500, "63000": 9000,
        "64000": 7000, "65000": 4000, "66000": 3000,
    }[account_id]
    # +/- 8% stable idiosyncratic offset per cost-center/account pair
    offset = rng.uniform(0.92, 1.08)
    return round(account_base * cc_scale * offset, 2)


def generate_rows(seed: int = 42) -> list:
    rng = random.Random(seed)
    rows = []

    # Stable per-(cc,account) budget baseline, with seasonal index for
    # a bakery business (higher in Nov/Dec for Konditorei/Marketing,
    # slightly lower in summer for heating/energy accounts).
    baseline = {}
    for cc_id, cc_name, division in COST_CENTERS:
        for acc_id, acc_name in ACCOUNTS:
            baseline[(cc_id, acc_id)] = _base_monthly_budget(cc_id, acc_id, rng)

    def seasonal_index(period: str, cc_id: str, acc_id: str) -> float:
        month = int(period.split("-")[1])
        idx = 1.0
        if acc_id == "62000":  # Energiekosten: higher in winter
            idx *= 1.25 if month in (11, 12, 1, 2) else 0.90
        if cc_id == "1100" and month in (11, 12):  # Konditorei: Weihnachtsgeschäft
            idx *= 1.45
        if cc_id == "6000" and month in (11, 12):  # Marketing: Weihnachtskampagne
            idx *= 1.6
        if cc_id in ("2000", "2100", "2200") and month in (6, 7, 8):
            idx *= 1.08  # summer retail bump
        return idx

    # A handful of cost centers get a *real, sustained* deviation pattern
    # (not just noise) so the variance analysis has genuine signal:
    #   - 3000 (Logistik): fuel/energy cost overrun starting mid-2026
    #   - 1000 (Produktion Backwaren): a one-off equipment repair spike
    #   - 6000 (Marketing): consistently under budget (campaign underspend)
    overrun_start = "2026-04"

    for period in MONTHS:
        for cc_id, cc_name, division in COST_CENTERS:
            for acc_id, acc_name in ACCOUNTS:
                base = baseline[(cc_id, acc_id)]
                s_idx = seasonal_index(period, cc_id, acc_id)
                budget = round(base * s_idx, 2)

                # Actual = budget * noise, plus structural patterns
                noise = rng.gauss(1.0, 0.06)
                actual_factor = noise

                if cc_id == "3000" and acc_id == "64000" and period >= overrun_start:
                    actual_factor *= 1.28  # sustained fuel cost overrun

                if cc_id == "1000" and acc_id == "66000" and period == "2026-03":
                    actual_factor *= 3.1  # one-off equipment repair spike

                if cc_id == "6000":
                    actual_factor *= 0.88  # consistent marketing underspend

                actual = round(budget * actual_factor, 2)

                rows.append(Row(
                    period=period, cost_center_id=cc_id, cost_center_name=cc_name,
                    division=division, account_id=acc_id, account_name=acc_name,
                    budget_eur=budget, actual_eur=actual,
                ))

    return rows


if __name__ == "__main__":
    rows = generate_rows()
    print(f"Generated {len(rows)} rows across {len(MONTHS)} months, "
          f"{len(COST_CENTERS)} cost centers, {len(ACCOUNTS)} accounts.")
