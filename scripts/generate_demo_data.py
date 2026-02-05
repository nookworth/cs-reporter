"""
Generate demo Excel files for CS Reporter demonstration.

Creates two Excel files (current and previous month) with realistic
dummy data that works with the reporter tool.
"""

import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

# Configuration
OUTPUT_DIR = Path(__file__).parent.parent / "demo_data"
CURRENT_MONTH_FILE = OUTPUT_DIR / "demo_january_2026.xlsx"
PREVIOUS_MONTH_FILE = OUTPUT_DIR / "demo_december_2025.xlsx"

# Sheet names (generic for demo)
RETAIL_SHEET = "Retail Tickets"
SUPPLIER_SHEET = "Supplier Tickets"

# Data values
SATISFACTION_RATINGS = [
    "good",
    "good",
    "good",
    "good",
    "good with comment",
    "good with comment",
    "good with comment",
    "bad",
    "bad",
    "",  # empty
    "",
]

SUPPORT_CATEGORIES = [
    "Technical Support",
    "Technical Support",
    "Technical Support",
    "Account Management",
    "Account Management",
    "Billing Inquiry",
    "Billing Inquiry",
    "Product Information",
    "Order Status",
    "Returns & Refunds",
    "Integration Support",
    "",  # empty - will become "Uncategorized"
    "",
]

ORGANIZATION_NAMES = [
    "Acme Corporation",
    "Acme Corporation",
    "Acme Corporation",
    "Acme Corporation",
    "Globex Industries",
    "Globex Industries",
    "Globex Industries",
    "Initech Solutions",
    "Initech Solutions",
    "Umbrella Corp",
    "Umbrella Corp",
    "Stark Enterprises",
    "Wayne Industries",
    "Cyberdyne Systems",
    "Soylent Corp",
    "Wonka Industries",
]

ASSIGNEE_NAMES = [
    "Sarah Chen",
    "Sarah Chen",
    "Sarah Chen",
    "Marcus Johnson",
    "Marcus Johnson",
    "Emily Rodriguez",
    "Emily Rodriguez",
    "David Kim",
    "Rachel Green",
]


def random_date_in_month(year: int, month: int) -> date:
    """Generate a random date within a given month."""
    if month == 12:
        next_month = date(year + 1, 1, 1)
    else:
        next_month = date(year, month + 1, 1)

    start = date(year, month, 1)
    days_in_month = (next_month - start).days

    return start + timedelta(days=random.randint(0, days_in_month - 1))


def generate_ticket_data(
    num_tickets: int,
    year: int,
    month: int,
    include_org_names: bool = False
) -> list[dict]:
    """Generate ticket data for a single sheet."""
    tickets = []

    for i in range(num_tickets):
        created_date = random_date_in_month(year, month)

        # Resolution time: mostly 0-3 days, some longer
        if random.random() < 0.85:
            # 85% resolved within 3 days
            resolution_days = random.randint(0, 3)
        else:
            # 15% take longer (4-7 days)
            resolution_days = random.randint(4, 7)

        # 10% of tickets are unsolved
        if random.random() < 0.10:
            solved_date = None
        else:
            solved_date = created_date + timedelta(days=resolution_days)

        ticket = {
            "Ticket ID": f"TKT-{year}{month:02d}-{i+1:04d}",
            "Ticket created - Date": created_date.isoformat(),
            "Ticket solved - Date": solved_date.isoformat() if solved_date else None,
            "Ticket satisfaction rating": random.choice(SATISFACTION_RATINGS),
            "Support Category": random.choice(SUPPORT_CATEGORIES),
            "Assignee name": random.choice(ASSIGNEE_NAMES),
        }

        if include_org_names:
            ticket["Ticket organisation name"] = random.choice(ORGANIZATION_NAMES)

        tickets.append(ticket)

    return tickets


def create_excel_file(filepath: Path, year: int, month: int, retail_count: int, supplier_count: int):
    """Create an Excel file with retail and supplier sheets."""
    # Generate data
    retail_data = generate_ticket_data(retail_count, year, month, include_org_names=False)
    supplier_data = generate_ticket_data(supplier_count, year, month, include_org_names=True)

    # Create DataFrames
    retail_df = pd.DataFrame(retail_data)
    supplier_df = pd.DataFrame(supplier_data)

    # Write to Excel with multiple sheets
    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        retail_df.to_excel(writer, sheet_name=RETAIL_SHEET, index=False)
        supplier_df.to_excel(writer, sheet_name=SUPPLIER_SHEET, index=False)

    print(f"Created: {filepath}")
    print(f"  - {RETAIL_SHEET}: {retail_count} tickets")
    print(f"  - {SUPPLIER_SHEET}: {supplier_count} tickets")


def main():
    # Create output directory
    OUTPUT_DIR.mkdir(exist_ok=True)

    # Set random seed for reproducibility
    random.seed(42)

    print("Generating demo Excel files for CS Reporter...\n")

    # Create current month file (January 2026)
    create_excel_file(
        CURRENT_MONTH_FILE,
        year=2026,
        month=1,
        retail_count=92,
        supplier_count=48
    )

    print()

    # Create previous month file (December 2025)
    create_excel_file(
        PREVIOUS_MONTH_FILE,
        year=2025,
        month=12,
        retail_count=78,
        supplier_count=41
    )

    print("\nDone! Demo files created in:", OUTPUT_DIR)
    print("\nTo use with the reporter, update config/mapping.yaml to use:")
    print(f'  - Sheet names: "{RETAIL_SHEET}" and "{SUPPLIER_SHEET}"')
    print("  - Or use config/demo_mapping.yaml")


if __name__ == "__main__":
    main()
