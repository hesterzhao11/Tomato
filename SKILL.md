---
name: reimbursement-reader
description: Read, classify, and reconcile Chinese sales reimbursement materials for OA expense submission. Use when Codex needs to process a reimbursement folder, read Excel expense ledgers, extract PDF invoice/trip/toll/hotel data, generate monthly OA填报清单/OA草稿/查漏清单, match invoices to mileage, meals, hotels, phone, tolls, parking, taxi/train travel, or assist with OA reimbursement entry after the user logs in.
---

# Reimbursement Reader

## Core Workflow

1. Ground in the folder before asking questions: list files, inspect the main Excel ledger, and identify month folders and candidate invoices.
2. Treat the user's original receipts, screenshots, invoices, and OA exports as read-only unless explicitly asked to reorganize files.
3. For spreadsheets, use the bundled spreadsheet/document runtime when available; otherwise use `openpyxl`/`pypdf` through the local Python runtime.
4. Generate a monthly OA work product with these sheets whenever the user asks to process a month:
   - `总览`
   - `OA填报清单`
   - `原始清单记录`
   - `发票索引`
   - `查漏清单`
   - `规则速查`
5. Before finalizing, verify the workbook opens, row counts and totals are plausible, and no obvious formula errors or blank core sheets exist.

## Image Mileage Handling

When reading Amap/Gaode screenshots, deduplicate before summing mileage. Treat a route card as the same trip when date/time, start, destination, and km match, even if it appears in overlapping screenshots. Count each route card once per day. Do not add a monthly total shown by the app on top of the individual route cards.

When updating the ledger from screenshots, write the daily total km as a numeric value and write fuel amount as `km × 1.4`, rounded to 2 decimals. If a visible screenshot is cut off and the km cannot be read, use existing ledger detail only when it clearly matches the visible same-day cards; otherwise mark it as needing confirmation.

## Default Project

Default to `/Users/hehe/Documents/报销` when the user says "报销", "读取报销", "整理报销", "OA清单", or a month such as `202605` without giving another path.

Use `2026/报销清单.xlsx` as the default ledger when present. Do not assume old rows are still pending; trust the user's latest edited workbook.

## Monthly Generation

Use `scripts/reimbursement_reader.py` for deterministic month workbooks:

```bash
python scripts/reimbursement_reader.py --root /Users/hehe/Documents/报销 --month 202605
```

The script writes:

```text
<root>/<year>/<year>.<month-number>/OA月度工作台_<yyyymm>.xlsx
```

If the user has not finished month-based filing yet, still run the script: it also pulls candidate files whose file name or PDF invoice date matches the target month.

## OA Entry Boundary

If assisting in OA after the user logs in:

- Fill fields, add OA rows, match/upload attachments, and save drafts.
- Stop before final submit, delete, withdraw, payment confirmation, or any irreversible action unless the user explicitly authorizes that exact action.
- Never ask for or store passwords, verification codes, or long-lived credentials.

## Expense Rules

Read `references/oa-rules.md` when preparing or reviewing OA rows. Keep the most important checks in the output:

- Private car mileage: amount = km × 1.4; needs Amap/Gaode historical route screenshots and enough fuel invoice value.
- Toll/ETC: needs ETC summary and toll invoices; dates should match the self-drive trip.
- Didi/taxi: needs both the electronic invoice and the trip reimbursement itinerary.
- High-speed rail/train: needs the railway e-ticket invoice; use travel date, not invoice issue date, for monthly matching.
- Travel meal allowance: for trips outside Guangzhou over 8 hours, use RMB 120/day in South China including Hong Kong/Macau, RMB 200/day outside South China, and deduct same-day reimbursed meals (lunch RMB 60, dinner RMB 80).
- Meal/entertainment: needs customer company, customer names, colleague names, lunch/dinner, and business context.
- Hotel: needs dates, location, folio/water bill plus invoice; if no folio, use order screenshot and mark approval need.
- Phone: monthly invoice preferred; recharge invoice requires phone bill statement.
- Parking: list separately with date/place and receipt or payment record.

## Output Style

Use Chinese for user-facing reimbursement deliverables. Lead with the files created or updated, the number of OA rows, total amount, pending/gap count, and the next concrete items the user must supply.
