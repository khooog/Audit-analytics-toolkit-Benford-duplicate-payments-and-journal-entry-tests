# Audit analytics toolkit: Benford, duplicate payments, round sums and journal entry tests

All data in this repository is sample / demonstration data generated with a fixed seed. No real company, vendor, employee or transaction is represented. Anomalies and benign look alike patterns were seeded on purpose so detection can be tested.

## Problem

Auditors must look for duplicate payments, threshold avoidance and unusual manual journal entries across populations too large to read line by line. This toolkit runs a set of standard analytics over a year of payments and journal entries, scores every item, and produces a ranked list for targeted follow up.

## Files

* audit_analytics.py: data generator, pandas implementation of every test, workbook builder, self test
* audit_analytics_workbook.xlsx: working file with live formulas (Parameters, Payments, Benford, Vendor_Benford, JE_Lines, JE_Entries, Exceptions, Answer_Key, Checks)
* verify_workbook.py: recalculates the workbook in LibreOffice and compares it row by row with the pandas tests
* audit_analytics_case_study.pdf: two page case study

## How to run

    pip install pandas openpyxl
    python audit_analytics.py out

This writes out/audit_analytics_workbook.xlsx and out/results.json and prints a summary.

Optional independent check (needs LibreOffice with soffice on the PATH):

    python verify_workbook.py out/audit_analytics_workbook.xlsx

## Tests

Payments: duplicate (vendor, normalised invoice number, amount), near duplicate within a window, round sums, just below the approval limit, same day splits above the limit, approval below required authority, vendor bank account equal to an employee bank account, and a Benford first digit test per vendor.

Population: Benford first digit test on all payments with z statistics, chi square and MAD.

Journal entries: unbalanced, weekend posting, late posting, year end window, unauthorised user, manual entry to revenue, round total, no description, manual entry just below the approval limit, manual entry to suspense, number sequence gap.

Every flag has a weight on the Parameters sheet. Items at or above the cutoff score are High.

## Results (sample data)

* 630 payments worth 9,499,168.93 and 791 journal entries (1,729 lines).
* Benford on all payments: MAD 0.0218 (nonconformity), chi square 54.06 against a critical value of 15.507. Excluding vendor V031, MAD is 0.0126. At vendor level V031 is the only vendor flagged (MAD 0.201).
* All 10 seeded duplicate payments found, all 3 unbalanced entries found, all 5 year end manual revenue entries found, and the vendor with an employee bank account found.
* The just below limit band test detected 39 of 45 seeded threshold payments. The vendor level Benford test caught the pattern.
* High list: 49 payments and 8 journal entries. 17 of the High payments are fixed rent and licence payments that a reviewer would clear against the contract.

## Controls

* Checks sheet ties counts, debits to credits, and the High lists to the flag columns. Overall result reads PASS.
* verify_workbook.py found no mismatches in 630 payments and 791 entries, and no formula errors.

## Limitations

* The answer key was written by the same author as the tests, so recall is optimistic. Real ledgers are messier.
* Some tests have low precision, for example the year end window test.
* Benford does not apply to assigned or fixed amounts. Nonconformity is a reason to look, not a finding.
* Thresholds and weights are judgment and must be set for each engagement.
* The output selects specific items for follow up. It is not an audit sample and results cannot be projected (ISA 530).
* The duplicate test does not catch transposed digits in invoice numbers.
* Real use needs checks on the completeness and accuracy of the data extract first.
* The revised ISA 240 applies to audits of periods beginning on or after 15 Dec 2026.

Author: Saqib Anwar Khan, CPA
