# Prompt library: CA edition

Shunya se Shikhar: Excel x AI, Rajkot, 24 Oct 2026. Replace the [bracketed] parts with your details. Remove PAN, Aadhaar, mobile and bank numbers before pasting anything into an AI tool.


## Build tools with AI

### 1. Spec first: let AI interview you

```text
I want an Excel VBA tool for my CA office that [describe the job in one line]. Before writing any code, ask me 10 questions about inputs, outputs, edge cases (blank rows, duplicates, dates in dd-mm-yyyy, Indian number format) and who will use it. Then write a one-page specification I can approve. Do not write code until I say 'approved'.
```

### 2. What is hidden in this file? (the macro from the talk)

```text
Write an Excel VBA macro for a CA office: for the active workbook, list every hidden sheet, very hidden sheet, hidden row, hidden column, collapsed group, white-on-white text, ;;; formatted cell, hidden name, cell note and hidden shape into a new sheet 'What Was Hidden' with sheet name, address, what was hidden and how to reveal it. It must not change any data. Use Option Explicit, late binding, and comment each step.
```

### 3. Build from an approved spec

```text
Here is my approved specification: [paste spec]. Write the complete VBA as separate modules. Rules: Option Explicit everywhere, late binding only (no Tools > References), no MsgBox inside loops, every public macro has an error handler that tells the user what went wrong in plain English, money rounded half-up to 2 decimals (not VBA Round), dates handled as dd-mm-yyyy regardless of the PC's locale. Finish with a list of the sheets and buttons I need to create.
```

### 4. Test like an auditor

```text
Give me 15 test cases for this macro with the input rows and the hand-computed expected result for each, including edge cases: blank rows, duplicate invoice numbers, negative amounts, credit notes, dates in dd-mm-yyyy, text that looks like numbers, and leading zeros in GSTIN / PAN. Then write a VBA test routine that runs them and prints PASS / FAIL. Macro: [paste code]
```

### 5. Fix my macro

```text
This VBA macro stops with the error '[paste error and line]'. Explain the cause in one paragraph a non-programmer understands, then give the corrected procedure in full. Do not change anything else. Code: [paste code]
```

### 6. Explain a macro to an article

```text
Explain this VBA macro line by line to a first-year articled assistant who knows Excel but not programming. Group the explanation by what each block does for the office, and list anything risky (deleting data, sending email, overwriting files). Code: [paste code]
```


## Reconciliations

### 7. GSTR-2B vs purchase register

```text
Write VBA that matches my purchase register (sheet 'Books': GSTIN, Invoice No, Invoice Date, Taxable, IGST, CGST, SGST) with GSTR-2B (sheet '2B', same columns) on GSTIN + normalised invoice number (ignore case, spaces, slashes, leading zeros and FY suffixes like /25-26). Flag tax differences above Rs 1. Output sheets: Matched, Amount Mismatch, Only in Books, Only in 2B, and a summary of ITC at risk by supplier.
```

### 8. 26AS / AIS vs books (TDS)

```text
Write VBA that reconciles TDS as per Form 26AS / AIS (sheet '26AS': Deductor TAN, Name, Section, Amount Paid, TDS) with my books (sheet 'Books': Party, TAN, Section, Income, TDS). Match on TAN + section, then fuzzy-match names when TAN is missing. Show matched, short deducted, not in 26AS, not in books, and a client-wise summary I can send.
```

### 9. Bank reconciliation auto-match

```text
Write VBA that matches a bank statement (Date, Narration, Debit, Credit) with the bank book from Tally (Date, Particulars, Vch No, Debit, Credit). Match same amount within +/- 5 days, use cheque / UTR numbers in the narration when present, and produce the BRS format: balance as per books, cheques issued not presented, deposits not credited, bank charges / interest not in books, balance as per bank.
```

### 10. Duplicate payment detector

```text
Write VBA that finds possible duplicate vendor payments in a payments register (Date, Vendor, Invoice No, Amount, Voucher No): same vendor + same amount within 30 days, same invoice number after removing spaces / slashes, and vendor names that are 90% similar. Rank by risk and show the total at risk.
```


## Excel formulas

### 11. Explain and simplify a formula

```text
Explain this Excel formula step by step for a junior, then suggest a simpler or faster version (LET, XLOOKUP, FILTER if available) and say when the old one would give a different answer: [paste formula]
```

### 12. Indian FY, quarter and month labels

```text
Give me Excel formulas that take a date in A2 and return: financial year as '2026-27', FY quarter as 'Q1'..'Q4' (April-June = Q1), the GST return period as 'Oct-2026', and the TDS return quarter's due date (31 Jul, 31 Oct, 31 Jan, 31 May). Show each formula with a one-line explanation.
```

### 13. Receivables ageing buckets

```text
I have invoice date in B, due date in C, outstanding in D and an 'as on' date in H1. Give me formulas for days overdue and buckets Not due / 0-30 / 31-60 / 61-90 / 91-180 / >180, a pivot-ready layout, and a SUMIFS summary table by customer.
```

### 14. Clean a Tally export with Power Query

```text
Write Power Query (M) steps to clean a Tally ledger export: remove the header block and page totals, split 'Particulars' into ledger and narration, convert Dr / Cr into a signed amount, convert dates in dd-mm-yyyy, and keep the voucher type and number. Explain each step so I can repeat it in the Power Query editor.
```


## Tax and compliance (verify every answer)

### 15. Explain capital gains to a client

```text
Using only these computed figures, write a one-page note in simple English for my client explaining their capital gains for FY [year] and the tax: [paste the summary table]. Explain FIFO, grandfathering (31-Jan-2018), short vs long term and the Rs 1.25 lakh exemption in plain words. Do not add any figure that is not in the table.
```

### 16. Due-date checklist with sources

```text
List the statutory due dates under GST, TDS, income tax and the Companies Act for a private limited company (monthly GST filer, TDS deductor, tax audit) for [months]. For each, give the provision or notification I should check and mark anything you are not sure about. I will verify against official sources before use.
```

### 17. First draft of a notice reply

```text
Draft a reply to this GST / income-tax notice for my review. Structure: facts, issues raised, our submissions point-wise, documents enclosed, prayer. Keep a respectful tone and mark every legal position you are unsure of with [CHECK]. Notice text (personal identifiers removed): [paste text]
```


## Client communication

### 18. Fee reminder in Gujarati

```text
Write a polite WhatsApp fee reminder in Gujarati for a long-standing client using exactly these facts: name [name], amount Rs [amount], oldest bill [invoice no] dated [date], UPI ID [upi id]. Under 60 words, warm and respectful, end with the firm name [firm].
```

### 19. Documents pending, in Hindi

```text
Write a short WhatsApp message in Hindi asking a client to send these pending documents for their [return] due on [date]: [list]. Respectful, mention that the return cannot be filed without them, under 70 words, end with [firm].
```

### 20. Explain a notice to a client

```text
Explain this tax notice to my client in simple Gujarati (or English) in 5 bullet points: what it is, why it came, what they must send, by when, and what happens if ignored. No legal jargon. Notice summary: [paste summary without PAN or other identifiers].
```


## Audit

### 21. Journal entry testing (SA 240)

```text
Write VBA that runs journal entry tests on a GL export (Date, Voucher No, Voucher Type, Ledger, Debit, Credit, Narration, Entered By, Entry Date): entries on Sundays / holidays, after the period end, round amounts, just below approval limits, manual entries to revenue or cash, blank or vague narrations, unusual users, and gaps in voucher numbers. One sheet per test, with counts and totals for the working paper.
```

### 22. Reproducible audit sample

```text
Write VBA that draws a reproducible audit sample from a population on the active sheet: monetary unit sampling and simple random sampling, with a seed I enter so the same sample can be redrawn. Output the selected items, the sampling interval, and a working-paper header (client, period, area, prepared by, date).
```

### 23. Working paper narrative

```text
From these test results, write a concise audit working-paper narrative: objective, procedure, population, sample, exceptions found, conclusion. Formal, past tense, no figures beyond those given: [paste results].
```


## Safe AI use

### 24. Mask personal data before you paste

```text
Write a VBA macro that makes a masked copy of the active sheet: PAN becomes XXXXX1234X, Aadhaar keeps only the last 4 digits, mobile numbers and bank account numbers keep only the last 4 digits, e-mail addresses are replaced with [EMAIL], and names in a column I choose become Client 1, Client 2... Never change the original sheet.
```

### 25. DPDP check before sharing

```text
Here are the column headers and 3 sample rows of a sheet I want to share (identifiers already masked): [paste]. List which columns are personal data or sensitive personal data under India's DPDP Act, what I should remove or mask, and whether client consent is needed to share it with a third-party tool.
```
