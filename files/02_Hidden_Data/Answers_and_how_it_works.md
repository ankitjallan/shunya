# "What you don't see": answers and how it works (slides 13–17, about 7 minutes)

**The story:** a client file from *Saurashtra Castings Pvt Ltd* (fictional), FY 2025-26 year-end, received on WhatsApp. Everything is fictional: the Aadhaar numbers deliberately fail the Verhoeff checksum, so none can be real.

---

## Slide 13: the brief (1 min)
Show the README sheet. Read the three questions:
1. Ledger: does the creditors ledger agree with the books?
2. Ageing: what is the **real** debtors figure?
3. Payroll: is this file safe to paste into ChatGPT?

> *"Thirty seconds each. Everything you need is in this file. You just can't see it yet."*

## Slide 14: Ledger (1.5 min), the hidden sheet
- **Visible:** total as per ledger **₹1,04,35,000**.
- **Hunt for 30 s:** "Does it agree with the books?" Nobody can tell.
- **Reveal:** right-click any tab › **Unhide** › `_verify`
  - Creditors as per ledger **₹1,04,35,000**
  - add JV-118 reversed on 31-Mar-26 (Kathiyawad Logistics) **₹4,85,000**
  - creditors as per books **₹1,09,20,000**
- **Say:** *"A ₹4.85 lakh reversal on the last day of the year, on a sheet nobody was meant to open."*
- **Bonus** (if the room is quick): Ctrl+A › font colour Automatic. Cell **G6** appears: *"JV-118 … NOT approved by director"* (white on white).

## Slide 15: Ageing (1.5 min), the real total
- **Visible:** TOTAL **₹2,04,90,000**.
- **Hunt for 30 s:** "What's the real debtors figure?"
- **Reveal:** **Ctrl+G** › type `THE_REAL_TOTAL` › Enter. The cursor flies to **BZ2000**: **₹2,30,65,000**. The label beside it in BY2000 reads *"Real total incl. Kathiyawad Traders (disputed ₹25,75,000, kept off the ageing)"*.
- **Say:** *"₹25.75 lakh parked two thousand rows down. Ctrl+End would have taken you there too."*

## Slide 16: Payroll (1.5 min), the risky columns
- **Visible:** Emp code, name, designation, gross, PF, net pay, mobile. It looks harmless.
- **Ask:** *"Who would paste this into ChatGPT to compute the TDS?"* Most hands go up.
- **Reveal:** select columns E:K › right-click › **Unhide**. Columns F–J appear: **PAN · Aadhaar · Bank A/c · DOB · Medical notes** (e.g. "Diabetic - insulin", "Pregnancy leave from May").
- **Say:** *"You didn't hide these columns. Somebody in the client's office did, months ago. And you were about to paste 12 people's Aadhaar and medical history into an AI tool. That's a DPDP problem, and it's yours."*

## Slide 17: one macro finds it all (1.5 min)
- Import **WhatIsHidden.bas** (Alt+F11 > File > Import File) and run **WhatIsHiddenHere** with the workbook active.
- In under a second, the **WHAT WAS HIDDEN** sheet shows **14 findings**. Read the ones the room never saw:
  | # | Found | Why a CA cares |
  |---|---|---|
  | 6 | Sales Register rows 28–42 hidden: **credit notes, taxable −₹8,07,000** | The total uses SUBTOTAL, so hidden credit notes silently drop out of the sales figure |
  | 11 | `_old_TB`: a **VERY hidden** sheet, the trial balance "given to the bank" with stock at ₹4.2 Cr (books ₹2.9 Cr) | It doesn't even appear in the Unhide list |
  | 8 | GST Reco rows 11–14 collapsed: **blocked credits u/s 17(5)** (director's car, Diwali party, club) claimed in books, ₹2,74,000 | ITC exposure |
  | 9 | GST Reco F5 in **1-pt font**: "ITC of ₹1,68,000 on director's car claimed in GSTR-3B" | The confession is in the file |
  | 7 | Note on Sales C10: "bill this in his brother's firm name" | Notes travel with the file |
  | 2 | Hidden text box: "DRAFT - NOT FOR BANK" | |
  | 12 | Hidden name `Commission_Rate` = 2.5% | Not visible in Name Manager |
  | 3 | Ageing column H: provision values blanked with `;;;` | Visible only in the formula bar |
  | 14 | Personal data: **12 PAN, 12 Aadhaar, 12 mobiles, 24 of them inside hidden columns** | DPDP |
- **Close:** *"Three by hand. Fourteen by one macro. Nobody in this room wrote a line of it. **Hidden does not mean absent.** Run it before you paste any client file into any AI tool. It's in your takeaway pack."*

## If something goes wrong
- **You revealed something too early:** close the stage file without saving and reopen it. It ships with everything hidden.
- **Reveal everything at the end (optional):** macro `Reveal_All` asks first, then shows every hidden item in red.

## What the finder checks (for Q&A)
1. Hidden sheets
2. Very hidden sheets
3. Hidden rows, with a total of their first number column
4. Hidden columns, with their headers
5. Collapsed groups
6. White-on-white text
7. `;;;` formats
8. Text in a font of 2 pt or less
9. Values parked far from the table
10. Named cells pointing far away
11. Hidden names
12. Notes
13. Hidden shapes
14. Links to other workbooks
15. PAN / Aadhaar / mobile patterns, including inside hidden areas

It only adds a report sheet. It never changes the data.
