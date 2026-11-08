

# Tier-Based Volume Discounts with Excel VLOOKUP (`TRUE` / `1`)

This repository demonstrates how to implement dynamic, tiered discount schedules in Excel using `VLOOKUP` with an approximate match (`TRUE` or `1`).

---

## Why This Is Important

In business scenarios—such as wholesale pricing, tiered shipping rates, tax brackets, and volume commission schedules—lookup values rarely match exact numbers. 

Using `VLOOKUP(..., FALSE)` requires every single possible number to exist in the lookup table, which is impractical for continuous ranges (e.g., quantities like 14, 27, 43). 

Setting the fourth argument to **`TRUE`** (or **`1`**) enables **range matching**, allowing a compact reference table to classify any number into its appropriate tier automatically.

---
![](vllokup%20true%20=1.webp)

## The Formula Breakdown

```excel
=VLOOKUP(D2, $F$11:$G$16, 2, TRUE)



The number **`2`** (the `col_index_num` argument) is a **relative count within the selected table range**, not the letter of the worksheet column.

---

### How the Column Index Number Works

When you write:

```excel
=VLOOKUP(D2, $F$11:$G$16, 2, TRUE)

```

Excel only counts columns inside the boundary defined in the table range (`$F$11:$G$16`):

1. **Column 1 is always the leftmost column of your selected range:**
* In `$F$11:$G$16`, the first column selected is **Column F**. Therefore, **Column F = 1**.


2. **Column 2 is the next column to the right:**
* One column over from F is **Column G**. Therefore, **Column G = 2**.



Because `2` tells Excel to pull data from the 2nd column of `$F$11:$G$16`, it pulls the value from **Column G** (the Discount %).

---

### What if the Range Started at Column A?

If your reference table started in Column A instead, the number you need depends on how wide your range is:

* **Example 1: Range is `$A$2:$C$2000**`
* Column A = 1
* Column B = 2
* Column C = 3
* *To return Column C, you enter `3`.*


* **Example 2: Range is `$A$11:$G$16**`
If you selected everything from A through G:
* Column A = 1
* Column B = 2
* Column C = 3
* Column D = 4
* Column E = 5
* Column F = 6
* Column G = 7
* *To return Column G here, you would have to write `7`, not `2`.*



The key takeaway is that `1` is always whichever column starts your table array. If the range starts at `F`, then `F` is `1` and `G` is `2`.
