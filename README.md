# excel-v-lookup-with-True-or-1-as-parameter


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
