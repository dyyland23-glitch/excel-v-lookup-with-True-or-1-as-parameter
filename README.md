# excel-v-lookup-with-True-or-1-as-parameter


# Tier-Based Volume Discounts with Excel VLOOKUP (`TRUE` / `1`)

This repository demonstrates how to implement dynamic, tiered discount schedules in Excel using `VLOOKUP` with an approximate match (`TRUE` or `1`).

---

## Why This Is Important

In business scenarios—such as wholesale pricing, tiered shipping rates, tax brackets, and volume commission schedules—lookup values rarely match exact numbers. 

Using `VLOOKUP(..., FALSE)` requires every single possible number to exist in the lookup table, which is impractical for continuous ranges (e.g., quantities like 14, 27, 43). 

Setting the fourth argument to **`TRUE`** (or **`1`**) enables **range matching**, allowing a compact reference table to classify any number into its appropriate tier automatically.

---
![VLOOKUP Formula Breakdown](https://lh3.googleusercontent.com/rd-gg/AJWXcNesWXObkVuetaijTC2IyD4mpNU_xmEjK8K3m6txDWORHU1q0c1kDmm69BUqkzX4eGlujg2DUQBywFRa8gP9vGxQfI0Mz_erWBh8WQkvVCtyl3qbO4pRC0JZP3Fqf1uGEPwqS8-CZzLEJLujHqKnd16Fq0yfGfM0oTsR6SeO8g9lXBgSzPA4uKnTqXW4hBlNk9xCjw96Uqn-9ltMjtkSvVBO8snHLzH9--GguQp1aDYle4HbP0l0G9XCMXdzJFnHxnzBRP9lgBN8zcysmKnlEKdquhQTeaBEMZZTqhA4QOomtLkzh9egn4fAGIg5y6TaM5edA7PixwT24hXc73ovnN5_b3pAlxatXXstIFl7RSUELuvTrzap2gtET7zxrjmrW4eNXzcZAS7mg93LvEwEEnBCROf0cNdT5r9tbSoRDk4wECHV7iskQJPrJAT7nu2lFbkJTQ-taciVV-EmeF_E0oJirKWjsOqN8NPvZyJujUfpx_RFTGDifZ-2PrEfEAYZ_CdqdBh4TiGTZUg0Ois67R3v4VxHYG3UgjTG7OqIc8l_fYcQHSI-e-0j4PhBOQdYXbt-26_bBhM02Mz-FAFdms9SRDCeoc8Fj_CRCOUXVbzi60X370CKBYUEy0dZ-wGta3I1SC9rMI0hr8n-ztoMwTvC-FOynTZ3BlqavA8bxVM9WZiftYELiv4Co6Erg767PjceOdiX1DLQs4-8yK-_gHl4O5C6jgGY9t4s2ZOszBMnbdeVbwlIvAjRcTtPtUm54uVRDHB_RhWgAIZDezVMT7LVrPQQQfRm8i82DvzlKUR_7q6Q-dkSnxRs2DjvYd6-uneDwmIRwSkKyq93E1wegs5npJKjnFQ-pZOJ8wnHwIjJJ3XUoA0sLuqIyI0Ue6rPXCYwU-BxN2yEoGUdAJ4cfJHuE-2Xqu1ROFFqwh1UhpXM_8barIbLXRrX05whfZKN7ExHWroAKexPHCzlh8gWeZidbH2x0Xc7qv5504AjPKOwYintPuBpafheFFcOb88pTDV0qMyp8o5cAtxuv5jF3dlO1A_CjuHj5wMMrtJhVPbVPFnZbkuW4hgzNoNubJkZxQzT7FSExmzoWngXGTd2CpYRfuFUH9v-HdvDapTC5V_Al95l_vWQOZo9YZbHoUdIqJHL98C50_pb0m-b1FXb_jJESMoBKZ_uf-XoPu_pN4jrx4PIhGQN5SUMsRcHUh0_zkTBPR3CXjWWato60KAp5AFUPZsjebcxOdvuujaResVLJ9I2AEdY8Zty4ZQWzzP8qP_9kFJXHQIWmy9nTm9oOKPRfNfYgcqHOZDmesiiRZ3SrH32PpKSBrPw7oJo0RHwJluX5gpJzX3aFtynVxGvg0xMe00VFIqYHr0gGj3SHz4GdTXxD5LPea7L1nQnSHNRFonw2ddpgiqkakNmvhqor3IZVUr7A8gFipfN-kW681HJaRLnb4xc7MrEE4XdtVIkLvJs4EPGxaLi=s1024-rw)

## The Formula Breakdown

```excel
=VLOOKUP(D2, $F$11:$G$16, 2, TRUE)
