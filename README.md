# Excel Formula: Extract Item Codes and Match Case Quantity

## Overview
This solution extracts item codes from SKUs and uses `VLOOKUP` to find corresponding case quantities from another file.

##
```excel
=IFERROR(VLOOKUP(LEFT(B3, MIN(IFERROR(FIND({"-"," "}, B3), LEN(B3)+1)) - 1),'[Hammont all items (1).xlsx]Sheet1'!$A$2:$C$640,3,FALSE),"No match found")
````
## How It Works:
`1.` Extracts the item code from B3 (before - or first space).

`2.` Uses VLOOKUP to find the case quantity from another file.

`3.` Handles errors to return "No match found" if there's no match.

## 📦 Example Data: Item Numbers and Case Quantities

| Item Number  | Name                                                      | CASE QTY |
|-------------|------------------------------------------------------------|----------|
| 242-GOLD   | Premium Gift Bags Gold with Ribbon Handles 9x7x4 (12 Pack)  | 12       |
| 243-SILVER | Premium Gift Bags Silver with Ribbon Handles 9x7x4 inches (12 Pack) | 12       |
| CB-052     | Polka Dot Candy Boxes 4.5x3.75x2.25 Inches (18 Pack)        | 16       |
| CB-053     | Its a Girl Candy Boxes 4.5x3.75x2.25 Inches (18 Pack)       | 16       |
| CB-054     | Its a Boy Candy Boxes 4.5x3.75x2.25 Inches (18 Pack)        | 16       |
| CB-055     | Upsherin Candy Boxes 4.5x3.75x2.25 Inches (18 Pack)         | 16       |
| CB-056     | Vacht Nacht Candy Boxes 4.5x3.75x2.25 Inches (18 Pack)      | 16       |
| CB-057     | Birthday Boys Candy Boxes 4.5x3.75x2.25 Inches (18 Pack)    | 16       |
| CB-058     | Birthday Girls Candy Boxes 4.5x3.75x2.25 Inches (18 Pack)   | 16       |


## Example Results:
| SKU                                | Extracted Item Code | Matched Case Quantity |
|------------------------------------|--------------------|--------------|
| TB-37 - Kraft Window Treat Box    | TB-37             | 12           |
| HM1213 Clear Acrylic Boxes        | HM1213            | 12           |
| HM1214 Clear Acrylic Boxes        | HM1214            | 24           |
| HM1212-Lucite Boxes 4''x4''x4''   | HM1212            | 8            |
| HM1215 Clear Acrylic Boxes        | HM1215            | 8            |
| HM1218 Clear Acrylic Boxes        | HM1218            | 18           |
| HM1216 Clear Acrylic Boxes        | HM1216            | 18           |
| HM1217 Clear Acrylic Boxes        | HM1217            | 24           |
| HM1449-Icecream bowl set          | HM1449            | 10           |
| HM1308 - Wooden Box               | HM1308            | 10           |
| HM1322 - 6 section Wooden box     | HM1322            | 20           |


## Why This Is Useful
Helps clean messy SKU data.
Automates matching without manual effort.
Works even if formats are inconsistent.
