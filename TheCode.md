# The Code
| [Total Row](#The-Total-Row) | [Depth Calcs](#Depth-Calculations) | [Junctions](#Junctions) | [Settings](#Settings) |
| :---: | :---: | :---: | :---: |

While much of the code relies on simple `VLOOKUP` and *Structured References* with tables some of it required careful planning of nested if statements (*N.B. In its original implementation the design was to avoid a macro enabled workbook*).

## The Total Row
In it's earliest implementation intended users would input their own total rows then re-using the amended sheet breaking functionality.
Achieving a Total row without using `VB Code` was not supported by any of the native Functions available in excel, particularly in automatically totalling only items from the last total row.

In order to achieve this functionality I realised the best way to achieve this would be to create flags that would identify the total row and items that needed totalling.
On the data entry page a column was added to provide a simple `boolean` flag:
```
IF any cell in the range of data entry CONTAINS "*total*" [case insensitive]
  THEN true
ELSE false
```

This flag is calculated for each row on the output sheets and a second column stores a count of the number of totals encountered up to each item:
```
IF totalFlag is true
  THEN counter is incremented by 1
ELSE counter equals its previous value [i.e. the row above]
```

A final flag is created to allocate items to the correct total.
```
IF totalFlag is true
  THEN blockFlag is null [Total rows are not included in any sum]
ELSE blockFlag is equal to the current counter
```

All columns holding data values have a short piece of code added to the beginning:
```
IF totalFlag is true
  THEN calculate the sum of all entries with blockFlag equal to current counter - 1
ELSE lookup current data value
```

## Depth Calculations
I'll try to represent this as simple as I can while remaining true to the complexity required by the limitations imposed.
The cost for drainage trenches is measured in price bands based upon their average depth.

For the first depth everything upto between the first and second depth is included in the first band.

While my sheet allows for this value to be customised typically the first depth is 0.50m with subsequent depths at 0.25m intervals resulting in the first band including depths up to 0.625m.  In my algorithm I will refer to the typical depths as this will keep it simpler.

*N.B. The first IF statement of each column contains the [totalFlag check](#The-Total-Row) described above which I will ignore.*

### First Depth (0.50m)
```
IF the lowest depth INCLUDING bed < 0.625m
  THEN IF the largest depth INCLUDING bed > 0.625m
          THEN the length of run at this depth = (0.625 - the lowest depth INC bed) * the gradient    [length per m fall] to the nearest 0.05m
       ELSE IF (the largest depth INC bed <= 0.625) AND (The largest depth != 0)
              THEN the length of run at this depth = the overall length of this run
            ELSE null
ELSE null
```

### Second Depth (0.75m) and Subsequent
This is essentially a switch in excel:
```
IF (the largest depth INC bed >= 0.875m) AND (the lowest depth INC bed <= 0.625m)
  THEN the length of run at this depth = 0.25m * the gradient
  
ELSE IF (the largest depth INC bed <= 0.875m) AND (the lowest depth INC bed >= 0.625m)
        THEN the length of run at this depth = the overall length of this run

ELSE IF (the largest depth INC bed > 0.625m) AND (the lowest depth INC bed <= 0.625m)
        THEN the length of run at this depth = (the largest depth INC bed - 0.625m) * the gradient
        
ELSE IF (the largest depth INC bed >= 0.875m) AND (the lowest depth INC bed < 0.875m)
        THEN the length of run at this depth = (0.875m - the smallest depth INC bed) * the gradient

ELSE null
```
## Junctions
Junctions come in various sizes and also different types and are often misplaced on a table.

To handle the various options I used a string search to flag items with a junction and concatenation junction type and joining pipe size.

To assign this to the correct run an identifier is created and included with the above.

Each column can then search this range of junction ids and count the number of junctions relating to each run.

## Settings
The settings are essentially a collection of constants that the user can amend as required.

In order to protect formula from users changing columns this involved an amended lookup using `INDEX` with `MATCH` to find the relative position of the column to be returned.  This allows users to customise column headings and even their order and maintain functionality.

