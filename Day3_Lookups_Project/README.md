# 📊 Lookup & Conditional Functions

## 📌 Overview

This repository contains my **Day 3 Excel practice** focused on lookup functions, conditional logic, error handling, absolute references, and practical data analysis.

The purpose of this practice was to understand how Excel can be used to search for information, retrieve related values, handle errors, apply multiple conditions, and solve real-world data analysis problems.

---

# 🧠 Topics Covered

## 1. XLOOKUP()

### Definition

`XLOOKUP()` searches for a value in one range and returns the corresponding value from another range.

### Syntax

```excel
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])

EXCEL DAY 3 – ALL FORMULAS
==========================

1. XLOOKUP
-----------

Basic XLOOKUP:

=XLOOKUP(lookup_value,lookup_array,return_array)

XLOOKUP with "Not Found":

=XLOOKUP(lookup_value,lookup_array,return_array,"Not Found")

Example:

=XLOOKUP(A2,A5:A10,B5:B10,"Not Found")


2. VLOOKUP
----------

Basic exact-match VLOOKUP:

=VLOOKUP(lookup_value,table_array,column_number,FALSE)

Example:

=VLOOKUP(G25,$G$20:$K$27,2,FALSE)

Another example:

=VLOOKUP(G22,$G$20:$J$27,4,FALSE)


3. HLOOKUP
----------

Basic exact-match HLOOKUP:

=HLOOKUP(lookup_value,table_array,row_number,FALSE)

Example:

=HLOOKUP("Mar",N27:R30,2,FALSE)

Another example:

=HLOOKUP("May",N27:R30,3,FALSE)


4. INDEX
--------

INDEX with one column:

=INDEX(O14:O18,3)

INDEX with one column:

=INDEX(Q14:Q18,4)

INDEX with rows and columns:

=INDEX(N14:Q18,4,2)

Another example:

=INDEX(N14:Q18,5,4)

Employee dataset example:

=INDEX(A2:E9,5,5)


5. MATCH
--------

Exact-match MATCH:

=MATCH(lookup_value,lookup_array,0)

Example:

=MATCH(W6,W2:W7,0)

Another example:

=MATCH("Rahul",W2:W7,0)


6. INDEX + MATCH
---------------

Basic INDEX + MATCH:

=INDEX(return_range,MATCH(lookup_value,lookup_range,0))

Example:

=INDEX(D2:D7,MATCH(G2,B2:B7,0))

Using absolute references:

=INDEX($D$2:$D$7,MATCH(G2,$B$2:$B$7,0))

INDEX + MATCH with a complete table:

=INDEX(B14:D19,MATCH(B15,B14:B19,0),3)

Another example:

=INDEX(B14:C19,MATCH(B18,B14:B19,0),2)

Another example:

=INDEX(B14:D19,MATCH(B19,B14:B19,0),3)


7. IFERROR
----------

Basic IFERROR:

=IFERROR(value,value_if_error)

Example:

=IFERROR(10/0,"ERROR")

XLOOKUP + IFERROR:

=IFERROR(XLOOKUP(N29,N27:N31,O27:O31),"ERROR")

VLOOKUP + IFERROR:

=IFERROR(VLOOKUP(N30,N27:P31,3,FALSE),"ERROR")

INDEX + MATCH + IFERROR:

=IFERROR(INDEX(N27:P31,MATCH(N31,N27:N31,0),3),"ERROR")

INDEX + MATCH + IFERROR with custom message:

=IFERROR(INDEX(N27:P31,MATCH(N31,N27:N31,0),3),"Not Found")


8. NESTED IF
------------

Basic Nested IF:

=IF(B2>=70000,"Excellent",IF(B2>=50000,"Good","Needs Improvement"))

Day 3 project performance formula:

=IF(D2>=80000,"Excellent",IF(D2>=60000,"Good","Needs Improvement"))


9. ABSOLUTE REFERENCES
----------------------

Absolute cell reference:

=$A$1

Absolute range:

=$A$1:$D$10

VLOOKUP with fixed table range:

=VLOOKUP(G25,$G$20:$K$27,2,FALSE)

INDEX + MATCH with fixed ranges:

=INDEX($D$2:$D$7,MATCH(G2,$B$2:$B$7,0))


10. DAY 3 PROJECT FORMULAS
--------------------------

Employee Name using XLOOKUP:

=IFERROR(XLOOKUP(G2,$A$2:$A$9,$B$2:$B$9),"Not Found")


Department using VLOOKUP:

=IFERROR(VLOOKUP(G2,$A$2:$E$9,3,FALSE),"Not Found")


Salary using INDEX + MATCH:

=IFERROR(INDEX($E$2:$E$9,MATCH(G2,$A$2:$A$9,0)),"Not Found")


Performance using Nested IF:

=IF(D2>=80000,"Excellent",IF(D2>=60000,"Good","Needs Improvement"))


HLOOKUP – Sales:

=HLOOKUP(G2,$B$13:$F$16,2,FALSE)


HLOOKUP – Profit:

=HLOOKUP(G2,$B$13:$F$16,3,FALSE)


INDEX example:

=INDEX(A2:E9,5,5)


11. IMPORTANT FORMULA PATTERNS
-----------------------------

XLOOKUP:

=XLOOKUP(what_to_find,where_to_search,what_to_return)


VLOOKUP:

=VLOOKUP(what_to_find,table,return_column,FALSE)


HLOOKUP:

=HLOOKUP(what_to_find,table,return_row,FALSE)


INDEX:

=INDEX(range,row,column)


MATCH:

=MATCH(what_to_find,where_to_search,0)


INDEX + MATCH:

=INDEX(return_range,MATCH(what_to_find,lookup_range,0))


IFERROR:

=IFERROR(formula,"message")


Nested IF:

=IF(condition1,result1,IF(condition2,result2,result3))


12. QUICK REMINDER
------------------

XLOOKUP  → Search and return a value

VLOOKUP  → Vertical lookup

HLOOKUP  → Horizontal lookup

INDEX    → Return value from a position

MATCH    → Find position

INDEX + MATCH → Find position + return value

IFERROR  → Replace errors with a custom result

Nested IF → Handle multiple conditions

$        → Lock/fix a cell or range
