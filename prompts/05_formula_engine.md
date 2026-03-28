# OpenSheet Formula Engine Specification
Version: 1.0
Language: C++17
Component: office/engine/formula_engine

This document defines the formulas that must be implemented in the OpenSheet spreadsheet engine.

The formula engine must behave similar to Microsoft Excel.

Supported features:
- Cell references (A1, B2, C10)
- Cell ranges (A1:A10)
- Nested formulas
- Logical evaluation
- Range iteration
- Date handling
- Numeric formatting
- Boolean TRUE/FALSE values

---

# 1 Basic Mathematical Functions

## SUM
Syntax
SUM(range)

Example
=SUM(A1:A10)

Description
Returns the total of all numeric values in the range.

Implementation
Iterate through range and accumulate numeric values.

---

## AVERAGE

Syntax
AVERAGE(range)

Example
=AVERAGE(A1:A10)

Description
Returns arithmetic mean.

Implementation
sum(range) / count(numbers)

---

## COUNT

Syntax
COUNT(range)

Description
Counts numeric values only.

---

## MAX

Syntax
MAX(range)

Description
Returns largest numeric value.

---

## MIN

Syntax
MIN(range)

Description
Returns smallest numeric value.

---

## MEDIAN

Syntax
MEDIAN(range)

Description
Returns the middle value after sorting.

Implementation
Sort values then return middle.

---

# 2 Time & Date Functions

## TODAY

Syntax
TODAY()

Returns today's date.

Implementation
std::chrono::system_clock

---

## NOW

Syntax
NOW()

Returns current date and time.

---

## YEAR

Syntax
YEAR(date)

Returns year component.

Example
YEAR("2005-07-16") → 2005

---

## MONTH

Syntax
MONTH(date)

Returns month number (1-12)

---

## DAY

Syntax
DAY(date)

Returns day number.

---

## DATEDIF

Syntax
DATEDIF(start, end, unit)

Unit options
"Y" years
"M" months
"D" days

Example

=DATEDIF(A1, A2, "Y")

Implementation
Calculate difference between two dates.

---

# 3 Logical Functions

## IF

Syntax

IF(condition, value_if_true, value_if_false)

Example

=IF(A1>100,"High","Low")

Implementation
Evaluate logical expression.

---

## AND

Syntax

AND(condition1, condition2)

Returns TRUE if all conditions are true.

---

## OR

Syntax

OR(condition1, condition2)

Returns TRUE if at least one condition is true.

---

# 4 Lookup Functions

## VLOOKUP

Syntax

VLOOKUP(lookup_value, table_array, column_index, exact_match)

Example

=VLOOKUP(A2,B2:E10,3,FALSE)

Behavior

Search first column of table_array.

Return value from specified column.

If exact_match=false return only exact match.

---

# 5 Financial Functions

## PMT

Syntax

PMT(rate, periods, present_value, future_value, type)

Example

=PMT(0.05/12,60,10000)

Description

Calculates loan payment.

---

# 6 Statistical Functions

## SUMIF

Syntax

SUMIF(range, criteria, sum_range)

Example

=SUMIF(A1:A10,"Apples",B1:B10)

---

## SUMIFS

Syntax

SUMIFS(sum_range,criteria_range1,criteria1,...)

Example

=SUMIFS(B1:B10,A1:A10,"Apples",C1:C10,">10")

---

## COUNTIF

Syntax

COUNTIF(range,criteria)

Example

=COUNTIF(A1:A10,">5")

---

## COUNTIFS

Syntax

COUNTIFS(range1,criteria1,range2,criteria2)

---

## AVERAGEIF

Syntax

AVERAGEIF(range,criteria,average_range)

---

## AVERAGEIFS

Syntax

AVERAGEIFS(avg_range,range1,criteria1,range2,criteria2)

---

# 7 Comparison Operators

The formula parser must support:

=   equal
<   less than
>   greater than
<=  less than equal
>=  greater than equal
<>  not equal

---

# 8 Boolean Constants

TRUE
FALSE

---

# 9 Range Parsing

Examples

A1:A10
B2:D10
C1:C100

Engine must iterate through cells.

---

# 10 Nested Formulas

The engine must support nesting.

Example

=SUM(A1:A10)+AVERAGE(B1:B10)

---

# 11 Error Handling

Return following errors:

#DIV/0
#VALUE
#NAME
#REF
#NUM

---

# 12 Formula Parser Architecture

Required modules

