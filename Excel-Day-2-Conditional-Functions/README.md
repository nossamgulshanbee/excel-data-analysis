METRICS
-------------------------------------------------------
Total Employees

=COUNTA(Sales_Data!A2:A11)
-------------------------------------------------------
South Employees

=COUNTIF(Sales_Data!B2:B11,"South")
-------------------------------------------------------
West Employees

=COUNTIF(Sales_Data!B2:B11,"West")
-------------------------------------------------------
Employees with Sales ≥ ₹70,000

=COUNTIF(Sales_Data!C2:C11,">=70000")
-------------------------------------------------------
South Employees with Sales ≥ ₹50,000

=COUNTIFS(Sales_Data!B2:B11,"South",Sales_Data!C2:C11,">=50000")
-------------------------------------------------------
Total South Sales

=SUMIF(Sales_Data!B2:B11,"South",Sales_Data!C2:C11)
-------------------------------------------------------
Total West Sales

=SUMIF(Sales_Data!B2:B11,"West",Sales_Data!C2:C11)
-------------------------------------------------------
Total North Sales ≥ ₹70,000

=SUMIFS(Sales_Data!C2:C11,Sales_Data!B2:B11,"North",Sales_Data!C2:C11,">=70000")
-------------------------------------------------------
Average South Sales

=AVERAGEIF(Sales_Data!B2:B11,"South",Sales_Data!C2:C11)
-------------------------------------------------------
Average West Sales with Rating ≥ 3

=AVERAGEIFS(Sales_Data!C2:C11,Sales_Data!B2:B11,"West",Sales_Data!D2:D11,">=3")
-------------------------------------------------------
Total North Sales

=AVERAGEIF(Sales_Data!B2:B11,"North",Sales_Data!C2:C11)
-------------------------------------------------------
Average West Sales

=AVERAGEIF(Sales_Data!B2:B11,"West",Sales_Data!C2:C11)
-------------------------------------------------------
Count of employee eligibility

=COUNTIF(Sales_Data!G2:G11,"Eligible")
-------------------------------------------------------
Count of excellent performance

=COUNTIF(Sales_Data!F2:F11,"Excellent")
-------------------------------------------------------
percentage of eligibity

=COUNTIF(Sales_Data!G2:G11,"Eligible")/COUNTA(Sales_Data!G2:G11)*100
-------------------------------------------------------
Sales ≥ ₹50,000 AND (South OR West) AND Rating ≥ 4

=COUNTIFS(Sales_Data!C2:C11,">=50000",Sales_Data!B2:B11,"South",Sales_Data!D2:D11,">=4")+COUNTIFS(Sales_Data!C2:C11,">=50000",Sales_Data!B2:B11,"West",Sales_Data!D2:D11,">=4")
