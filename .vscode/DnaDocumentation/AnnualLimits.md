# Annual Limits Functions

This module contains procedures used to return various contribution and salary limits.

*CFGENA has following remarks. Not sure if we are implementing them or if we need them?*

Results: This function returns 9.999999e19 (a very large value) as the value of the limit if no limit was in effect for that year. For example, this would be the result for Medicare Wage Base for years after 1993.

The function returns an error when the limit is not available. This could mean that: (a) the law implementing the limit occurred after the year you specified; (b) the function does not contain limits as far back as the year you specified; or (c) the law was revised to the extent that the limit no longer applies (as in the case of HCE upper and lower limits).

You can make the function return n/a (as a string, not the N/A Excel error), when a limit is not available by placing an exclamation point in front of Name. With an exclamation point, the result does not produce a CFGENA error when the limit is not available.

Note that rounding also enables a according to the law mode. That is, according to the law, many of the rounded values are not allowed to be less than the value for some limit-specific year. This becomes an issue when, like in 2010, the actual COLA was less than zero and produced unrounded limits that were less than those for 2009. This means, that the function will return the same values (for many limits) for 2010 as it does for 2009. If you ask the function to produce unrounded numbers, this extra test is not done.

Projecting limits: projection is done from the unrounded limit then the appropriate rounding is applied.

By placing an asterisk in front of Name (within the quotes), the unrounded limit is returned (if applicable). For example, AnnualLimit (*415DBLimit, 2001,0.0,2001) would yield $141,075 rather than $140,000.

Click one of the links in the following list to see detailed help about the function.

Function | Description
---|---
{FUNCTIONS}

[Back to RBLe Framework](RBLe.md)