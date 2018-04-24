# AQS
## about_AQS

# SHORT DESCRIPTION
Advanced Query Syntax (AQS) is a language that we can use to filter EWS results.
This document describes basic functionality offered by it.

# LONG DESCRIPTION
Anytime filtering is required in Exchange Web Services (EWS), we can use 
Advanced Query Syntax (AQS). This filtering language is pretty natural and 
in general - easy to use.

Syntax for a single property search:
```
property:value
otherProperty:"Value with spaces"
```

For some values (like dates or sizes) we can also use compare operators:
```
size:>=100000
```

Finally, we can use range operator:
```
size:700..800
```

For dates we can use keywords (today, tomorrow, coming year)

We can also use brackets and logical AND and OR operators:
```
from:(Adam or Ewa) AND subject:paradise
```



# EXAMPLES
```powershell
Get-EWSItem -Filter subject:PowerShell
```
Search for any item with PowerShell in the subject.

```powershell
Get-EWSItem -Filter isread:false
```
Search for items that were not read.


# SEE ALSO
https://msdn.microsoft.com/en-us/library/office/dn579420(v=exchg.150).aspx

# KEYWORDS

- AQS
- Advanced Query Syntax
