# KatalonDiscussion9632

run Test Case `TC1` and see the log. 
You will find such log messages emitted by the TC1:

```
=TEXT(WORKDAY(TODAY(),1),"dd-mmm-yyyy"), 21-9-2018
=TEXT(WORKDAY(TODAY(),2),"dd-mmm-yyyy"), 18-9-2018
=TEXT(WORKDAY(TODAY(),3),"dd-mmm-yyyy"), 25-9-2018
=TEXT(WORKDAY(TODAY(),4),"dd-mmm-yyyy"), 26-9-2018
=TEXT(WORKDAY(TODAY(),5),"dd-mmm-yyyy"), 27-9-2018
=TEXT(WORKDAY(TODAY(),6),"dd-mmm-yyyy"), 28-9-2018
=TEXT(WORKDAY(TODAY(),7),"dd-mmm-yyyy"), 25-9-2018
=TEXT(WORKDAY(TODAY(),8),"dd-mmm-yyyy"), 02-10-2018
```

This message implies something is wrong in Katalon Studio.

I have doubt about com.kms.katalon.core.testdata.reader.SheetPOI class, internallyGetCellText() method

