# kmc
<h3>VBA used:</h3>
<b>1. ConCat UDF (user defined function)</b><br>
Contributed by Rick Rothstein<br>
https://www.contextures.com/rickrothsteinexcelvbatext.html#About<br>
<br><!--reduce line height-->
Syntax:<br>
=ConCat("delimiter",["text"],[number],[cellLocation],[cellRanges])<br>
<br>
Examples:<br>
=ConCat("",A1:A3,C1,"HELLO",D1:D2)  //delimiter-empty string<br>
=ConCat(,A1:A3,C1,"HELLO",D1:D2)    //delimiter-empty string<br>
=ConCat("-",A1:A3,C1,"HELLO",D1:D2)<br>
<br>
<b>2. TripCount UDF (user defined function)</b><br>
Contributed by Bernie Deitrick<br>
https://answers.microsoft.com/en-us/msoffice/forum/all/excel-find-number-of-cells-within-a-merged-section/84d27fcb-1c0f-4405-8d62-e357743c813e<br>
https://answers.microsoft.com/en-us/profile/20889d9c-421e-4290-8bef-0de4837323b3?sort=LastReplyDate&dir=Desc&tab=Threads&forum=allcategories&meta=&status=&mod=&advFil=&postedAfter=undefined&postedBefore=undefined&threadType=All&page=1<br>
<br>
Syntax:<br>
=TripCount(A1:AE1,"Trip")<br>
<br>
<br>
<h3>Steps to Install UDF in your workbook.(Will not work in other workbooks)</h3>
1. Press Alt+F11 to go into the VB editor.<br>
2. Click Insert->Module from its menu bar.<br>
3. Then, Copy/Paste the code(present in file NameOfFunction_UDF.txt) into the code window that opened up. That's it.<br>
<br>
<h3>Using the ConCat UDF in Other Workbooks</h3><br>
If you install the ConCat UDF into a workbooks, then the function will travel with the file if you distribute it to others.<br>
<br>
If you find this ConCat UDF useful and want it available for use on any worksheets that only YOU will <br>
work on, just install it in your personal.xls file... just remember, though, if you install it to your <br>
personal.xls file and use if from there, then the function will NOT travel with any worksheets you<br>
distribute to others (meaning cells using ConCat will produce a #NAME! error on their computers) unless, <br>
of course, they install the function to their own personal.xls file as well.<br>
<br>
If you want to pursue the personal.xls file route, and you don't now have one, you can find out how to create one here...<br>
<br>
https://office.microsoft.com/en-us/excel-help/deploy-your-excel-macros-from-a-central-file-HA001087296.aspx<br>
<h3>Disadvantage:</h3><br>
1. You need to save your spreedsheet as a  macro-enabled spreadsheet(.xlsm)<br>
