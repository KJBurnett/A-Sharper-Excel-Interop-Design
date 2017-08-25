# A Sharper Excel Interop Design
Understanding how to use Excel Interop for spreadsheet manipulation in C# without hours of googling.

## How to add Excel Interop to your project

1. Create your project
2. In the Solution Explorer, right click on your project -> Add -> Reference...
3. On the left-hand side, select the COM tab.
4. In the "Search COM" box in the top right, search for "excel"
5. One of your options should be "Microsoft Excel version_number Object Library". In my case version_number = 16.0.
6. Check the box and press OK.

You have successfully added the Excel Interop library to your project! Now, all you have to do is add this using statement to your file:

```c#
using Excel = Microsoft.Office.Interop.Excel;
```
