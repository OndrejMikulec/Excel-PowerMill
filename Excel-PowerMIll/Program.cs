/*
 * Created by SharpDevelop.
 * 
The MIT License (MIT)

Copyright (c) 2015 Ondrej Mikulec
o.mikulec@seznam.cz
Vsetin, Czech Republic

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
 */
using System;

namespace Excel_PowerMIll
{
	class Program
	{
		public static void Main(string[] args)
		{
			PowerMILL.Application pMAppliacation = (PowerMILL.Application) System.Runtime.InteropServices.Marshal.GetActiveObject("PowerMill.Application");
			
			Microsoft.Office.Interop.Excel.Application exApplication = new Microsoft.Office.Interop.Excel.Application();
			Microsoft.Office.Interop.Excel.Workbook exWorkbook = exApplication.Workbooks.Open(@"e:\ExTest.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t",false, false, 0, true, 1, 0);
			Microsoft.Office.Interop.Excel._Worksheet exWorksheet = (Microsoft.Office.Interop.Excel._Worksheet)exWorkbook.Sheets[1];
			Microsoft.Office.Interop.Excel.Range exRange = exWorksheet.UsedRange;
			
			int rowCount = exRange.Rows.Count;
			int colCount = exRange.Columns.Count;
			
			for (int i = 1; i <= rowCount; i++) {
				for (int j = 1; j <= colCount; j++) {
					string valueForPM = null;
					try {
						valueForPM = (string)(exRange.Cells[i, j] as Microsoft.Office.Interop.Excel.Range).Value.ToString();
					} catch {}
					
					if (valueForPM != null) {
						pMAppliacation.DoCommand(@"MESSAGE INFO """+ valueForPM +@"""");
					}
				}
			}
		}
	}
}