using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Diagnostics;
using System.Net;
using System.Windows.Forms;

namespace ExcelAutoPaste
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            ClipboardNotification.ClipboardUpdate += ClipboardNotification_ClipboardUpdate;
        }

        private void ClipboardNotification_ClipboardUpdate(object sender, EventArgs e)
        {
            if (toggleReceive.Checked)
            {
                try
                {
                    Range rng = (Range)Globals.ThisAddIn.Application.ActiveCell;
                    object cellValue = rng.Value;
                    int row = rng.Row;
                    int column = rng.Column;
                    if (Clipboard.ContainsText())
                    {
                        IDataObject iData = Clipboard.GetDataObject();
                        Worksheet activeSheet = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                        string clipboardData = (string)iData.GetData(DataFormats.Text);
                        activeSheet.Paste();
                        Range newCell;
                        int numLines;
                        if (pasteDirection.SelectedItemIndex == 0)
                        {
                            numLines = clipboardData.Split('\n').Length;
                            if (numLines < 1)
                                numLines = 1;
                            newCell = (Range)activeSheet.Cells[row + numLines, column];
                        }
                        else
                        {
                            numLines = clipboardData.Split('\t').Length;
                            if (numLines < 1)
                                numLines = 1;
                            newCell = (Range)activeSheet.Cells[row, column + numLines];
                        }
                        newCell.Select();
                    }
                }
                catch (Exception ex)
                {
                    Debug.Print(ex.ToString());
                }
            }

        }


        private void ToggleReceive_Click(object sender, RibbonControlEventArgs e)
        {   
            if (toggleReceive.Checked)
            {
                toggleReceive.Label = "Watching Clipboard!";
                Debug.Print("Listening...");
            }
            else
            {
                toggleReceive.Label = "Watch clipboard";
                Debug.Print("Stopping...");
            }
        }

        private void pasteDirection_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Debug.Print(pasteDirection.SelectedItemIndex.ToString());
        }
    }
}
