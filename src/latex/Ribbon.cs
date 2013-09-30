using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;
using System.Collections;

namespace latex
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void tableButton_Click(object sender, RibbonControlEventArgs e)
        {
            Range range = Globals.ThisAddIn.Application.Selection;
            if (range.Count == 0)
            {
                MessageBox.Show("Selection is not valid!");
            }
            else
            {

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("\\begin{table}\\centering");
                sb.AppendLine("\\caption{Add caption here!}");
                sb.AppendLine("\\label{tbl:}");
                sb.AppendFormat("\\begin{{tabular}}{{{0}}}", new String('c', range.Columns.Count));
                sb.AppendLine();
                sb.AppendLine("\\hline");
                Range item;
                int headerRow = 1;
                ArrayList clineBegin = new ArrayList();
                ArrayList clineEnd = new ArrayList();
                for (int row = 1; row <= range.Rows.Count; ++row)
                {
                    for (int col = 1; col <= range.Columns.Count; ++col)
                    {
                        item = range[row, col];
                        if (item.MergeCells == false)
                        {
                            sb.Append(item.Text);
                            if (col != range.Columns.Count)
                            {
                                sb.Append("&");
                            }
                        }
                        else
                        {
                            if (item.Text != "")
                            {
                                char align = ((XlHAlign)item.MergeArea.HorizontalAlignment).ToString().ToLower()[8];
                                if (item.MergeArea.Rows.Count == 1) // Merge columns
                                {
                                    sb.AppendFormat("\\multicolumn{{{0}}}{{{1}}}{{{2}}}", item.MergeArea.Columns.Count, align, item.Text);
                                    if (row < headerRow)
                                    {
                                        clineBegin.Add(col);
                                        clineEnd.Add(col + item.MergeArea.Columns.Count - 1);
                                    }
                                }
                                else // Merge rows
                                {
                                    if (row == 1)
                                    {
                                        headerRow = item.MergeArea.Rows.Count;
                                    }
                                    sb.AppendFormat("\\multirow{{{0}}}{{*}}{{{1}}}", item.MergeArea.Rows.Count, item.Text);
                                }
                                if (item.MergeArea.Columns.Count + col - 1 < range.Columns.Count)
                                {
                                    sb.Append("&");
                                }
                            }
                            else
                            {
                                if ((item.MergeArea.Rows.Count > 1) && (col < range.Columns.Count))
                                {
                                    sb.Append("&");
                                }
                            }
                        }
                    }
                    sb.AppendLine("\\\\");
                    if ((row <= headerRow) || row == range.Rows.Count)
                    {
                        if (row < headerRow)
                        {
                            object[] begin = clineBegin.ToArray();
                            object[] end = clineEnd.ToArray();
                            for (int i = 0; i < clineBegin.Count; ++i)
                            {
                                sb.AppendFormat("\\cline{{{0}-{1}}}", begin[i].ToString(), end[i].ToString());
                            }
                            sb.AppendLine();
                            clineBegin.Clear();
                            clineEnd.Clear();
                        }
                        else
                        {
                            sb.AppendLine("\\hline");
                        }
                    }
                }
                sb.AppendLine("\\end{tabular}");
                sb.AppendLine("\\end{table}");
                if (clipboardBox.Checked == true)
                {
                    Clipboard.SetText(sb.ToString());
                }
                if (fileBox.Checked == true)
                {
                    SaveFileDialog fileDlg = new SaveFileDialog();
                    fileDlg.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                    fileDlg.RestoreDirectory = true;
                    if (fileDlg.ShowDialog() == DialogResult.OK)
                    {
                        StreamWriter sw = new StreamWriter(fileDlg.FileName);
                        sw.Write(sb.ToString());
                        sw.Flush();
                        sw.Close();
                    }
                }
            }
        }

        private void tableGroup_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(String.Format("{0}\n\n{1}\n\n{2}\n\n{3}", "LaTeXAddIn 1.0", 
                "Convert Excel data to LaTeX table environment.", "Copyright ©  2011 RiverStudio (Jun Ma). All rights reserved.",
                "The program is provided AS IS with NO WARRANTY OF ANY KIND, INCLUDING THE WARRANTY OF DESIGN, MERCHANTABLILITY AND FITNESS FOR A PARTICULAR PURPOSE."));
        }
    }
}
