using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FeisWork
{
    public partial class Form1 : Form
    {
        private const string GROUP1TITLE = "一、员工银行发放工资";
        private const string GROUP2TITLE = "二、员工现金发放工资";
        private const string GROUP3TITLE = "三、外来人员发放工资";

        private string filePath;
        private List<Employee> listGroup1Employes = new List<Employee>();
        private List<Employee> listGroup2Employes = new List<Employee>();
        private List<Employee> listGroup3Employes = new List<Employee>();

        public Form1()
        {
            InitializeComponent();

            this.cmdEmplyeeNameCol.SelectedIndex = 5;
            this.cmbSalaryCol.SelectedIndex = 33;
        }

        private void btnLoadFile_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "xlsx文件|*.xlsx";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.filePath = this.txtFilePath.Text = this.openFileDialog1.FileName;
                
            }
        }
        
        private Workbook LoadFile(string filePath)
        {
            try
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(filePath);

                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = workbook.Worksheets[i];
                    if (sheet.Name.Trim() == (i + 1).ToString() + "月份")
                    {
                        for (int r = sheet.FirstRow; r <= sheet.LastRow; r++)
                        {
                            string col1CellValue = sheet[r, 1].Value.Trim();
                            if (col1CellValue == GROUP1TITLE)
                            {
                                r = ReadGroupEmployee(sheet, r + 4, this.listGroup1Employes, i + 1);
                            }
                            else if (col1CellValue == GROUP2TITLE)
                            {
                                r = ReadGroupEmployee(sheet, r + 1, this.listGroup2Employes, i + 1);
                            }
                            else if (col1CellValue == GROUP3TITLE)
                            {
                                r = ReadGroupEmployee(sheet, r + 1, this.listGroup3Employes, i + 1);
                            }



                            //if (string.IsNullOrEmpty(sheet[r, 1].Value.Trim()))
                            //{
                            //    break;
                            //}

                            Console.WriteLine(sheet[r, 1].Value.Trim());
                        }
                    }
                }


                MessageBox.Show("导入完成");

                return workbook;
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
                return null;
            }
        }

        private string[] resultTableColHeader = new string[]
        {
            "员工编号",
            "员工姓名",
            "1月份",
            "2月份",
            "3月份",
            "4月份",
            "5月份",
            "6月份",
            "7月份",
            "8月份",
            "9月份",
            "10月份",
            "11月份",
            "12月份",
            "合计"
        };
        
        private void Save(Workbook workbook)
        {
            try
            {
                Worksheet sheet = workbook.Worksheets[workbook.Worksheets.Count - 1];
                if (sheet.Name != "工资汇总(自动生成)")
                {
                    sheet = workbook.CreateEmptySheet("工资汇总(自动生成)");
                }

                CellStyle cellStyle = sheet.GetDefaultRowStyle(1);
                cellStyle.VerticalAlignment = VerticalAlignType.Center;
                cellStyle.HorizontalAlignment = HorizontalAlignType.Center;

                for (int c = 1; c <= 15; c++)
                {
                    string cName = ColIndexToText(c);
                    sheet.Range[string.Format(cName + "{0}:" + cName + "{1}", 1, 3)].Merge();
                    sheet.Range[string.Format(cName + "{0}:" + cName + "{1}", 1, 3)].Text = resultTableColHeader[c - 1];

                }

                int row = 4;
                int g1EndRow = row = WriteGroupEmployee(sheet, listGroup1Employes, row, "银行发放小计");
                int g2EndRow = row = WriteGroupEmployee(sheet, listGroup2Employes, row + 1, "现金发放小计");

                row++;
                sheet.Range[string.Format("A{0}:B{0}", row)].Merge();
                sheet.Range[string.Format("A{0}:B{0}", row)].Text = "员工工资合计";

                for (int c = 0; c < 13; c++)
                {
                    string cName = ColIndexToText(c + 3);
                    sheet.Range[string.Format(cName + "{0}", row)].Value2 = string.Format("=SUM(" + cName + "{0}+" + cName + "{1})", g1EndRow, g2EndRow);
                }

                int g3EndRow = row = WriteGroupEmployee(sheet, listGroup3Employes, row + 1, "外来人员小计");
                row++;
                sheet.Range[string.Format("A{0}:B{0}", row)].Merge();
                sheet.Range[string.Format("A{0}:B{0}", row)].Text = "外账工资报表合计";

                for (int c = 0; c < 13; c++)
                {
                    string cName = ColIndexToText(c + 3);
                    sheet.Range[string.Format(cName + "{0}", row)].Value2 = string.Format("=SUM(" + cName + "{0}+" + cName + "{1})", g1EndRow, g3EndRow);
                }


                workbook.Save();
                MessageBox.Show("汇总成功");
            }
            catch (Exception exc)
            {
                MessageBox.Show("汇总失败。" + exc.Message);
            }
        }

        private int WriteGroupEmployee(Worksheet sheet, List<Employee> list, int row, string sumTitle)
        {
            int startRow = row;
            for (int i = 0; i < list.Count; i++)
            {
                sheet.Range[string.Format("A{0}", row)].Value = (i + 1).ToString();
                sheet.Range[string.Format("B{0}", row)].Text = list[i].Name;

                for (int c = 0; c < 12; c++)
                {
                    string cName = ColIndexToText(c + 3);
                    sheet.Range[string.Format(cName + "{0}", row)].Value = list[i].Salaries[c].ToString();
                }
                sheet.Range[string.Format("O{0}", row)].Value2 = string.Format("=SUM(C{0}:N{0})", row);

                row++;
            }


            sheet.Range[string.Format("A{0}:B{0}", row)].Merge();
            sheet.Range[string.Format("A{0}:B{0}", row)].Text = sumTitle;

            for (int c = 0; c < 13; c++)
            {
                string cName = ColIndexToText(c + 3);
                sheet.Range[string.Format(cName + "{0}", row)].Value2 = string.Format("=SUM(" + cName + "{0}:" + cName + "{1})", startRow, row - 1);
            }

            return row;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="colIndex">from 1</param>
        /// <returns></returns>
        private string ColIndexToText(int colIndex)
        {
            if (colIndex == 1)
            {
                return "A";
            }
            if (colIndex == 2)
            {
                return "B";
            }
            if (colIndex == 3)
            {
                return "C";
            }
            if (colIndex == 4)
            {
                return "D";
            }
            if (colIndex == 5)
            {
                return "E";
            }
            if (colIndex == 6)
            {
                return "F";
            }
            if (colIndex == 7)
            {
                return "G";
            }
            if (colIndex == 8)
            {
                return "H";
            }
            if (colIndex == 9)
            {
                return "I";
            }
            if (colIndex == 10)
            {
                return "J";
            }
            if (colIndex == 11)
            {
                return "K";
            }
            if (colIndex == 12)
            {
                return "L";
            }
            if (colIndex == 13)
            {
                return "M";
            }
            if (colIndex == 14)
            {
                return "N";
            }
            if (colIndex == 15)
            {
                return "O";
            }
            if (colIndex == 16)
            {
                return "P";
            }
            if (colIndex == 17)
            {
                return "Q";
            }
            if (colIndex == 18)
            {
                return "R";
            }
            if (colIndex == 19)
            {
                return "S";
            }
            if (colIndex == 20)
            {
                return "T";
            }
            if (colIndex == 21)
            {
                return "U";
            }
            return "Z";
        }

        private int ReadGroupEmployee(Worksheet sheet, int startRow, List<Employee> list, int month)
        {
            int r = startRow;
            for (; r <= sheet.LastRow; r++)
            {
                string name = sheet[r, this.cmdEmplyeeNameCol.SelectedIndex + 1].Value.Trim();
                if (string.IsNullOrEmpty(name))
                {
                    break;
                }
                int salary = 0;
                try
                {
                    object txtSalary = sheet[r, this.cmbSalaryCol.SelectedIndex + 1].FormulaValue;
                    salary = Convert.ToInt32(txtSalary);
                }
                catch (Exception)
                {
                    MessageBox.Show("表［" + sheet.Name + "]工资列 第" + r.ToString() + "行非数值，汇总失败请检查！");
                    break;
                }

                Employee emp = new Employee();

                emp = list.FirstOrDefault(e => e.Name == name);
                if (emp == null)
                {
                    emp = new Employee()
                    {
                        Name = name
                    };
                    list.Add(emp);
                }
                emp.Salaries[month - 1] = salary;


                Console.WriteLine(sheet[r, 1].Value.Trim());
            }

            return r;
        }

        private void btnLookUp_Click(object sender, EventArgs e)
        {
            this.listGroup1Employes.Clear();
            this.listGroup2Employes.Clear();
            this.listGroup3Employes.Clear();

            Workbook workbook = this.LoadFile(this.filePath);
            if (workbook != null)
            {
                this.Save(workbook);
            }
        }
    }

    public class Employee
    {
        public string Name;
        public int[] Salaries = new int[12];
    }
}
