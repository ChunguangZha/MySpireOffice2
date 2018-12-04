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

namespace MySpireOffice2
{
    public partial class Form1 : Form
    {
        string hostNoPrefix = "2302815900";
        Dictionary<string, Family> dicFamilies = new Dictionary<string, Family>();
        string lastHostName = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void btnLoadSrcTable5_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "xlsx文件|*.xlsx";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.txtSrcTable5FilePath.Text = this.openFileDialog1.FileName;
                this.LoadTable5(this.openFileDialog1.FileName);
            }
        }

        private void btnBuildTable4_Click(object sender, EventArgs e)
        {
            foreach (var family in dicFamilies.Values)
            {
                Workbook book = new Workbook();

                Worksheet sheet = book.Worksheets[0];
                sheet.DefaultRowHeight = 27;

                sheet.Name = family.hostHostName;
                sheet.Range["A1:G1"].Merge();
                sheet.Range["A1:G1"].Text = "村确认家庭人口调查表  （户口薄）";
                sheet.Range["A1:G1"].Style.Font.IsBold = true;
                sheet.Range["A1:G1"].Style.Font.Size = 18;
                sheet.Range["A1:G1"].Style.Font.FontName = "仿宋";
                sheet.Range["A1:G1"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["A1:G1"].Style.HorizontalAlignment = HorizontalAlignType.Center;

                int index = 0;
                foreach (var p in family.people.OrderBy(item => item.birthday).OrderBy(item => item.isHost))
                {
                    sheet.Range[string.Format("A{0}:A{1}", (index * 4) + 2, (index * 4) + 5)].Merge();
                    sheet.Range[string.Format("A{0}:A{1}", (index * 4) + 2, (index * 4) + 5)].Text = p.isHost ? "户主情况" : "家庭人员";
                    sheet.Range[string.Format("B{0}", (index * 4) + 2)].Text = "姓名";
                    sheet.Range[string.Format("C{0}", (index * 4) + 2)].Text = p.isHost ? p.hostName : p.name;
                    sheet.Range[string.Format("D{0}", (index * 4) + 2)].Text = "性别";
                    sheet.Range[string.Format("E{0}", (index * 4) + 2)].Text = p.sex;
                    sheet.Range[string.Format("F{0}", (index * 4) + 2)].Text = p.isHost ? "户编号" : "与户主关系";
                    sheet.Range[string.Format("G{0}", (index * 4) + 2)].Text = p.isHost ? family.hostNo : p.relation;

                    sheet.Range[string.Format("B{0}", (index * 4) + 3)].Text = "身份证号";
                    sheet.Range[string.Format("C{0}:E{1}", (index * 4) + 3, (index * 4) + 3)].Merge();
                    sheet.Range[string.Format("C{0}:E{1}", (index * 4) + 3, (index * 4) + 3)].Text = p.idNo.Substring(0, 18);
                    sheet.Range[string.Format("F{0}", (index * 4) + 3)].Text = "民族";
                    sheet.Range[string.Format("G{0}", (index * 4) + 3)].Text = p.nation;

                    sheet.Range[string.Format("B{0}", (index * 4) + 4)].Text = "现居住地址";
                    sheet.Range[string.Format("C{0}:E{1}", (index * 4) + 4, (index * 4) + 4)].Merge();
                    sheet.Range[string.Format("C{0}:E{1}", (index * 4) + 4, (index * 4) + 4)].Text = p.location;
                    sheet.Range[string.Format("F{0}", (index * 4) + 4)].Text = "婚姻状况";
                    sheet.Range[string.Format("G{0}", (index * 4) + 4)].Text = "";

                    sheet.Range[string.Format("B{0}", (index * 4) + 5)].Text = "备注";
                    sheet.Range[string.Format("C{0}:G{1}", (index * 4) + 5, (index * 4) + 5)].Merge();

                    index++;
                }
                
                CellStyle cellStyle = sheet.GetDefaultRowStyle(1);
                cellStyle.Font.Size = 14;
                cellStyle.Font.FontName = "仿宋";
                cellStyle.Font.IsBold = false;
                cellStyle.VerticalAlignment = VerticalAlignType.Center;
                cellStyle.HorizontalAlignment = HorizontalAlignType.Center;
                for (int i = sheet.FirstRow + 1; i <= sheet.LastRow; i++)
                {
                    sheet.SetDefaultRowStyle(i, cellStyle);
                }

                for (int i = 0; i < family.people.Count; i++)
                {
                    sheet.Range[string.Format("A{0}:G{1}", (i * 4) + 2, (i * 4) + 5)].BorderInside(LineStyleType.Thin, ExcelColors.Black);
                    sheet.Range[string.Format("A{0}:G{1}", (i * 4) + 2, (i * 4) + 5)].BorderAround(LineStyleType.Medium, ExcelColors.Black);
                }

                sheet.SetRowHeight(1, 69);
                sheet.SetColumnWidth(1, 4.71);
                sheet.SetColumnWidth(2, 14.71);
                sheet.SetColumnWidth(3, 13.71);
                sheet.SetColumnWidth(4, 11.14);
                sheet.SetColumnWidth(5, 10.71);
                sheet.SetColumnWidth(6, 14.71);
                sheet.SetColumnWidth(7, 12.86);

                sheet.Range["A2:A" + sheet.LastRow].Style.WrapText = true;

                book.SaveToFile("Output\\" + family.hostHostName + "_" + family.hostNo + ".xlsx", ExcelVersion.Version2010);

            }

            MessageBox.Show("Save OK");
        }

        private void LoadTable5(string filePath)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filePath);

            Worksheet sheet = workbook.Worksheets[0];

            Family family = null;

            string hostName = "";
            for (int r = sheet.FirstRow + 2; r <= sheet.LastRow; r++)
            {
                string hostNo = sheet[r, 1].Value.Trim();
                hostName = sheet[r, 2].Value.Trim();
                string liveState = sheet[r, 14].Value.Trim();
                if (liveState == "死亡")
                {
                    continue;
                }

                //if (hostNo.Length > 5)
                //{
                //    hostNoPrefix = hostNo.Substring(0, hostNo.Length - 5);
                //}
                //else if (!string.IsNullOrEmpty(hostNo))
                //{
                //    hostNo = hostNoPrefix + hostNo;
                //}

                if (this.lastHostName != hostName)
                {
                    family = new Family()
                    {
                        hostNo = hostNo,
                        hostHostName = hostName
                    };

                    this.lastHostName = hostName;
                    this.dicFamilies.Add(hostName, family);
                }

                Person p = new Person()
                {
                    hostNo = hostNo,
                    hostName = hostName,
                    name = sheet[r, 3].Value.Trim(),
                    relation = sheet[r, 4].Value.Trim(),
                    idNo = sheet[r, 5].Value.Trim().Substring(0,18),
                    birthday = new DateTime(int.Parse(sheet[r, 5].Value.Trim().Substring(6, 4)), int.Parse(sheet[r, 5].Value.Trim().Substring(10, 2)), int.Parse(sheet[r, 5].Value.Trim().Substring(12, 2))),
                    sex = sheet[r, 7].Value.Trim(),
                    huY_renY = sheet[r, 8].Value.Trim(),
                    huY_renN = sheet[r, 9].Value.Trim(),
                    huN_renY = sheet[r, 10].Value.Trim(),
                    huN_renN = sheet[r, 11].Value.Trim(),
                    nation = "",
                    group = sheet[r, 12].Value.Trim(),
                    isTuDiChengbao = sheet[r, 13].Value.Trim(),
                    lifeState = liveState,
                    marryState = sheet[r, 15].Value.Trim(),
                    location = "",
                    education = "",
                    job = ""
                };
                p.isHost = p.hostName == p.name;

                family.people.Add(p);

            }

            MessageBox.Show("导入成功！");
        }

        private void OrderFamily()
        {
            foreach (var family in this.dicFamilies.Values)
            {
                this.getHostNo(family);
                family.people.OrderBy(p => p.birthday).OrderBy(p => p.isHost); ;
            }
        }

        private string getHostNo(Family family)
        {
            string hostNo = "";
            foreach (var person in family.people)
            {
                if (!string.IsNullOrEmpty(person.hostNo))
                {
                    hostNo = person.hostNo;
                }
            }

            if (hostNo.Length == 5)
            {
                hostNo = this.hostNoPrefix + hostNo;
            }
            family.hostNo = hostNo;
            return hostNo;
        }
    }

    class Family
    {
        public string hostNo;
        public string hostHostName;

        public List<Person> people = new List<Person>();
    }

    class Person
    {
        public bool isHost;

        public string hostName;

        public string hostNo;

        public string name;

        public string relation;

        public string sex;

        public string nation;

        public string idNo;

        public DateTime birthday;

        public string group;

        public string location;

        public string education;

        public string job;

        public string huY_renY;

        public string huY_renN;

        public string huN_renY;

        public string huN_renN;

        public string isTuDiChengbao;

        public string lifeState;

        public string marryState;


    }
}