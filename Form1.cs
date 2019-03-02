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

        string inputSymRight = "";
        string symbolSquareRight = "";
        string symbolSquareError = "";
        string symbolSquareNull = "";
        string symbolRight = "";

        List<Person> listperson = new List<Person>();

        public Form1()
        {
            InitializeComponent();
        }

        private void btnLoadSrcTablePeopleInfo_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "xlsx文件|*.xlsx";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.txtPeopleInfoTablePath.Text = this.openFileDialog1.FileName;
                this.LoadTablePeopleInfo(this.openFileDialog1.FileName);
            }
        }


        private void btnLoad户籍信息表_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "xlsx文件|*.xlsx";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.txtSrcTable5FilePath.Text = this.openFileDialog1.FileName;
                this.Load户籍信息表(this.openFileDialog1.FileName);
            }
        }


        private void btnBuildTable4家庭人员调查表_Click(object sender, EventArgs e)
        {
            foreach (var family in dicFamilies.Values)
            {
                Workbook book = new Workbook();

                Worksheet sheet = book.Worksheets[0];

                sheet.Name = family.hostHostName;
                sheet.Range["A1:G1"].Merge();
                sheet.Range["A1:G1"].Text = "村确认家庭人口调查表  （户口薄）";
                sheet.Range["A1:G1"].Style.Font.IsBold = true;
                sheet.Range["A1:G1"].Style.Font.Size = 14;
                sheet.Range["A1:G1"].Style.Font.FontName = "仿宋";
                sheet.Range["A1:G1"].Style.VerticalAlignment = VerticalAlignType.Center;
                sheet.Range["A1:G1"].Style.HorizontalAlignment = HorizontalAlignType.Center;
                
                this.getHostNo(family);

                int index = 0;
                foreach (var p in family.people.OrderBy(item => item.birthday).OrderByDescending(item => item.isHost))
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
                    sheet.Range[string.Format("C{0}:E{1}", (index * 4) + 4, (index * 4) + 4)].Text = string.IsNullOrEmpty(p.location) ? "讷河市龙河镇国庆村" : p.location;
                    sheet.Range[string.Format("F{0}", (index * 4) + 4)].Text = "婚姻状况";
                    sheet.Range[string.Format("G{0}", (index * 4) + 4)].Text = p.marryState;

                    sheet.Range[string.Format("B{0}", (index * 4) + 5)].Text = "备注";
                    sheet.Range[string.Format("C{0}:G{1}", (index * 4) + 5, (index * 4) + 5)].Merge();

                    index++;
                }
                
                CellStyle cellStyle = sheet.GetDefaultRowStyle(1);
                cellStyle.Font.Size = 12;
                cellStyle.Font.FontName = "仿宋";
                cellStyle.Font.IsBold = false;
                cellStyle.VerticalAlignment = VerticalAlignType.Center;
                cellStyle.HorizontalAlignment = HorizontalAlignType.Center;
                for (int i = sheet.FirstRow + 1; i <= sheet.LastRow; i++)
                {
                    sheet.SetDefaultRowStyle(i, cellStyle);
                    sheet.SetRowHeight(i, 15);
                }

                for (int i = 0; i < family.people.Count; i++)
                {
                    sheet.Range[string.Format("A{0}:G{1}", (i * 4) + 2, (i * 4) + 5)].BorderInside(LineStyleType.Thin, ExcelColors.Black);
                    sheet.Range[string.Format("A{0}:G{1}", (i * 4) + 2, (i * 4) + 5)].BorderAround(LineStyleType.Medium, ExcelColors.Black);
                }

                sheet.SetRowHeight(1, 69);
                sheet.SetColumnWidth(1, 4.71);
                sheet.SetColumnWidth(2, 12.86);
                sheet.SetColumnWidth(3, 11.86);
                sheet.SetColumnWidth(4, 6.71);
                sheet.SetColumnWidth(5, 7.29);
                sheet.SetColumnWidth(6, 15.29);
                sheet.SetColumnWidth(7, 22.29);

                sheet.Range["A2:A" + sheet.LastRow].Style.WrapText = true;

                book.SaveToFile("家庭人员调查表\\" + family.hostHostName + "_" + family.hostNo + ".xlsx", ExcelVersion.Version2010);

            }

            MessageBox.Show("Save OK");
        }


        private void btnBuild3Table人口摸底调查表_Click(object sender, EventArgs e)
        {
            foreach (var family in dicFamilies.Values)
            {
                Workbook book = new Workbook();

                Worksheet sheet = book.Worksheets[0];

                sheet.Name = family.hostHostName;
                this.getHostNo(family);

                const int itemRowCount = 13;
                int itemIndex = 0;

                foreach (var p in family.people.OrderBy(item => item.birthday).OrderByDescending(item => item.isHost))
                {
                    int row = itemIndex * itemRowCount + 1;
                    sheet.Range[string.Format("A{0}:H{1}", row, row)].Merge();
                    sheet.Range[string.Format("A{0}:H{1}", row, row)].Text = "人员摸底调查表";
                    sheet.Range[string.Format("A{0}:H{1}", row, row)].Style.Font.IsBold = true;
                    sheet.Range[string.Format("A{0}:H{1}", row, row)].Style.Font.Size = 12;
                    sheet.Range[string.Format("A{0}:H{1}", row, row)].Style.Font.FontName = "仿宋";
                    sheet.Range[string.Format("A{0}:H{1}", row, row)].Style.VerticalAlignment = VerticalAlignType.Center;
                    sheet.Range[string.Format("A{0}:H{1}", row, row)].Style.HorizontalAlignment = HorizontalAlignType.Center;

                    row += 1;
                    sheet.Range[string.Format("A{0}:H{1}", row, row)].Merge();
                    sheet.Range[string.Format("A{0}:H{1}", row, row)].Text = "      国庆村     " + txtGroup.Text + "屯                       调查时间:2018年12月";

                    row += 1;
                    sheet.Range[string.Format("A{0}", row)].Text = "姓 名";
                    sheet.Range[string.Format("B{0}", row)].Text = p.name;
                    sheet.Range[string.Format("C{0}", row)].Text = "性 别";
                    sheet.Range[string.Format("D{0}", row)].Text = p.sex;
                    sheet.Range[string.Format("E{0}", row)].Text = "民 族";
                    sheet.Range[string.Format("F{0}", row)].Text = p.nation;
                    sheet.Range[string.Format("G{0}", row)].Text = "出生日期";
                    sheet.Range[string.Format("H{0}", row)].Text = p.birthday.ToShortDateString();

                    row += 1;
                    sheet.Range[string.Format("A{0}", row)].Text = "学 历";
                    sheet.Range[string.Format("B{0}", row)].Text = p.education;
                    sheet.Range[string.Format("C{0}:D{1}", row, row)].Merge();
                    sheet.Range[string.Format("C{0}:D{1}", row, row)].Text = "身份证号";
                    sheet.Range[string.Format("E{0}:F{1}", row, row)].Merge();
                    sheet.Range[string.Format("E{0}:F{1}", row, row)].Text = p.idNo;
                    sheet.Range[string.Format("G{0}", row)].Text = "联系电话";
                    sheet.Range[string.Format("H{0}", row)].Text = p.phone;

                    row += 1;
                    sheet.Range[string.Format("A{0}", row)].Text = "兵役状况";
                    sheet.Range[string.Format("B{0}:D{1}", row, row)].Merge();
                    sheet.Range[string.Format("B{0}:D{1}", row, row)].Text = "年月日至年月日";
                    sheet.Range[string.Format("E{0}:G{1}", row, row)].Merge();
                    sheet.Range[string.Format("E{0}:G{1}", row, row)].Text = "对本集体经济组织特殊贡献情况";
                    sheet.Range[string.Format("H{0}", row)].Text = "（有/无）无";

                    row += 1;
                    sheet.Range[string.Format("A{0}", row)].Text = "婚姻状况";
                    sheet.Range[string.Format("B{0}:H{1}", row, row)].Merge();

                    string merryStateText = (p.marryState == "未婚" ? this.symbolSquareRight : this.symbolSquareNull) + "未婚 " +
                                            (p.marryState == "已婚" ? this.symbolSquareRight : this.symbolSquareNull) + "已婚 " +
                                            (p.marryState == "离异" ? this.symbolSquareRight : this.symbolSquareNull) + "离异 " +
                                            (p.marryState == "丧偶" ? this.symbolSquareRight : this.symbolSquareNull) + "丧偶 " +
                                            "  婚姻状况变动日期：    年   月    日";

                    sheet.Range[string.Format("B{0}:H{1}", row, row)].Text = merryStateText;

                    row += 1;
                    sheet.Range[string.Format("A{0}:B{1}", row, row)].Merge();
                    sheet.Range[string.Format("A{0}:B{1}", row, row)].Text = "取得家庭承包地情况";

                    string jiatingchengbaodiqingkuang = (p.isTuDiChengbao == "是" ? this.symbolSquareRight + "是；" + this.symbolSquareNull + "否" : this.symbolSquareNull + "是；" + this.symbolSquareRight + "否") +
                                                        "  原因：" +
                                                        (p.lifeState == "新生" ? this.symbolRight : "") + "1、新生" + " " +
                                                        (p.lifeState == "婚入" ? this.symbolRight : "") + "2、婚入" + " " +
                                                        (p.lifeState == "世居" ? this.symbolRight : "") + "3、世居" + " " +
                                                        (p.lifeState != "新生" && p.lifeState != "婚入" && p.lifeState != "世居" ? this.symbolRight : "") + "4、其他   ";
                    sheet.Range[string.Format("C{0}:H{1}", row, row)].Merge();
                    sheet.Range[string.Format("C{0}:H{1}", row, row)].Text = jiatingchengbaodiqingkuang;

                    row += 1;
                    sheet.Range[string.Format("A{0}:B{1}", row, row)].Merge();
                    sheet.Range[string.Format("A{0}:B{1}", row, row)].Text = "户口性质变动情况";

                    string hukouxingzhibiandong = "       年  月  日 因                    转为非农业";
                    sheet.Range[string.Format("C{0}:H{1}", row, row)].Merge();
                    sheet.Range[string.Format("C{0}:H{1}", row, row)].Text = hukouxingzhibiandong;

                    row += 1;
                    sheet.Range[string.Format("A{0}:B{1}", row, row)].Merge();
                    sheet.Range[string.Format("A{0}:B{1}", row, row)].Text = "户籍地变动情况";

                    string hujidibiandongqingkuang = "       年  月  日 因      迁出（入）至";
                    sheet.Range[string.Format("C{0}:H{1}", row, row)].Merge();
                    sheet.Range[string.Format("C{0}:H{1}", row, row)].Text = hujidibiandongqingkuang;

                    row += 1;
                    sheet.Range[string.Format("A{0}:B{1}", row, row)].Merge();
                    sheet.Range[string.Format("A{0}:B{1}", row, row)].Text = "成员身份认定情况";

                    string chengyuanshenfenrendingqingkuang = "  2019年  1月 30日 被认定为    国庆村  集体经济组织成员";
                    sheet.Range[string.Format("C{0}:H{1}", row, row)].Merge();
                    sheet.Range[string.Format("C{0}:H{1}", row, row)].Text = chengyuanshenfenrendingqingkuang;

                    row += 1;
                    sheet.Range[string.Format("A{0}", row)].Text = "其他情况";                    
                    sheet.Range[string.Format("B{0}:H{1}", row, row)].Merge();

                    row += 1;
                    sheet.Range[string.Format("A{0}:C{1}", row, row)].Merge();
                    sheet.Range[string.Format("A{0}:C{1}", row, row)].Text = "目前在本集体经济组织状态";

                    string jitijingjizuzhizhuangtai = (string.IsNullOrEmpty(p.huY_renY) ? this.symbolSquareNull : this.symbolSquareRight) + "户在人在；" +
                                                      (string.IsNullOrEmpty(p.huY_renN) ? this.symbolSquareNull : this.symbolSquareRight) + "户在人不在；" +
                                                      (string.IsNullOrEmpty(p.huN_renY) ? this.symbolSquareNull : this.symbolSquareRight) + "人在户不在；" +
                                                      (string.IsNullOrEmpty(p.huN_renN) ? this.symbolSquareNull : this.symbolSquareRight) + "人户都不在";
                    sheet.Range[string.Format("D{0}:H{1}", row, row)].Merge();
                    sheet.Range[string.Format("D{0}:H{1}", row, row)].Text = jitijingjizuzhizhuangtai;
                    
                    itemIndex++;
                }

                CellStyle cellStyle = sheet.GetDefaultRowStyle(1);
                cellStyle.Font.Size = 12;
                cellStyle.Font.FontName = "仿宋";
                cellStyle.Font.IsBold = false;
                cellStyle.VerticalAlignment = VerticalAlignType.Center;
                cellStyle.HorizontalAlignment = HorizontalAlignType.Center;
                for (int i = sheet.FirstRow; i <= sheet.LastRow; i++)
                {
                    //if ((i - 1) % 13 != 0)
                    //{
                        sheet.SetDefaultRowStyle(i, cellStyle);
                        sheet.SetRowHeight(i, 15);
                    //}
                    //else
                    //{
                    //    sheet.SetRowHeight(i, 69);
                    //}
                }

                for (int i = 0; i < family.people.Count; i++)
                {
                    int start = (i * itemRowCount) + 3;
                    int end = start + 9;
                    sheet.Range[string.Format("A{0}:H{1}", start, end)].BorderInside(LineStyleType.Thin, ExcelColors.Black);
                    sheet.Range[string.Format("A{0}:H{1}", start, end)].BorderAround(LineStyleType.Medium, ExcelColors.Black);
                }

                sheet.SetColumnWidth(1, 10);
                sheet.SetColumnWidth(2, 10);
                sheet.SetColumnWidth(3, 6);
                sheet.SetColumnWidth(4, 6);
                sheet.SetColumnWidth(5, 13);
                sheet.SetColumnWidth(6, 7);
                sheet.SetColumnWidth(7, 13);
                sheet.SetColumnWidth(8, 14.86);

                sheet.Range["A2:A" + sheet.LastRow].Style.WrapText = true;

                book.SaveToFile("人口摸底调查表\\" + family.hostHostName + "_" + family.hostNo + ".xlsx", ExcelVersion.Version2010);

            }

            MessageBox.Show("Save OK");
        }



        private void btnLoadSymbols_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "xlsx文件|*.xlsx";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(this.openFileDialog1.FileName);

                Worksheet sheet = workbook.Worksheets[0];
                this.inputSymRight = sheet.Range["A1"].Text;
                this.symbolSquareRight = sheet.Range["A2"].Text;
                this.symbolSquareError = sheet.Range["B2"].Text;
                this.symbolSquareNull = sheet.Range["C2"].Text;
                this.symbolRight = sheet.Range["D2"].Text;

                workbook.Dispose();

                MessageBox.Show("符号导入成功");
            }
        }

        private void LoadTablePeopleInfo(string filePath)
        {
            this.lastHostName = "";
            this.dicFamilies.Clear();

            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filePath);

            Worksheet sheet = workbook.Worksheets[0];

            Family family = null;

            string hostName = "";
            for (int r = sheet.FirstRow + 2; r <= sheet.LastRow; r++)
            {
                if (string.IsNullOrEmpty(sheet[r, 3].Value.Trim()))
                {
                    break;
                }
                string hostNo = sheet[r, 1].Value.Trim();
                hostName = sheet[r, 2].Value.Trim();
                string liveState = sheet[r, 14].Value.Trim();
                if (liveState == "死亡" || liveState == "")
                {
                    continue;
                }
                if (sheet[r, 5].Value.Trim() == "")
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
                    if (this.dicFamilies.ContainsKey(hostName))
                    {
                        family = this.dicFamilies[hostName];
                    }
                    else
                    {
                        family = new Family()
                        {
                            hostNo = hostNo,
                            hostHostName = hostName
                        };

                        this.dicFamilies.Add(hostName, family);
                    }

                    this.lastHostName = hostName;
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
                    job = "",
                    phone = sheet[r, 17].Value.Trim()
                };
                p.isHost = p.hostName == p.name;
                Person pfrom5 = this.listperson.Find(item => item.name == p.name);
                if (pfrom5 != null)
                {
                    p.nation = pfrom5.nation;
                    p.location = pfrom5.location;
                    p.education = pfrom5.education;
                    p.job = pfrom5.job;
                }

                family.people.Add(p);

            }

            MessageBox.Show("导入成功！");
        }

        private void Load户籍信息表(string filePath)
        {
            this.listperson.Clear();
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filePath);

            Worksheet sheet = workbook.Worksheets[0];
            
            for (int r = sheet.FirstRow + 1; r <= sheet.LastRow; r++)
            {
                string id = sheet[r, 6].Value.Trim();
                Person p = new Person()
                {
                    hostName = sheet[r, 18].Value.Trim(),
                    name = sheet[r, 3].Value.Trim(),
                    relation = sheet[r, 2].Value.Trim(),
                    sex = sheet[r, 4].Value.Trim(),
                    nation = sheet[r, 5].Value.Trim(),
                    idNo = id.Length > 18 ? id.Substring(0, 18) : id,
                    group = sheet[r, 12].Value.Trim(),
                    location = sheet[r, 15].Value.Trim(),
                    education = sheet[r, 16].Value.Trim(),
                    job = sheet[r, 17].Value.Trim(),
                };
                p.isHost = p.hostName == p.name;

                this.listperson.Add(p);

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

            if (!string.IsNullOrEmpty(hostNo) && hostNo.Length <= 5)
            {
                if (hostNo.Length < 5)
                {
                    string zeroes = "";
                    for (int i = 0; i < 5 - hostNo.Length; i++)
                    {
                        zeroes += "0";
                    }
                    hostNo = zeroes + hostNo;
                }
                hostNo = this.hostNoPrefix + hostNo;
            }
            family.hostNo = hostNo;
            return hostNo;
        }

        private void btnFormatDate_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "xlsx文件|*.xlsx";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = this.openFileDialog1.FileName;

                Workbook workbook = new Workbook();
                workbook.LoadFromFile(filePath);

                Worksheet sheet = workbook.Worksheets[0];
                
                for (int r = sheet.FirstRow + 6; r <= sheet.LastRow - 6; r++)
                {
                    string date = sheet[r, 1].Value.Trim();
                    if (date.Length == 4)
                    {
                        sheet[r, 1].Value = date + "/12/31";
                    }
                }

                workbook.SaveToFile("FormatDate2.xlsx");
                MessageBox.Show("完成");
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

            public string phone;

        }

        private void btnLoad人口登记表_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "xlsx文件|*.xlsx";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.txt人口登记表路径.Text = this.openFileDialog1.FileName;
                this.Load人口登记表(this.openFileDialog1.FileName);
            }
        }

        private void Load人口登记表(string filePath)
        {
            this.lastHostName = "";
            this.dicFamilies.Clear();

            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filePath);

            Worksheet sheet = workbook.Worksheets[0];

            Family family = null;

            string hostName = "";
            for (int r = sheet.FirstRow + 2; r <= sheet.LastRow; r++)
            {
                if (string.IsNullOrEmpty(sheet[r, 3].Value.Trim()))
                {
                    break;
                }
                string hostNo = sheet[r, 1].Value.Trim().Trim('\'');
                hostName = sheet[r, 2].Value.Trim();
                string liveState = sheet[r, 14].Value.Trim();
                if (liveState == "死亡" || liveState == "")
                {
                    continue;
                }
                if (sheet[r, 5].Value.Trim() == "")
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
                    if (this.dicFamilies.ContainsKey(hostName))
                    {
                        family = this.dicFamilies[hostName];
                    }
                    else
                    {
                        family = new Family()
                        {
                            hostNo = hostNo,
                            hostHostName = hostName
                        };

                        this.dicFamilies.Add(hostName, family);
                    }

                    this.lastHostName = hostName;
                }

                string idNumber = sheet[r, 5].Value.Trim().Trim('\'').Substring(0, 18);
                Person p = new Person()
                {
                    hostNo = hostNo,
                    hostName = hostName,
                    name = sheet[r, 3].Value.Trim(),
                    relation = sheet[r, 4].Value.Trim(),
                    idNo = idNumber,
                    birthday = new DateTime(int.Parse(idNumber.Substring(6, 4)), int.Parse(idNumber.Substring(10, 2)), int.Parse(idNumber.Substring(12, 2))),
                    sex = sheet[r, 7].Value.Trim(),
                    huY_renY = sheet[r, 8].Value.Trim(),
                    huY_renN = sheet[r, 9].Value.Trim(),
                    huN_renY = sheet[r, 10].Value.Trim(),
                    huN_renN = sheet[r, 11].Value.Trim(),
                    group = sheet[r, 12].Value.Trim(),
                    isTuDiChengbao = sheet[r, 13].Value.Trim(),
                    lifeState = liveState,
                    marryState = sheet[r, 15].Value.Trim(),
                    phone = sheet[r, 16].Value.Trim(),
                    nation = sheet[r, 18].Value.Trim(),
                    location = "",
                    education = sheet[r, 19].Value.Trim(),
                    job = "粮农"
                };
                p.isHost = p.hostName == p.name;

                family.people.Add(p);

            }
            MessageBox.Show("导入成功！");
        }
    }
}