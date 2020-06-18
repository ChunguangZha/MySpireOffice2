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
    public partial class GuQuanZhengGongZuoBiao : Form
    {
        public List<F> fs = new List<F>();

        public GuQuanZhengGongZuoBiao()
        {
            InitializeComponent();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "xlsx文件|*.xlsx";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.LoadData(this.openFileDialog1.FileName);
            }
        }

        private void LoadData(string filePath)
        {
            this.fs.Clear();
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filePath);

            Worksheet sheet = workbook.Worksheets[5];

            for (int r = sheet.FirstRow + 1; r <= sheet.LastRow; r++)
            {
                string hostName = sheet[r, 5].Value.Trim();
                string stockNo = sheet[r, 7].Value.Trim();
                string pName = sheet[r, 10].Value.Trim();
                if (string.IsNullOrEmpty(hostName) && string.IsNullOrEmpty(pName))
                {
                    break;
                }

                if (!string.IsNullOrEmpty(hostName))
                {
                    string pSex = sheet[r, 11].Value.Trim().Substring(0, 1);
                    string hostid = sheet[r, 14].Value.Trim();
                    if (hostid.Length > 18)
                    {
                        hostid = hostid.Substring(0, 18);
                    }
                    F f = new F()
                    {
                        HostName = hostName,
                        HostSex = pSex,
                        HostID = hostid,
                        StockNo = stockNo
                    };
                    this.fs.Add(f);


                    int offset = 0;
                    do
                    {
                        if (string.IsNullOrEmpty(sheet[r + offset, 10].Value.Trim()))
                        {
                            break;
                        }
                        string pRelation = sheet[r + offset, 13].Value.Trim();
                        if (pRelation == "妻" || pRelation == "夫" || pRelation == "妻子")
                        {
                            pRelation = "配偶";
                        }

                        int memberStock = Convert.ToInt32(sheet[r + offset, 15].Value.Trim());
                        float ageStock = Convert.ToSingle(sheet[r + offset, 16].Value.Trim());
                        float countStock = memberStock + ageStock;
                        float pmoney = 0;
                        float.TryParse(sheet[r + offset, 20].Value.Trim(), out pmoney);
                        string id = sheet[r + offset, 14].Value.Trim();
                        if (id.Length > 18)
                        {
                            id = id.Substring(0, 18);
                        }

                        P p1 = new P()
                        {
                            Name = sheet[r + offset, 10].Value.Trim(),
                            Sex = sheet[r + offset, 11].Value.Trim().Substring(0, 1),
                            Age = Convert.ToInt32(sheet[r + offset, 12].Value.Trim()),
                            Relation = pRelation,
                            IDNo = id,
                            MemeberStock = memberStock,
                            AgeStock = ageStock,
                            CountStock = countStock,
                            PMoney = pmoney
                        };

                        if (p1.Relation == "配偶")
                        {
                            if (f.People.FirstOrDefault(item => item.Relation == "配偶") != null)
                            {
                                Console.WriteLine(p1.Name + "  两媳妇!");
                                MessageBox.Show(p1.Name + "  两媳妇!");
                            }
                        }
                        if (p1.Relation == "户主")
                        {
                            if (f.People.FirstOrDefault(item => item.Relation == "户主") != null)
                            {
                                Console.WriteLine(p1.Name + "  两户主!");
                            }
                        }

                        foreach (var item in this.fs)
                        {
                            P existP = item.People.Find(p => p.IDNo == p1.IDNo);
                            if (existP != null)
                            {
                                Console.WriteLine("重复身份证号：" + existP.Name + " -- " + existP.IDNo + "   " + p1.Name + " -- " + p1.IDNo);
                            }
                        }
                        
                        f.People.Add(p1);
                        f.PersonCount++;
                        f.FStock += p1.CountStock;
                        f.FMoney += p1.PMoney;
                        offset++;
                    } while (string.IsNullOrEmpty(sheet[r + offset, 5].Value.Trim()));

                    r = r + offset - 1;
                }

            }

            MessageBox.Show("导入成功！");
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {

            Workbook book = new Workbook();
            Worksheet sheet = book.Worksheets[0];
            sheet.Name = "股权证工作表";
            int row = 1;

            CellStyle cellStyle = sheet.GetDefaultRowStyle(1);
            cellStyle.Font.Size = 12;
            cellStyle.Font.FontName = "宋体";

            for (int i = 0; i < fs.Count; i++)
            {
                F f = fs[i];

                int endRow = row + f.PersonCount - 1;
                sheet.Range[string.Format("A{0}:A{1}", row, endRow)].Merge();
                sheet.Range[string.Format("A{0}:A{1}", row, endRow)].Text = "讷河市";
                sheet.Range[string.Format("D{0}", row)].Text = f.StockNo;
                if (f.HostName == f.People[0].Name)
                {
                    sheet.Range[string.Format("F{0}", row)].Text = f.HostName;
                }
                else
                {
                    sheet.Range[string.Format("F{0}", row)].Text = f.People[0].Name;
                    f.People[0].Relation = "户主";
                    sheet.Range[string.Format("F{0}", row)].Style.Color = Color.Red;
                }
                sheet.Range[string.Format("G{0}", row)].Text = f.HostSex;
                sheet.Range[string.Format("H{0}", row)].Text = f.Nation;
                sheet.Range[string.Format("I{0}", row)].Text = f.HostID;
                sheet.Range[string.Format("J{0}", row)].Value = f.PersonCount.ToString();
                sheet.Range[string.Format("V{0}:V{1}", row, endRow)].Merge();
                sheet.Range[string.Format("V{0}", row)].Value = f.FStock.ToString("0.00");
                sheet.Range[string.Format("W{0}:W{1}", row, endRow)].Merge();
                sheet.Range[string.Format("W{0}", row)].Value = f.FMoney.ToString("0.00");

                for (int j = 0; j < f.People.Count; j++)
                {
                    P p = f.People[j];

                    sheet.Range[string.Format("B{0}", row)].Text = "龙河镇";
                    sheet.Range[string.Format("C{0}", row)].Text = "讷河市龙河镇国庆村股份经济合作社";
                    sheet.Range[string.Format("K{0}", row)].Text = p.Name;
                    sheet.Range[string.Format("L{0}", row)].Text = p.Sex;
                    sheet.Range[string.Format("M{0}", row)].Value = p.Age.ToString();
                    sheet.Range[string.Format("N{0}", row)].Text = p.Relation;
                    sheet.Range[string.Format("O{0}", row)].Text = p.IDNo;
                    sheet.Range[string.Format("P{0}", row)].Value = p.MemeberStock.ToString();
                    sheet.Range[string.Format("Q{0}", row)].Value = p.AgeStock.ToString();
                    sheet.Range[string.Format("T{0}", row)].Value = p.CountStock.ToString();
                    sheet.Range[string.Format("U{0}", row)].Value = p.PMoney.ToString();

                    row++;
                }
            }

            book.SaveToFile("国庆村股权证工作表.xlsx", ExcelVersion.Version2010);


            MessageBox.Show("Save OK");
        }
    }

    public class P
    {
        public string Name;
        public string Sex;
        public int Age;
        public string Relation;
        public string IDNo;
        public int MemeberStock;
        public float AgeStock;
        public float CountStock;
        public float PMoney;
    }

    public class F
    {
        public string StockNo;
        public string HostName;
        public string HostSex;
        public string Nation = "汉族";
        public string HostID;
        public int PersonCount;
        public float FStock;
        public float FMoney;

        public List<P> People = new List<P>();
    }
}
