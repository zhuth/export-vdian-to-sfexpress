using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Diagnostics;
using NPOI.XSSF.UserModel;

namespace WeidianExportToSF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = ((string[])e.Data.GetData(DataFormats.FileDrop));
            var file = files[0];
            FileStream fs = new FileStream(@"sfexpress.xlsx", FileMode.Open, FileAccess.Read);
            var templateWorkbook = new XSSFWorkbook(fs);
            var sheet = templateWorkbook.GetSheetAt(0);

            FileStream inputfs = new FileStream(file, FileMode.Open, FileAccess.Read);
            var inputWorkbook = new XSSFWorkbook(inputfs);
            var input = inputWorkbook.GetSheetAt(0);
            
            for (int i = 1; i < 1000; ++i)
            {
                var dataRow = input.GetRow(i);
                var tmplRow = sheet.CreateRow(i);

                if (dataRow == null) break;

                var s_cols = "订单编号	订单金额	订单状态	订单类型	下单时间	付款时间	发货时间	买家确认收货时间	收件人姓名	收件人手机	商品编码	购买数量	商品价格	省	市	区	收货详细地址	物流公司	物流单号	商品总件数	订单描述	优惠金额	运费	推广费	订单优惠	订单备注	微信	备注	分销商店铺ID	分销商注册姓名	分销商手机号	分成金额	下单账号	是否已成团	身份证号".Split('\t');
                var t_cols = "收件人姓名	收件人手机		省	市	区	收货详细地址	书	1	商品总件数	商品总件数	顺丰次日	寄付现结".Split('\t');

                Dictionary<string, string> rec = new Dictionary<string, string>();
                for (int j = 0; j < s_cols.Length; ++j)
                {
                    var colname = s_cols[j];
                    var col = dataRow.GetCell(j);
                    var v = col.StringCellValue;
                    if (string.IsNullOrWhiteSpace(v)) continue;
                    rec.Add(colname, v.Trim());
                }

                if (rec["省"] == "上海" || rec["省"] == "北京" || rec["省"] == "天津" || rec["省"] == "重庆")
                {
                    rec["区"] = rec["市"];
                    rec["市"] = rec["省"] + "市";
                }
                else
                {
                    rec["省"] += "省";
                }

                for (int j = 0; j < t_cols.Length; ++j)
                {
                    var colname = t_cols[j];
                    tmplRow.CreateCell(j).SetCellValue(rec.ContainsKey(colname) ? rec[colname] : colname);
                }
            }

            using (var outputfs = new FileStream("output.xlsx", FileMode.Create))
                templateWorkbook.Write(outputfs);

            Process.Start("excel.exe", "output.xlsx");
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {

            string file = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false))
            {
                e.Effect = DragDropEffects.All;
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
