using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Drawing;
using System.Collections;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace 附和水准近似平差
{
    class Class1
    {
        public static void daochu1(StreamWriter sw, DataGridView d)
        {
            List<string> StrArray = new List<string>();
            string str = null;
            for (int i = 0; i < d.Columns.Count; i++)
            {
                StrArray.Add(string.Format("{0,-12}", d.Columns[i].HeaderText));
            }
            str = string.Join("  ", StrArray);
            sw.WriteLine(str);
            for (int i = 0; i < d.Rows.Count; i++)
            {
                str = null;
                StrArray.Clear();
                for (int i1 = 0; i1 < d.Columns.Count; i1++)
                {
                    StrArray.Add(string.Format("{0,-16}", d.Rows[i].Cells[i1].Value));
                }
                str = string.Join("  ", StrArray);
                sw.WriteLine(str);
            }
        }
        #region
        public static void daochu2(int h,Excel.Worksheet worksheet1, DataGridView d)
        {
            for (int i = 0; i < d.Columns.Count; i++)//表头
            {
                worksheet1.Cells[h, i + 1].Value = d.Columns[i].HeaderText;
            }
            for (int i = 0; i < d.Rows.Count; i++)
            {
                for (int j = 0; j < d.Columns.Count; j++)
                {
                    worksheet1.Cells[h + 1 + i, j + 1].Value = d.Rows[i].Cells[j].Value;
                }
            }
            worksheet1.Columns.AutoFit();//自动调整列宽
        }
        #endregion
        #region
        #endregion
        #region 绘制三角
        public static void sanjiao(Graphics g, PointF pf)
        {
            //绘制填充多边形的原理
            Bitmap bt1 = new Bitmap(20, 20);//画板
            PointF[] pfs2 = { new PointF(20, 10), new PointF(1, 0), new PointF(1, 20) };//三角的三个点
            Graphics g1 = Graphics.FromImage(bt1);
            g1.FillPolygon(Brushes.White, pfs2);//填充
            g1.DrawPolygon(new Pen(Color.Blue, 1.5f), pfs2);//绘制
            g.DrawImage((Image)bt1, pf.X - 10, pf.Y - 10);//图形绘制的位置
        }
        #endregion
        #region 绘制注记
        public static void ziti(Graphics g, PointF pf, string dianhao)
        {
            Bitmap bt2 = new Bitmap(50, 50);
            Graphics g2 = Graphics.FromImage(bt2);
            g2.RotateTransform(90);
            g2.TranslateTransform(0, -30);//划定原点位置
            g2.DrawString(dianhao, new Font("宋体", 20), Brushes.Green, new Point(5, 5));
            g.DrawImage((Image)bt2, pf.X - 25, pf.Y - 25);
        }
        #endregion
    }
}
