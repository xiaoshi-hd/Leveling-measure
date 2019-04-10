using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Drawing2D;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace 附和水准近似平差
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        #region 时间控件
        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView3.AllowUserToAddRows = false;
            dataGridView4.AllowUserToAddRows = false;

            toolStripStatusLabel3.Text = DateTime.Now.ToString();
            timer1.Enabled = true;
            timer1.Interval = 1000;
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            toolStripStatusLabel3.Text = DateTime.Now.ToString();
        }
        #endregion
        #region 变量声明
        Image image;
        List<string> dianhao;//点号
        List<string> cezhan;//测站
        List<double> dengji;//存储尺常数

        List<double> shangsihou;//后尺上丝
        List<double> xiasihou;//后尺下丝
        List<double> shangsiqian;//前尺上丝
        List<double> xiasiqian;//前尺下丝
        List<double> jibenhou;//后尺黑面（基础）分划
        List<double> jibenqian;//前尺黑面（基础）分划
        List<double> fuzhuqian;//前尺红面（辅助）分划
        List<double> fuzhuhou;//后尺红面（辅助）分划

        List<double> qianshiju;//前视距
        List<double> houshiju;//后视距
        List<double> shijucha;//视距差
        List<double> shijuchaleiji;//累计视距差
        List<double> shiju;//视距
        List<double> gaocha1;//黑面（基础）高差
        List<double> gaocha2;//红面（辅助）高差
        List<double> houcicha;//后尺差
        List<double> qiancicha;//前尺差
        List<double> houjianqian;//后减前差值
        List<double> gaochazhong;//高差中数

        List<double> gaizhengshu;//高差改正数
        List<double> gaizhenghougaocha;//改正后高差
        List<double> gaocheng;//高程
        List<double> yzuobiao;//绘图用，存储距离
        #endregion
        #region 变量初始化
        public void chushihua()
        {
            dianhao = new List<string>();
            cezhan = new List<string>();
            dengji = new List<double>();

            shangsihou = new List<double>();
            xiasihou = new List<double>();
            shangsiqian = new List<double>();
            xiasiqian = new List<double>();
            jibenhou = new List<double>();
            jibenqian = new List<double>();
            fuzhuqian = new List<double>();
            fuzhuhou = new List<double>();

            qianshiju = new List<double>();
            houshiju = new List<double>();
            shijucha = new List<double>();
            shijuchaleiji = new List<double>();
            shiju = new List<double>();
            gaocha1 = new List<double>();
            gaocha2 = new List<double>();
            gaochazhong = new List<double>();
            houcicha = new List<double>();
            qiancicha = new List<double>();
            houjianqian = new List<double>();

            gaizhengshu = new List<double>();
            gaizhenghougaocha = new List<double>();
            gaocheng = new List<double>();
            yzuobiao = new List<double>();
        }
        #endregion
        #region 文件打开
        private void 打开ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            rdb_4687.Checked = false;
            rdb_4787.Checked = false;

            openFileDialog1.Title = "附和水准数据打开";
            openFileDialog1.Filter = "文本文件(*.txt)|*.txt|Excel旧版本文件(*.xls)|*.xls|Excel新版本文件(*.xlsx)|*.xlsx";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                #region txt文档
                if (openFileDialog1.FilterIndex == 1)
                {
                    StreamReader sr = new StreamReader(openFileDialog1.FileName,Encoding.Default);
                    sr.ReadLine();
                    string[] str = sr.ReadLine().Split(',');
                    int i = 0, j = 0;
                    while (!sr.EndOfStream)
                    {
                        if (str[0] != "观测数据：")
                        {
                            dataGridView1.Rows.Add();
                            dataGridView1.Rows[i].Cells[0].Value = str[0];
                            dataGridView1.Rows[i].Cells[1].Value = str[1];
                            str = sr.ReadLine().Split(',');
                            i++;
                        }
                        else
                        { 
                            while (!sr.EndOfStream)
                            {
                                dataGridView2.Rows.Add();
                                str = sr.ReadLine().Split(',');
                                for (int a = 0; a < str.Length; a++)
                                {
                                    dataGridView2.Rows[j].Cells[a].Value = str[a];
                                }
                                j++;
                            }
                        }
                    }
                    sr.Close();
                }
                #endregion
                #region Excel文件
                else
                {
                    Excel.Application excel = new Excel.Application();
                    excel.Visible = false;
                    Excel.Workbook wb = excel.Application.Workbooks.Open(openFileDialog1.FileName);
                    Excel.Worksheet ws = excel.Workbooks[1].Worksheets[1];
                    int rows = ws.UsedRange.Rows.Count;
                    int columns = ws.UsedRange.Columns.Count;
                    for (int i = 0; i < 2; i++)
                    {
                        dataGridView1.Rows.Add();
                        dataGridView1.Rows[i].Cells[0].Value = ws.Cells[i + 2, 1].Value;
                        dataGridView1.Rows[i].Cells[1].Value = ws.Cells[i + 2, 2].Value;
                    }
                    for (int i = 0; i < rows - 4; i++)
                    {
                        dataGridView2.Rows.Add();
                        for (int j = 0; j < columns; j++)
                        {
                            dataGridView2.Rows[i].Cells[j].Value = ws.Cells[i + 5, j + 1].Value;
                        }
                    }
                    wb.Close();
                }
                #endregion
            }
        }
        #endregion
        #region 计算
        private void 计算ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            chushihua();
            dataGridView3.Rows.Clear();
            dataGridView4.Rows.Clear();
            dataGridView1.AllowUserToAddRows = false;
            dataGridView2.AllowUserToAddRows = false;
            #region 数据检核
            #region 等级判断
            if (rdb_2.Checked)
            {
                dengji.Add(301.550);
                dengji.Add(301.550);
            }
            else
            {
                if (rdb_4787.Checked)
                {
                    dengji.Add(4787);
                    dengji.Add(4687);
                }
                else if (rdb_4687.Checked)
                {
                    dengji.Add(4687);
                    dengji.Add(4787);
                }
                else
                {
                    MessageBox.Show("请选择四等起始尺常数！");
                    return;
                }
            }
            #endregion
            #region 数据导入
            try
            {
                dianhao.Add(dataGridView1.Rows[0].Cells[0].Value.ToString());
                gaocheng.Add(Convert.ToDouble(dataGridView1.Rows[0].Cells[1].Value));
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    dianhao.Add(dataGridView2.Rows[i].Cells[0].Value.ToString());
                    cezhan.Add(dataGridView2.Rows[i].Cells[0].Value.ToString());//测站
                    shangsihou.Add(Convert.ToDouble(dataGridView2.Rows[i].Cells[1].Value.ToString().Replace(" ", "")));
                    xiasihou.Add(Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value.ToString().Replace(" ", "")));
                    shangsiqian.Add(Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value.ToString().Replace(" ", "")));
                    xiasiqian.Add(Convert.ToDouble(dataGridView2.Rows[i].Cells[4].Value.ToString().Replace(" ", "")));
                    jibenhou.Add(Convert.ToDouble(dataGridView2.Rows[i].Cells[5].Value.ToString().Replace(" ", "")));
                    jibenqian.Add(Convert.ToDouble(dataGridView2.Rows[i].Cells[6].Value.ToString().Replace(" ", "")));
                    fuzhuqian.Add(Convert.ToDouble(dataGridView2.Rows[i].Cells[7].Value.ToString().Replace(" ", "")));
                    fuzhuhou.Add(Convert.ToDouble(dataGridView2.Rows[i].Cells[8].Value.ToString().Replace(" ", "")));
                }
                dianhao[dataGridView2.Rows.Count] = dataGridView1.Rows[1].Cells[0].Value.ToString();
                gaocheng.Add(Convert.ToDouble(dataGridView1.Rows[1].Cells[1].Value));
            }
            catch
            {
                MessageBox.Show("请输入正确的数据！");
            }
            #endregion
            #region 计算
            double a = 0, b = 0;
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                if (rdb_2.Checked)//二等精确到0.01mm
                {
                    houshiju.Add(Math.Round(shangsihou[i] - xiasihou[i], 2));//计算出来的视距需要乘以100再化成米，二等以cm为单位处以100
                    qianshiju.Add(Math.Round(shangsiqian[i] - xiasiqian[i],2));
                    shijucha.Add(Math.Round(houshiju[i] - qianshiju[i],2));
                    a = a + shijucha[i];
                    shijuchaleiji.Add(Math.Round(a,2));
                    shiju.Add(Math.Round(houshiju[i] + qianshiju[i],2));
                    b = b + shiju[i];
                    yzuobiao.Add(Math.Round(b,2));
                    houcicha.Add(Math.Round((jibenhou[i] + dengji[i % 2] - fuzhuhou[i]) * 10, 2));//后尺差以mm为单位
                    qiancicha.Add(Math.Round((jibenqian[i] + dengji[-(i % 2) + 1] - fuzhuqian[i]) * 10,2));
                    houjianqian.Add(Math.Round((houcicha[i] - qiancicha[i]),2));
                    gaocha1.Add(Math.Round(jibenhou[i] - jibenqian[i], 3));//高程以cm为单位
                    gaocha2.Add(Math.Round(fuzhuhou[i] - fuzhuqian[i],3));
                    gaochazhong.Add(Math.Round(gaocha1[i] - houjianqian[i] /10 / 2,3));
                }
                else//四等精确到1mm
                {
                    houshiju.Add(Math.Round((shangsihou[i] - xiasihou[i]) / 10, 1));//计算出来的视距需要乘以100再化成米，二等以mm为单位处以1000
                    qianshiju.Add(Math.Round((shangsiqian[i] - xiasiqian[i]) / 10, 1));
                    shijucha.Add(Math.Round((houshiju[i] - qianshiju[i]), 1));
                    a = a + shijucha[i];
                    shijuchaleiji.Add(Math.Round(a, 1));
                    shiju.Add(Math.Round(houshiju[i] + qianshiju[i], 1));
                    b = b + shiju[i];
                    yzuobiao.Add(Math.Round(b, 1));
                    houcicha.Add(Math.Round((jibenhou[i] + dengji[i % 2] - fuzhuhou[i])));//后尺差以mm为单位
                    qiancicha.Add(Math.Round((jibenqian[i] + dengji[-(i % 2) + 1] - fuzhuqian[i])));
                    houjianqian.Add(Math.Round((houcicha[i] - qiancicha[i])));
                    gaocha1.Add(Math.Round((jibenhou[i] - jibenqian[i]) / 10, 1));//高程以cm为单位
                    gaocha2.Add(Math.Round((fuzhuhou[i] - fuzhuqian[i]) / 10, 1));
                    gaochazhong.Add(Math.Round(gaocha1[i] - houjianqian[i] / 10 / 2, 2));
                }
            }
            #endregion
            #region 数据导出
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                dataGridView4.Rows.Add();
                dataGridView4.Rows[i].Cells[0].Value = cezhan[i];
                dataGridView4.Rows[i].Cells[1].Value = houshiju[i];
                dataGridView4.Rows[i].Cells[2].Value = qianshiju[i];
                dataGridView4.Rows[i].Cells[3].Value = shijucha[i];
                dataGridView4.Rows[i].Cells[4].Value = shijuchaleiji[i];
                dataGridView4.Rows[i].Cells[5].Value = shiju[i];
                dataGridView4.Rows[i].Cells[6].Value = houcicha[i];
                dataGridView4.Rows[i].Cells[7].Value = qiancicha[i];
                dataGridView4.Rows[i].Cells[8].Value = houjianqian[i];
                dataGridView4.Rows[i].Cells[9].Value = gaocha1[i];
                dataGridView4.Rows[i].Cells[10].Value = gaocha2[i];
                dataGridView4.Rows[i].Cells[11].Value = gaochazhong[i];
                #region 超限判断
                if (rdb_2.Checked)
                {
                    if (Math.Abs(houshiju[i]) >= 50)
                    { dataGridView4.Rows[i].Cells[1].Style.ForeColor = Color.Red; }
                    if (Math.Abs(qianshiju[i]) >= 50)
                    { dataGridView4.Rows[i].Cells[2].Style.ForeColor = Color.Red; }
                    if (Math.Abs(shijucha[i]) >= 1)
                    { dataGridView4.Rows[i].Cells[3].Style.ForeColor = Color.Red; }
                    if (Math.Abs(shijuchaleiji[i]) >= 3)
                    { dataGridView4.Rows[i].Cells[4].Style.ForeColor = Color.Red; }
                    if (Math.Abs(houcicha[i]) >= 0.4)
                    { dataGridView4.Rows[i].Cells[6].Style.ForeColor = Color.Red; }
                    if (Math.Abs(qiancicha[i]) >= 0.4)
                    { dataGridView4.Rows[i].Cells[7].Style.ForeColor = Color.Red; }
                    if (Math.Abs(houjianqian[i]) >= 0.6)
                    { dataGridView4.Rows[i].Cells[8].Style.ForeColor = Color.Red; }
                }
                else
                {
                    if (Math.Abs(houshiju[i]) >= 80)
                    { dataGridView4.Rows[i].Cells[1].Style.ForeColor = Color.Red; }
                    if (Math.Abs(qianshiju[i]) >= 80)
                    { dataGridView4.Rows[i].Cells[2].Style.ForeColor = Color.Red; }
                    if (Math.Abs(shijucha[i]) >= 5)
                    { dataGridView4.Rows[i].Cells[3].Style.ForeColor = Color.Red; }
                    if (Math.Abs(shijuchaleiji[i]) >= 10)
                    { dataGridView4.Rows[i].Cells[4].Style.ForeColor = Color.Red; }
                    if (Math.Abs(houcicha[i]) >= 3)
                    { dataGridView4.Rows[i].Cells[6].Style.ForeColor = Color.Red; }
                    if (Math.Abs(qiancicha[i]) >= 3)
                    { dataGridView4.Rows[i].Cells[7].Style.ForeColor = Color.Red; }
                    if (Math.Abs(houjianqian[i]) >= 5)
                    { dataGridView4.Rows[i].Cells[8].Style.ForeColor = Color.Red; }
                }
                #endregion
            }
            #endregion
            #endregion
            #region 水准平差
            for (int i = 0; i < dianhao.Count; i++)
            {
                dataGridView3.Rows.Add();
                dataGridView3.Rows[i].Cells[0].Value = dianhao[i];
            }
            double bhc = gaochazhong.Sum() - (gaocheng[1] * 100 - gaocheng[0] * 100);//观测值减去真实值
            if (rdb_2.Checked)
            {
                for (int i = 0; i < gaochazhong.Count; i++)
                {
                    gaizhengshu.Add(bhc * shiju[i] / shiju.Sum());
                    gaizhenghougaocha.Add(gaochazhong[i] - gaizhengshu[i]);
                    gaocheng.Insert(i + 1, gaocheng[i] + gaizhenghougaocha[i] / 100);
                    dataGridView3.Rows[i + 1].Cells[1].Value = Math.Round(shiju[i], 2);
                    dataGridView3.Rows[i + 1].Cells[2].Value = Math.Round(gaochazhong[i], 3);
                    dataGridView3.Rows[i + 1].Cells[3].Value = Math.Round(gaizhengshu[i], 3);
                    dataGridView3.Rows[i + 1].Cells[4].Value = Math.Round(gaizhenghougaocha[i], 3);
                }
                for (int i = 0; i < gaocheng.Count - 1; i++)
                {
                    dataGridView3.Rows[i].Cells[5].Value = Math.Round(gaocheng[i], 5);
                }
                dataGridView3.Rows.Add();
                dataGridView3.Rows[dianhao.Count].Cells[3].Value = Math.Round(bhc,3);
            }
            else
            {
                for (int i = 0; i < gaochazhong.Count; i++)
                {
                    gaizhengshu.Add(bhc * shiju[i] / shiju.Sum());
                    gaizhenghougaocha.Add(gaochazhong[i] - gaizhengshu[i]);
                    gaocheng.Insert(i + 1, gaocheng[i] + gaizhenghougaocha[i] / 100);
                    dataGridView3.Rows[i + 1].Cells[1].Value = Math.Round(shiju[i], 1);
                    dataGridView3.Rows[i + 1].Cells[2].Value = Math.Round(gaochazhong[i], 2);
                    dataGridView3.Rows[i + 1].Cells[3].Value = Math.Round(gaizhengshu[i], 2);
                    dataGridView3.Rows[i + 1].Cells[4].Value = Math.Round(gaizhenghougaocha[i], 2);
                }
                for (int i = 0; i < gaocheng.Count - 1; i++)//高程多一个，所以减1不把它输出
                {
                    dataGridView3.Rows[i].Cells[5].Value = Math.Round(gaocheng[i], 4);
                }
                dataGridView3.Rows.Add();
                dataGridView3.Rows[dianhao.Count].Cells[3].Value = Math.Round(bhc, 4);
            }
            #endregion
            yzuobiao.Insert(0, 0);
        }
        #endregion
        #region 文件保存
        private void 保存ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Title = "附和水准计算保存";
            saveFileDialog1.Filter = "文本文件(*.txt)|*.txt|Excel数据文件(*.xls)|*.xls|Excel表格(*.xlsx)|*.xlsx";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                #region TXT文档保存
                if (saveFileDialog1.FilterIndex == 1)
                {
                    StreamWriter sw = new StreamWriter(saveFileDialog1.FileName);
                    sw.WriteLine("附合水准平差计算\n");
                    Class1.daochu1(sw, dataGridView4);
                    sw.WriteLine("高程配赋平差计算\n");
                    Class1.daochu1(sw, dataGridView3);
                    sw.Flush();
                    MessageBox.Show("保存成功！");
                }
                #endregion
                #region Excel文档保存
                else
                {
                    Excel.Application excel1 = new Excel.Application();//创建一个excel对象
                    Excel.Workbook workbook1 = excel1.Workbooks.Add(true);//为该excel对象添加一个工作簿
                    Excel.Worksheet worksheet1 = excel1.Workbooks[1].Worksheets[1];//获取工作簿中的第一个工作表
                    Class1.daochu2(1, worksheet1, dataGridView4);
                    Class1.daochu2(dataGridView2.Rows.Count + 2, worksheet1, dataGridView3);
                    worksheet1.SaveAs(saveFileDialog1.FileName);//保存工作表
                    workbook1.Close();
                    MessageBox.Show("保存成功！");
                }
                #endregion
            }
        }
        #endregion
        #region 绘图
        private void 绘图ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Pen p = new Pen(Color.Black, 1);//高差乘以10
            image = new Bitmap((int)(yzuobiao.Max() - yzuobiao.Min()) + 200, (int)(gaocheng.Max() * 10 - gaocheng.Min() * 10) + 200);//显示图形范围
            Graphics g = Graphics.FromImage(image);
            g.RotateTransform(-90);//旋转为测量坐标系
            g.TranslateTransform(-(int)(gaocheng.Max() * 10 + 100), -(int)(yzuobiao.Min() - 100));//划定原点位置
            PointF[] pf = new PointF[yzuobiao.Count];//yzuobiao个数为15个
            for (int i = 0; i < yzuobiao.Count; i++)
            {
                pf[i].X = (float)gaocheng[i] * 10;
                pf[i].Y = (float)yzuobiao[i];
            }
            g.DrawLines(p, pf);
            //绘制三角
            Class1.sanjiao(g, pf[0]);
            Class1.sanjiao(g, pf[pf.Length - 1]);
            //绘制字体
            for (int i = 0; i < dianhao.Count; i++)
            {
               Class1.ziti(g, pf[i], dianhao[i].ToString());
            }
            pictureBox1.Image = (Image)image;
        }
        #endregion
        #region bmp图形保存
        private void bmp图形保存ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog save1 = new SaveFileDialog();
            save1.Filter = "图像文件(*.bmp)|*.bmp";
            if (save1.ShowDialog() == DialogResult.OK)
            {
                image.Save(save1.FileName);
            }
            MessageBox.Show("保存成功！");
        }
        #endregion
        #region dxf图形保存
        private void dxf图形保存ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Title = "保存dxf文件";
            saveFileDialog1.Filter = "AutoCAD dxf文件(*.dxf)|*.dxf";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                using (StreamWriter sw = new StreamWriter(saveFileDialog1.FileName))
                {
                    sw.Write("0\nSECTION\n");//第一个段的开始
                    #region 表段
                    sw.Write("2\nTABLES\n");//表段的开始 不需要END结束
                    sw.Write("0\nTABLE\n");//每个表都由带有标签TABLE的组码0引入
                    #region 图层
                    sw.Write("2\nLAYER\n");//由组码2标识具体表，此处为图层

                    sw.Write("0\nLAYER\n");//每个表条目包括指定条目类型的组码0引入,实体图层
                    sw.Write("70\n");//表中最大条目数
                    sw.Write("0\n");
                    sw.Write("2\nshiti\n");//设置图层名称
                    sw.Write("62\n");//颜色代码
                    sw.Write("10\n");//10表示红色，50表示黄色，170表示蓝色，90表示绿色，130表示青色
                    sw.Write("6\n");//线型名称
                    sw.Write("CONTINUOUS\n");//表示直线

                    sw.Write("0\nLAYER\n");//注记图层
                    sw.Write("70\n");
                    sw.Write("0\n");
                    sw.Write("2\nzhuji\n");
                    sw.Write("62\n");
                    sw.Write("50\n");
                    sw.Write("6\n");
                    sw.Write("CONTINUOUS\n");
                    #endregion
                    sw.Write("0\nENDTAB\n");//TABLE段结束
                    #endregion
                    sw.Write("0\nENDSEC\n");//第一段结束

                    sw.Write("0\nSECTION\n");//第二个段的开始
                    #region 实体段
                    sw.Write("2\nENTITIES\n");//实体段开始
                    #region 绘制三角形
                    sw.Write(san(yzuobiao[0], gaocheng[0] / 10));
                    sw.Write(san(yzuobiao[yzuobiao.Count - 1], gaocheng[gaocheng.Count - 1]));
                    #endregion
                    #region 绘制线路
                    sw.Write("0\nPOLYLINE\n");//多线段绘制，为一个整体线段
                    sw.Write("8\n");//图层
                    sw.Write("shiti\n");//没有图层的话会创建一个，有图层可以调用创建的图层，以默认设置
                    sw.Write("66\n");//不太懂，应该是多线个数
                    sw.Write("1\n");
                    for (int i = 0; i < yzuobiao.Count; i++)
                    {
                        sw.Write("0\nVERTEX\n");//多线段标识
                        sw.Write("8\n");//图层
                        sw.Write("shiti\n");
                        sw.Write("10\n");//X坐标
                        sw.Write(yzuobiao[i] + "\n");
                        sw.Write("20\n");//Y坐标
                        sw.Write(gaocheng[i] + "\n");
                    }
                    sw.Write("0\nSEQEND\n");//多线段结束
                    #endregion
                    #region 文字注记
                    for (int i = 0; i < yzuobiao.Count; i++)
                    {
                        sw.Write("0\nTEXT\n");//单行文字
                        sw.Write("8\n");
                        sw.Write("zhuji\n");
                        sw.Write("10\n");//字体起点X
                        sw.Write(yzuobiao[i] - 5 + "\n");
                        sw.Write("20\n");//字体起点Y
                        sw.Write(gaocheng[i] - 5 + "\n");
                        sw.Write("40\n15\n");//字体高度
                        sw.Write("1\n" + dianhao[i] + "\n");//文字内容
                    }
                    #endregion
                    #endregion
                    sw.Write("0\nENDSEC\n");//第二段结束
                    sw.Write("0\nEOF\n");//文件结束
                    MessageBox.Show("保存成功");
                }
            }
        }
        #region 三角存储
        public static string san(double x, double y)
        {
            string m;
            m = "0\nPOLYLINE\n8\nshiti\n66\n1\n";//用多线画的三角形
            m = m + "0\nVERTEX\n8\nshiti\n10\n";
            m = m + Convert.ToString(x - 5) + "\n20\n";
            m = m + Convert.ToString(y - 5) + "\n";
            m = m + "0\nVERTEX\n8\nshiti\n10\n";
            m = m + Convert.ToString(x + 5) + "\n20\n";
            m = m + Convert.ToString(y - 5) + "\n";
            m = m + "0\nVERTEX\n8\nshiti\n10\n";
            m = m + Convert.ToString(x) + "\n20\n";
            m = m + Convert.ToString(y + 5) + "\n";
            m = m + "0\nVERTEX\n8\nshiti\n10\n";
            m = m + Convert.ToString(x - 5) + "\n20\n";
            m = m + Convert.ToString(y - 5) + "\n" + "0\nSEQEND\n";
            return m;
        }
        #endregion
        #endregion
        #region 刷新
        private void 刷新ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();
            dataGridView4.Rows.Clear();
            pictureBox1.Image = null;
            dataGridView1.AllowUserToAddRows = true;
            dataGridView2.AllowUserToAddRows = true;
            rdb_4687.Checked = false;
            rdb_4787.Checked = false;
        }
        #endregion

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }
    }
}
