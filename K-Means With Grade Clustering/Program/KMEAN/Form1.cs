using System;
using System.Text;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.Collections;

namespace KMEAN
{
    public partial class Form1 : Form
    {
        public static int ClassNum = 5;
        public static double EndFlag = 1;
        public static int Dimension = 5;
        public static int LOFk = 0;

        public int[] OrderClass;
        public int RowCount;
        public int[] CenterArray;
        public double[,] CenterArrayParams;
        public double[,] RenewCenterArrayParams;
        public double[] GradePercent;
        public double[] GradeFurMax;
        public double[] GradeHalfMax;
        public int[] BaseGrade; 
        public static double vurMax = 99999;
        public static double BalEndFlag;
        public int ClusterCount = 0;
        public ArrayList[] ClusterAssem;
        public int[,] NkLOF;
        public double[] lrdk;
        public double[] LOFK;
        public static double maxLof;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CenterArray = new int[ClassNum];
            OrderClass = new int[ClassNum];
            CenterArrayParams = new double[ClassNum, Dimension];
            RenewCenterArrayParams = new double[ClassNum, Dimension];
            ClusterAssem = new ArrayList[5];
            BaseGrade = new int[5];
            GradeFurMax = new double[ClassNum];
            GradeHalfMax = new double[ClassNum];
            BalEndFlag = 99999999;
            for (int i = 0; i < ClassNum; i++) {
                ClusterAssem[i] = new ArrayList();
            }
        }

        private void LoadDataXLs_Click(object sender, EventArgs e)
        {
            try
            {
                string tmppath = Application.StartupPath + xlsDo.m_xlsPath;
                StringBuilder sbr = new StringBuilder();
                FileStream fs = File.OpenRead(tmppath);
                HSSFWorkbook wk = new HSSFWorkbook(fs);
                ISheet sheet = wk.GetSheetAt(0);
                RowCount = sheet.LastRowNum;

                ListViewItem tmpLine = new ListViewItem();
                for (int j = 1; j <= sheet.LastRowNum; j++)
                {
                    IRow row = sheet.GetRow(j);
                    if (row != null)
                    {
                        tmpLine = new ListViewItem(row.GetCell(0).ToString());
                        for (int k = 1; k <= row.LastCellNum; k++)  
                        {
                            ICell cell = row.GetCell(k);  
                            if (cell != null)
                            {
                                tmpLine.SubItems.Add(cell.ToString());
                            }
                        }
                        tmpLine.SubItems.Add("");
                        tmpLine.SubItems.Add("");
                    }
                    XlsDataSh.Items.Add(tmpLine);
                }
                LoadDataXLs.Enabled = false;
                richTextBox1.Text += "读取Excel文件成功.";
                kMEANBtn.Enabled = true;
                GradePercent = new double[RowCount];
            }
            catch (Exception ex) {
                richTextBox1.Text += "\n读取Exce文件发生错误,原因是:" + ex.ToString();
            }
        }

        private void kMEANBtn_Click(object sender, EventArgs e)
        {
            kMEANBtn.Enabled = false;
            //richTextBox1.Text += "\nK - Means开始";
            //KMeanMain();
            LOFDist();
        }

        private void KMeanMain() {
            InitCenter();
            richTextBox1.Text += "\n聚类中心:Center1:" + CenterArray[0] + ";Center2:" + CenterArray[1] + ";Center3:" + CenterArray[2] + ";Center4:" + CenterArray[3] + ";Center5:" +CenterArray[4];
            do
            {
                Cluster();
                RenewCenter();
                ClusterCount++;
            } while (CalEndFlag());
            richTextBox1.Text += "\n已经收敛,共进行聚类" + System.Convert.ToString(ClusterCount) +"次" ;
            OutPExBtn.Enabled = true;
        }

        private Boolean CalEndFlag() {
            double tmpDifferDis = 0, tmpSameDis = 0;

            for (int i = 0; i < ClassNum; i++) {
                tmpSameDis = 0;
                for (int j = 0; j < Dimension; j++) {
                    tmpSameDis += Math.Pow((RenewCenterArrayParams[i, j] - CenterArrayParams[i, j]),2);
                }
                tmpDifferDis += Math.Pow(tmpSameDis,1.0/2);
            }
            if ((BalEndFlag - tmpDifferDis) <= EndFlag) return false;
            else {
                for (int i = 0; i < ClassNum; i++) {
                    for (int j = 0; j < Dimension; j++) {
                        CenterArrayParams[i, j] = RenewCenterArrayParams[i, j];
                    }
                }
                BalEndFlag = tmpDifferDis;
                tmpDifferDis = 0;
                return true;
            }
        }

        private void InitCenter()
        {
            Random tmpRandSeed = new Random();
            double tmpDistanceSame = 0,tmpDistanceDiffer = 0,tmpMaxDifferDis = 0;
            
            CenterArray[0] = tmpRandSeed.Next(0, RowCount - 1);
            CenterArray[1] = tmpRandSeed.Next(0, RowCount - 1);
            CenterArray[2] = tmpRandSeed.Next(0, RowCount - 1);
            CenterArray[3] = tmpRandSeed.Next(0, RowCount - 1);
            CenterArray[4] = tmpRandSeed.Next(0, RowCount - 1);
            for (int i = 0; i < ClassNum; i++) {
                for (int w = 3; w < 3 + Dimension; w++) {
                    CenterArrayParams[i, w - 3] = System.Convert.ToDouble(XlsDataSh.Items[CenterArray[i]].SubItems[w].Text);
                }
            }

        }

        private void Cluster()
        {
            int tmpclass = 0;
            double tmpClusDis = 0, tmpClusMinDis = 0;

            for (int i = 0; i < ClassNum; i++)
            {
                ClusterAssem[i].Clear();
            }

            for (int i = 0; i < RowCount - 1; i++)
            {
                tmpClusMinDis = vurMax;
                for (int j = 0; j < ClassNum; j++)
                {
                    tmpClusDis = 0;
                    for (int k = 3; k < 3 + Dimension; k++)
                    {
                        tmpClusDis += Math.Pow((System.Convert.ToDouble(XlsDataSh.Items[i].SubItems[k].Text) - CenterArrayParams[j, k - 3]), 2);
                    }
                    if (tmpClusDis < tmpClusMinDis)
                    {
                        tmpclass = j;
                        tmpClusMinDis = tmpClusDis;
                    }
                }
                ClusterAssem[tmpclass].Add(i);
            }
            richTextBox1.Text += ("\n第"+ (ClusterCount+1) +"次 S1:" + ClusterAssem[0].Count + ",S2:" + ClusterAssem[1].Count + ",S3:" + ClusterAssem[2].Count + ",S4:" + ClusterAssem[3].Count + ";S5:" + ClusterAssem[4].Count);
            if (ClusterAssem[0].Count == 0 || ClusterAssem[1].Count == 0 || ClusterAssem[2].Count == 0 || ClusterAssem[3].Count == 0 || ClusterAssem[4].Count == 0)
            {
                richTextBox1.Text += "\n重新初始化";
                InitCenter();
                Cluster();
            }
        }

        private void RenewCenter() {
            double tmpSameDis = 0;

            for (int i = 0; i < ClassNum; i++) {
                for (int k = 3; k < 3 + Dimension; k++)
                {
                    tmpSameDis = 0;
                    foreach (object n in ClusterAssem[i]) {
                        tmpSameDis += System.Convert.ToDouble(XlsDataSh.Items[System.Convert.ToInt16(n)].SubItems[k].Text);
                    }
                    RenewCenterArrayParams[i, k - 3] = (tmpSameDis * 1.0 / (ClusterAssem[i].Count+1));
                }
            }
        }

        private void OutPExBtn_Click(object sender, EventArgs e)
        {
            richTextBox1.Text += "\n正在导出Excel";
            //OrderCenter();
            //GradeJudge();
            //FillClus();
            //CreateXls();
            //MessageBox.Show("C1:" + GradeHalfMax[0]);
        }

        private void OrderCenter() {
            double[] tmpMax = new double[ClassNum];
            double tmpOrder = 0;

            for (int i = 0; i < ClassNum; i++) {
                tmpMax[i] = 0;
            }

            for (int i = 0; i < ClassNum; i++)
            {
                tmpMax[i] = (CenterArrayParams[i, 0] + CenterArrayParams[i, 1] + CenterArrayParams[i, 2] + CenterArrayParams[i, 3] + CenterArrayParams[i, 4]) * 1.0 / 5;
            }

            //MessageBox.Show(tmpMax[0].ToString() + "|" + tmpMax[1].ToString() + "|" + tmpMax[2].ToString() + "|" + tmpMax[3].ToString() + "|" + tmpMax[4].ToString());

            tmpOrder = tmpMax[0];
            for (int i = 0; i < ClassNum; i++) {
                for (int j = 0; j < ClassNum; j++) {
                    if (tmpMax[j] > tmpOrder)
                    {
                        OrderClass[i] = j;
                        tmpOrder = tmpMax[j];
                    }
                }
                tmpOrder = 0;
                tmpMax[OrderClass[i]] = 0;
            }

            for (int i = 0; i < ClassNum; i++) {
                BaseGrade[OrderClass[i]] = 100 - 10 * i - 5;
            }

            richTextBox1.Text += ("\n排序簇结果(降序):\n" + (OrderClass[0]+1).ToString() + "|" + (OrderClass[1]+1).ToString() + "|" + (OrderClass[2]+1).ToString() + "|" + (OrderClass[3]+1).ToString() + "|" + (OrderClass[4]+1).ToString());
        }

        private void GradeJudge() {
            FindMaxDis();

            double tmpDis = 0;

            for (int i = 0; i < ClassNum; i++) {
                foreach (object Clus in ClusterAssem[i]) {
                    tmpDis = 0;
                    for (int k = 0;k < Dimension; k++) {
                        tmpDis += Math.Pow((System.Convert.ToDouble(XlsDataSh.Items[System.Convert.ToInt16(Clus)].SubItems[k + 3].Text) - CenterArrayParams[i, k]), 2);
                    }
                    //XlsDataSh.Items[System.Convert.ToInt16(Clus)].SubItems[8].Text = System.Convert.ToString(((Math.Pow(tmpDis,1.0/2) - GradeHalfMax[i])*5) + BaseGrade[i]);
                    //MessageBox.Show(GradeHalfMax[i].ToString() + "," + Math.Pow(tmpDis, 1.0 / 2));
                    XlsDataSh.Items[System.Convert.ToInt16(Clus)].SubItems[Dimension+4].Text = System.Convert.ToString((Math.Pow(tmpDis,1.0/2)- 2 * GradeHalfMax[i])/GradeFurMax[i] * 5 + BaseGrade[i]);
                }
            }
        }

        private void FindMaxDis() {
            double tmpFurmMax = 0, tmpFurPer = 0; ;

            for (int i = 0; i < ClassNum; i++)
            {
                tmpFurmMax = 0;
                foreach (object Clus in ClusterAssem[i])
                {
                    tmpFurPer = 0;
                    for (int k = 0; k < Dimension; k++)
                    {
                        tmpFurPer += Math.Pow((System.Convert.ToDouble(XlsDataSh.Items[System.Convert.ToInt16(Clus)].SubItems[k + 3].Text) - CenterArrayParams[i, k]), 2);
                    }
                    if (Math.Pow(tmpFurPer, 1.0 / 2) > tmpFurmMax)
                    {
                        tmpFurmMax = tmpFurPer;
                    }
                }
                GradeFurMax[i] = Math.Pow(tmpFurmMax, 1.0/2);
                GradeHalfMax[i] = Math.Pow(tmpFurmMax, 1.0 / 2) / 2;
            }
        }

        private void FillClus() {
            for (int i = 0; i < ClassNum; i++)
            {
                foreach (object Clus in ClusterAssem[i])
                {
                    XlsDataSh.Items[System.Convert.ToInt16(Clus)].SubItems[Dimension + 3].Text = (i+1).ToString();
                }
            }
        }
        private void CreateXls() {
            HSSFWorkbook wk = new HSSFWorkbook();
            string tmpsavepath;

            ISheet tb = wk.CreateSheet("聚类结果");
            IRow row = tb.CreateRow(0);
            ICell tmpcell = row.CreateCell(0);
            tmpcell.SetCellValue("编号");
            tmpcell = row.CreateCell(1);
            tmpcell.SetCellValue("学号");
            tmpcell = row.CreateCell(2);
            tmpcell.SetCellValue("姓名");
            tmpcell = row.CreateCell(3);
            tmpcell.SetCellValue("计划书");
            tmpcell = row.CreateCell(4);
            tmpcell.SetCellValue("小组评价");
            tmpcell = row.CreateCell(5);
            tmpcell.SetCellValue("小组报告");
            tmpcell = row.CreateCell(6);
            tmpcell.SetCellValue("小组互评");
            tmpcell = row.CreateCell(7);
            tmpcell.SetCellValue("记录得分");
            tmpcell = row.CreateCell(8);
            tmpcell.SetCellValue("第几类");
            tmpcell = row.CreateCell(9);
            tmpcell.SetCellValue("分数");
            for (int i = 1; i <= RowCount - 1;i++) {
                row = tb.CreateRow(i);
                for (int j = 0; j < 3 + Dimension + 2;j++) {
                    ICell cell = row.CreateCell(j);
                    cell.SetCellValue(XlsDataSh.Items[i].SubItems[j].Text);
                }
            }

            XLsSaveDlg.FileName = "聚类成绩单";
            XLsSaveDlg.Filter = "Xls files(*.xls)|*.xls";
            XLsSaveDlg.DefaultExt = "xls";
            XLsSaveDlg.AddExtension = true;

            DialogResult result = XLsSaveDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                tmpsavepath = XLsSaveDlg.FileName.ToString();
            }
            else
            {
                return;
            }
            FileStream fs = File.OpenWrite(tmpsavepath);
            wk.Write(fs);
            fs.Close();
            richTextBox1.Text += "\n已经创建成功.";
        }

        private void LOFDist() {
            //xlxw.ver
            LOFk = 20;     //文章中的k
            maxLof = 3;    //文章中的阈值

            NkLOF = new int[RowCount, LOFk];
            lrdk = new double[RowCount];
            LOFK = new double[RowCount];
            double tmpLofSum = 0;

            for (int i = 0; i < RowCount - 1; i++) {
                Nk(i);  //确定数据的领域Nk(p)以及计算出lrdk
            }
            richTextBox1.Text += "\nLOF计算中";
            //计算LOFK
            for (int i = 0; i < RowCount; i++) {
                tmpLofSum = 0;
                for (int j = 0; j < LOFk; j++) {
                    tmpLofSum += lrdk[NkLOF[i, j]]/ lrdk[i];
                }
                LOFK[i] = tmpLofSum / LOFk;
            }
            richTextBox1.Text += "\n分析得到的离群点:";
            //LOFK显示在ListView
            for (int i = 0; i < RowCount - 1; i++) {
                if (LOFK[i] > maxLof) richTextBox1.Text += "\n第" + (i+1).ToString() + "个数据";
                XlsDataSh.Items[i].SubItems[8].Text = LOFK[i].ToString();
            }
        }

        private void Nk(int DataPtS) {
            double[] TmpDis = new double[RowCount];
            double tmpdis = 0,tmpNpKMax = 0,tmpSum = 0;

            //计算所有的距离
            for (int i = 0; i < RowCount - 1; i++) {
                tmpdis = 0;
                for (int j = 3; j < 7; j++) {
                    tmpdis += Math.Pow((System.Convert.ToDouble(XlsDataSh.Items[i].SubItems[j].Text) - (System.Convert.ToDouble(XlsDataSh.Items[DataPtS].SubItems[j].Text))), 2);
                }
                TmpDis[i] = Math.Pow(tmpdis, 1.0/2);
            }

            double[] copy = new double[TmpDis.Length];
            TmpDis.CopyTo(copy, 0);

            //确定数据的领域Nk(p)
            for (int i = 0; i < LOFk; i++) {
                tmpNpKMax = 0;
                for (int j = 0; (j < RowCount - 1); j++) {
                    if (copy[j] > tmpNpKMax) {
                        tmpNpKMax = copy[j];
                        NkLOF[DataPtS, i] = j;
                    }
                }
                copy[NkLOF[DataPtS, i]] = 0;
            }

            //计算出lrdk
            for (int i = 0; i < LOFk; i++) {
                tmpSum += TmpDis[NkLOF[DataPtS, i]];
            }
            lrdk[DataPtS] = LOFk / tmpSum;
        }
    }
}
