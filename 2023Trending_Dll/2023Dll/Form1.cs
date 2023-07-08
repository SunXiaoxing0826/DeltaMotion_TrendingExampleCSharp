using RMCLinkNET;
using System;
using System.IO;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Trend2023;

namespace _2023Dll
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// .NET Framwork 4.8
        /// 安装RMCLINK
        /// Plot Samples 设定大于512 可避免读取时丢点
        /// Plot Samples RMC75:%MD 32+N.1 RMC150:%MD 96+N.1 RMC200 :%MD 512+N.1
        /// Trend.DLL 默认设置ReadSamples=512 控制器内Plot Samples设置大于ReadSamples
        /// </summary>

        #region Data

        public RMCLink RMC;
        public static string IP = "192.168.0.3";
        public float[] dataF = new float[17];
        public float[] dataF1 = new float[1];
        public float[] dataF2 = new float[1];
        public float[] dataF3 = new float[2];

        Trending Trender = null;
        public static string RMCType = "RMC70";
        public static DeviceType RMCTypes;
        public int sampleSpan = 0;
        public int sampleSpacing = 0;
        public int NumberOfPlotTrending = 0;
        public static int PlotNumber = 0;
        double XVal = 0.0, YVal1 = 0.0;
        float CachedSamplePeriod = 0;

        int CachedDataSetCount = 16;
        int CurrentSampleCount = 0;

        string select = "";
        FileStream fs = null;
        StreamWriter sw = null;

        Random rd = new Random();
        string save_inter_char = "/";
        #endregion

        private void UpdateData()
        {
            switch (comboBox1.SelectedIndex)
            {
                case 0: RMCTypes = DeviceType.RMC70; break;
                case 1: RMCTypes = DeviceType.RMC150; break;
                case 2: RMCTypes = DeviceType.RMC200; break;
                default: RMCTypes = DeviceType.RMC70; break;
            }

            PlotNumber = int.Parse(textBox2.Text);
        }
        private void StartTrending()
        {
            try
            {
                Trender = new Trending(RMCTypes, IP, PlotNumber, 200);
                Trender.StartTrend();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "unable to connect");
                Trender = null;
                return;
            }
            if (Trender.Running)
            {
                CachedSamplePeriod = Trender.PlotSamplePeriod;
                timer1.Interval = 125;
                timer1.Enabled = true;
                timer1.Start();
            }
        }

        private void StopTrending()
        {
            if (Trender != null && Trender.Running)
            {
                Trender.StopTrend();
            }
            timer1.Stop();
        }

        public void insertDataToMSChart(Chart chart)
        {
                if (Trender == null)
                {
                    return;
                }
                sampleSpan = int.Parse(textBox3.Text);
                sampleSpacing = 100;
                sampleSpacing = Trender.DataSetCount / 4;  
                if (sampleSpacing < 1)
                    sampleSpacing = 1;
                lock (Trender.TrendDataLock)
                {
                    int CurrentSampleCount = Trender.TrendData[0].Count;

                    sampleSpan = sampleSpan > CurrentSampleCount ? CurrentSampleCount : sampleSpan;
                    for (int i = 0; i < CachedDataSetCount; i++)
                    {
                        chart.Series[i].Points.Clear();
                    }
                    
                    for (int j = (CurrentSampleCount - sampleSpan); j < CurrentSampleCount - sampleSpacing; j += sampleSpacing)
                    {
                        int k = j, l = k + sampleSpacing;
                        if (k < Trender.TrendData[1].Count && l < Trender.TrendData[1].Count)//?
                        {
                            XVal = Math.Round(Trender.TrendData[0][j] / CachedSamplePeriod) / 1000;
                            for (int n = 1; n < CachedDataSetCount + 1; n++)
                            {
                                YVal1 = Trender.TrendData[n][j];
                                chart.Series[n-1].Points.AddXY(XVal, YVal1);
                            }
                            
                        }
                    }
                }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            chart1.Series.Clear();
            chart1.Legends.Clear();         
            chart1.Legends.Clear();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
            switch (comboBox1.SelectedIndex)
            {
                case 0: RMCTypes = DeviceType.RMC70;  break;
                case 1: RMCTypes = DeviceType.RMC150; break;
                case 2: RMCTypes = DeviceType.RMC200; break;
                default: RMCTypes = DeviceType.RMC70; break;
            };
            IP = textBox1.Text;
            PlotNumber = int.Parse(textBox2.Text);

            if (Trender != null && Trender.Running)
            {

            }
            else
            {
                chart1.Series.Clear();
                chart1.Legends.Clear();
                chart1.Legends.Clear();
               
                UpdateData();
                StartTrending();

                CachedDataSetCount = Trender.DataSetCount;

                for (int i = 0; i < CachedDataSetCount; i++)
                {
                    Legend legend = new Legend();
                    legend.Name = "曲线" + i.ToString();

                    Series series = new Series();
                    series.ChartType = SeriesChartType.Line;
                    series.Name = i.ToString();
                    
                    chart1.Series.Add(series);
                    chart1.Legends.Add(legend);
                    chart1.ChartAreas[0].AxisY.IsStartedFromZero = false;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (Trender != null && Trender.Running)
            {
                StopTrending();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (Trender != null)
            {
                if (Trender.Running)
                {
                    if (MessageBox.Show("保存数据操作会终止曲线运行？", "确认", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        StopTrending();
                        SaveData();
                    }
                    
                }
                else
                {
                    SaveData();
                }
            }
        }

        private void SaveData()
        {
            select = Application.StartupPath + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".txt";
            fs = new FileStream(select, FileMode.Create);
            sw = new StreamWriter(fs);

            for (int i = 0; i < Trender.TrendData[0].Count; i = i + 1)
            {
                for (int j = 0; j < CachedDataSetCount + 1; j++)
                {
                    if (j == CachedDataSetCount) save_inter_char = "";
                    else save_inter_char = "/";
                    sw.Write(Trender.TrendData[j][i].ToString("0.000") + save_inter_char);
                }
                sw.WriteLine();
            }

            sw.Flush();

            //关闭流
            sw.Close();
            fs.Close();

            MessageBox.Show("试验已完成，数据已保存" + select);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (Trender != null && Trender.Running)
            {
                if (WindowState != FormWindowState.Minimized) insertDataToMSChart(chart1);

                textBox4.Text = CachedDataSetCount.ToString();
                textBox5.Text = CachedSamplePeriod.ToString();

            }
            else
            {
                StopTrending();
            }
        }
    }
}
