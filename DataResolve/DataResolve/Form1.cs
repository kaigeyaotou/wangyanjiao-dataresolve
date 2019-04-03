using DataResolve.Model;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace DataResolve
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            LoadLine();
        }

        private void LoadLine()
        {
            this.chart1.Palette = ChartColorPalette.EarthTones;
            System.Windows.Forms.DataVisualization.Charting.Series series = new System.Windows.Forms.DataVisualization.Charting.Series();
            series.BorderColor = Color.Green;
            series.ChartType = SeriesChartType.Line;
            series.Enabled = true;
            series.Points.AddXY("September", 100);
            series.Points.AddXY("Obtober", 300);
            series.Points.AddXY("November", 800);
            series.Points.AddXY("December", 200);
            series.Points.AddXY("January", 600);
            series.Points.AddXY("February", 400);
            this.chart1.Series.Add(series);


            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            series1.BorderColor = Color.Yellow;
            series1.ShadowColor = Color.Black;
            series1.Color = Color.Green;
            series1.ChartType = SeriesChartType.Line;
            series1.Enabled = true;
            series1.Points.AddXY("September", 110);
            series1.Points.AddXY("Obtober", 310);
            series1.Points.AddXY("November", 810);
            series1.Points.AddXY("December", 210);
            series1.Points.AddXY("January", 610);
            series1.Points.AddXY("February", 410);
            this.chart1.Series.Add(series1);
        }

        private void ShowLine(List<TypeGra> typeGra, string Type)
        {
            this.chart1.Series.Clear();
            Series series = new Series();
            series.Name = "hello world";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filename = @"E:\1.xlsx";
            DataHandle dataHandle = new DataHandle(filename);
            string errorInfo = string.Empty;
            var list = dataHandle.GetData(out errorInfo);
        }
    }
}
