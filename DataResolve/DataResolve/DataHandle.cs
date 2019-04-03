using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataResolve.Model;
using NPOI.XSSF.UserModel;

namespace DataResolve
{
    public class DataHandle
    {
        private string FileName = string.Empty;
        private IWorkbook workbook = null;
        private FileStream fs = null;

        // 起始坐标
        private int StartIndex = 3;
        // 几个点构成一个折线图
        private int PointCount = 8;
        // 几个点构成一组折线图
        private int LineCount = 4;

        public DataHandle(string filename)
        {
            this.FileName = filename;
        }

        public IEnumerable<TypeGra> GetData(out string errorInfo)
        {
            errorInfo = string.Empty;
            fs = new FileStream(FileName, FileMode.Open, FileAccess.Read);
            workbook = new XSSFWorkbook(fs);
            List<Point> PointList = new List<Point>();
            List<LineGra> LineList = new List<LineGra>();
            List<TypeGra> TypeList = new List<TypeGra>();
            try
            {
                ISheet sheet = null;
                sheet = workbook.GetSheet("ALL");
                int CountAll = sheet.LastRowNum;
                int start = StartIndex;

                IRow row;
                for (int lop1 = StartIndex; lop1 <= CountAll; lop1++)
                {
                    row = sheet.GetRow(lop1);
                    if ((lop1 - start + 1) % PointCount == 0)
                    {
                        //生成点
                        int index = lop1 % PointCount - StartIndex;
                        Point point = GeneratePoint(row, index);
                        PointList.Add(point);
                        LineGra Line = new LineGra();
                        Line.Line = PointList;
                        LineList.Add(Line);
                        if ((lop1 - start + 1) % (PointCount * LineCount) == 0)
                        {
                            TypeGra type = new TypeGra();
                            type.Name = lop1.ToString();
                            type.Type = LineList;
                            TypeList.Add(type);
                            LineList = new List<LineGra>();
                        }
                        PointList = new List<Point>();
                    }
                    else
                    {
                        //生成点
                        int index = lop1 % PointCount - StartIndex;
                        Point point = GeneratePoint(row, index);
                        PointList.Add(point);
                    }
                }

            }
            catch (Exception ex)
            {
                errorInfo = ex.Message;
            }
            return TypeList;
        }

        private decimal GetValue(IRow row, int index)
        {
            decimal ret;
            ret = row.GetCell(index).ToString() == "NA" || row.GetCell(index).ToString() == "NF" ? 0M : Convert.ToDecimal(row.GetCell(index).ToString());
            return ret;
        }

        private Point GeneratePoint(IRow row, int index)
        {
            Point point = new Point()
            {
                Index = index,
                ACS = GetValue(row, 4),
                BTA = GetValue(row, 8),
                CAN = GetValue(row, 12),
                CBZ = GetValue(row, 16),
                DCF = GetValue(row, 20),
                FAA = GetValue(row, 24),
                GAB = GetValue(row, 28),
                GPL = GetValue(row, 32),
                IOM = GetValue(row, 36),
                IOP = GetValue(row, 40),
                MBT = GetValue(row, 44),
                MTP = GetValue(row, 48),
                OLM = GetValue(row, 52),
                PRL = GetValue(row, 56),
                SMX = GetValue(row, 60),
                VAL = GetValue(row, 64),
                VLX = GetValue(row, 68),
                VSA = GetValue(row, 72)
            };

            return point;
        }

    }
}
