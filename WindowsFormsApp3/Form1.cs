using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing.Printing;

using Microsoft.Win32;
using System.Runtime.InteropServices;



using System.Data.SqlClient;
using MySql.Data.MySqlClient;  //참조로 추가한다

namespace WindowsFormsApp3
{
    
    public partial class Form1 : Form
    {
        class data
        {
            public int room;
            public string name;
            public int sum;
            public int price;
            public int total_price;


            public float normal_prvele;
            public float normal_curele;
            public float normal_usage;
            public float normal_price;
            public int normal_money;

            public float night_prvele;
            public float night_curele;
            public float night_usage;
            public float night_price;
            public int night_money;

            public int pipe_heatingcost;
            public int expendables;
            public int medical_refund;
            public int phonebill;
            public int cablefee;
            public int preliving_expenses;
            public int living_expenses;
            public int 선납잔액;
        }

        class 미수금_선납금data
        {
            public int room;
            public string name;

            public int 미수금_선납금;
        }


        Dictionary<int, data> dataMap = new Dictionary<int, data>();
        Dictionary<int, 미수금_선납금data> 미수금_선납금dataMap = new Dictionary<int, 미수금_선납금data>();

        List<data> dataList = new List<data>();
        List<미수금_선납금data> 미수금_선납금dataList = new List<미수금_선납금data>();


        struct out_data
        {
            public string name;
            public string hosu;
            public string published_date; //발행일
            public string published_month;
            public string prvmeter_normal;//전월검침 일반
            public string prvmeter_night; //전월검침 심야
            public string curmeter_normal;//금월검침 일반
            public string curmeter_night; //금월검침 심야
            public string usage_normal;   //사용량   일반;
            public string usage_night;    //사용량   심야;
            public string kw_normal;      //단가(kw) 일반;
            public string kw_night;       //단가(kw) 심야;  
            public string electbill_normal;//전기요금 일반;
            public string electbill_night; //전기요금 심야;
            public string electbill_sum;   //전기요금 합계
            public string public_water;    //공용수도광열비;
            public string food_expenses;   //식비
            public string food_expenses_sum;//식비합계
            public string expendables;      //소모품
            public string medical_expense_refund;//의료비환급금
            public string phone_bill;      //전화비
            public string premonth_living_expenses;//전월 생활비
            public string month_living_expenses;//월 생활비
            public string cable_expense;   //케이블이용료
            public string 소계;
            public string 관리비;
            public string 청구금액;
            public string 선납금액;

            public out_data(string i_name,
                            string i_hosu,
                            string i_published_date,
                            string i_published_month,
                            string i_prvmeter_normal,
                            string i_prvmeter_night,
                            string i_curmeter_normal,
                            string i_curmeter_night,
                            string i_usage_normal,
                            string i_usage_night,
                            string i_kw_normal,
                            string i_kw_night,
                            string i_electbill_normal,
                            string i_electbill_night,
                            string i_electbill_sum,
                            string i_public_water,
                            string i_food_expenses,
                            string i_food_expenses_sum,
                            string i_expendables,
                            string i_medical_expense_refund,
                            string i_phone_bill,
                            string i_premonth_living_expenses,
                            string i_month_living_expenses,
                            string i_cable_expense,
                            string i_소계,
                            string i_관리비,
                            string i_청구금액,
                            string i_선납금액)
            {
                name = i_name;
                hosu = i_hosu;
                published_date = i_published_date;
                published_month = i_published_month;
                prvmeter_normal = i_prvmeter_normal;
                prvmeter_night = i_prvmeter_night;
                curmeter_normal = i_curmeter_normal;
                curmeter_night = i_curmeter_night;
                usage_normal = i_usage_normal;
                usage_night = i_usage_night;
                kw_normal = i_kw_normal;
                kw_night = i_kw_night;
                electbill_normal = i_electbill_normal;
                electbill_night = i_electbill_night;
                electbill_sum = i_electbill_sum;
                public_water = i_public_water;
                food_expenses = i_food_expenses;
                food_expenses_sum = i_food_expenses_sum;
                expendables = i_expendables;
                medical_expense_refund = i_medical_expense_refund;
                phone_bill = i_phone_bill;
                premonth_living_expenses = i_premonth_living_expenses;
                month_living_expenses = i_month_living_expenses;
                cable_expense = i_cable_expense;
                소계 = i_소계;
                관리비 = i_관리비;
                청구금액 = i_청구금액;
                선납금액 = i_선납금액;
            }
        };



        //
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        object misValue;
        Excel.Range chartRange;


        object[,] filddata;

        // object[,] data;
        int curYear = 0;
        int curMonth = 0;
        int curDay = 0;

        public Form1()
        {

            InitializeComponent();

            //
            /*
            DataTable table = new DataTable();

            // column을 추가합니다.
            table.Columns.Add("ID", typeof(string));
            table.Columns.Add("제목", typeof(string));
            table.Columns.Add("구분", typeof(string));
            table.Columns.Add("생성일", typeof(string));
            table.Columns.Add("수정일", typeof(string));

            // 각각의 행에 내용을 입력합니다.
            table.Rows.Add("ID 1", "제목 1번", "사용중", "2019/03/11", "2019/03/18");
            table.Rows.Add("ID 2", "제목 2번", "미사용", "2019/03/12", "2019/03/18");
            table.Rows.Add("ID 3", "제목 3번", "미사용", "2019/03/13", "2019/03/18");
            table.Rows.Add("ID 4", "제목 4번", "사용중", "2019/03/14", "2019/03/18");

            // 값들이 입력된 테이블을 DataGridView에 입력합니다.
           // dataGridView1.DataSource = table;
    //
    */
            curYear = DateTime.Now.ToLocalTime().Year;
            int nYear = curYear;
            curMonth = DateTime.Now.ToLocalTime().Month - 1;

            for (int i = 0; i < 2; i++)
            {
                comboBox_year.Items.Add(nYear.ToString());
                nYear--;
            }

            if (curMonth == 0)
            {
                curMonth = 12;
                nYear--;
            }

            int day = DateTime.Now.ToLocalTime().Day;

            comboBox_year.SelectedItem = curYear.ToString();
            comboBox_month.SelectedItem = curMonth.ToString();
            comboBox_day.SelectedItem = day.ToString();

            object misValue = System.Reflection.Missing.Value;

            

        }

        void file_export(string path)
        {
            xlWorkSheet.Cells[1, 1] = "구분";
            xlWorkSheet.Cells[2, 2] = "성명";
            xlWorkSheet.Cells[2, 3] = "호실";


            xlWorkSheet.Cells[1, 4] = "일반전기";
            xlWorkSheet.Cells[2, 4] = "전월검침";
            xlWorkSheet.Cells[2, 5] = "금월검침";
            xlWorkSheet.Cells[2, 6] = "사용량";
            xlWorkSheet.Cells[2, 7] = "금액";
            xlWorkSheet.Cells[2, 8] = "단가";

            xlWorkSheet.Cells[1, 9] = "심야전기";
            xlWorkSheet.Cells[2, 9] = "전월검침";
            xlWorkSheet.Cells[2, 10] = "금월검침";
            xlWorkSheet.Cells[2, 11] = "사용량";
            xlWorkSheet.Cells[2, 12] = "금액";
            xlWorkSheet.Cells[2, 13] = "단가";

            xlWorkSheet.Cells[1, 14] = "식비";
            xlWorkSheet.Cells[2, 14] = "식수";
            xlWorkSheet.Cells[2, 15] = "단가";
            xlWorkSheet.Cells[2, 16] = "금액";

            xlWorkSheet.Cells[1, 17] = "공용수도관열비";

            xlWorkSheet.Cells[1, 18] = "소모품";
            xlWorkSheet.Cells[2, 18] = "금액";

            xlWorkSheet.Cells[1, 19] = "의료환급금";
            
            xlWorkSheet.Cells[1, 20] = "전화요금";

            xlWorkSheet.Cells[1, 21] = "케이블이용료";

            xlWorkSheet.Cells[1, 22] = "전월 생활비";
            xlWorkSheet.Cells[1, 23] = "생활비";

            xlWorkSheet.Cells[1, 24] = "선납 잔액";


            xlWorkSheet.get_Range("a1", "c1").Merge(false);
            xlWorkSheet.get_Range("d1", "h1").Merge(false);
            xlWorkSheet.get_Range("i1", "m1").Merge(false);
            xlWorkSheet.get_Range("n1", "p1").Merge(false);
            xlWorkSheet.get_Range("q1", "q2").Merge(false);

            xlWorkSheet.get_Range("s1", "s2").Merge(false);
            xlWorkSheet.get_Range("t1", "t2").Merge(false);
            xlWorkSheet.get_Range("u1", "u2").Merge(false);
            xlWorkSheet.get_Range("v1", "v2").Merge(false);
            xlWorkSheet.get_Range("w1", "w2").Merge(false);
            xlWorkSheet.get_Range("x1", "x2").Merge(false);


            chartRange = xlWorkSheet.get_Range("a1", "x2");
            chartRange.Interior.Color = System.Drawing.Color.FromArgb(164, 199, 243);

            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 2;

            chartRange.Font.Bold = false;
            chartRange.Font.Size = 10;
      
            for(int i = 1; i <= 2; i++)
            {
                chartRange = xlWorkSheet.get_Range("a" + i, "a" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange.EntireColumn.ColumnWidth = 8;
                chartRange = xlWorkSheet.get_Range("b" + i, "b" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange.EntireColumn.ColumnWidth = 15;
                chartRange = xlWorkSheet.get_Range("c" + i, "c" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("d" + i, "d" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("e" + i, "e" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("f" + i, "f" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("g" + i, "g" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("h" + i, "h" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("i" + i, "i" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("j" + i, "j" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("k" + i, "k" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("l" + i, "l" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("m" + i, "m" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("n" + i, "n" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("o" + i, "o" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                chartRange = xlWorkSheet.get_Range("p" + i, "p" + i);
                if (i < 3) chartRange.Interior.Color = System.Drawing.Color.FromArgb(164, 199, 243);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                
                chartRange = xlWorkSheet.get_Range("q" + i, "q" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange.EntireColumn.ColumnWidth = 15;

                chartRange = xlWorkSheet.get_Range("r" + i, "r" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                chartRange = xlWorkSheet.get_Range("s" + i, "s" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
               // chartRange.EntireColumn.ColumnWidth = 15;

                chartRange = xlWorkSheet.get_Range("t" + i, "t" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                chartRange = xlWorkSheet.get_Range("u" + i, "u" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                //chartRange.EntireColumn.ColumnWidth = 15;

                chartRange = xlWorkSheet.get_Range("v" + i, "v" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange.EntireColumn.ColumnWidth = 15;


                chartRange = xlWorkSheet.get_Range("w" + i, "w" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange.EntireColumn.ColumnWidth = 15;

                chartRange = xlWorkSheet.get_Range("x" + i, "x" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange.EntireColumn.ColumnWidth = 15;
            }

            int check = 0;
            int line = 3;
            foreach (KeyValuePair<int, data> row in dataMap)
            {
                xlWorkSheet.Cells[line, 1] = line - 2;
                xlWorkSheet.Cells[line, 2] = row.Value.name;
                xlWorkSheet.Cells[line, 3] = row.Value.room;

                if (check == 0 && row.Value.name == null)
                    check = line - 1;

                string normal_money = string.Format("{0:#,0}", row.Value.normal_money);
                string normal_price = string.Format("{0:#,0}", row.Value.normal_price);

                string night_money = string.Format("{0:#,0}", row.Value.night_money);
                string night_price = string.Format("{0:#,0}", row.Value.night_price);

                string sum = string.Format("{0:#,0}", row.Value.sum);
                string price = string.Format("{0:#,0}", row.Value.price);
                string total_price = string.Format("{0:#,0}", row.Value.total_price);

                string pipe_heatingcost = string.Format("{0:#,0}", row.Value.pipe_heatingcost);
                string expendables = string.Format("{0:#,0}", row.Value.expendables);
                string medical_refund = string.Format("{0:#,0}", row.Value.medical_refund);
                string phonebill = string.Format("{0:#,0}", row.Value.phonebill);
                string cablefee = string.Format("{0:#,0}", row.Value.cablefee);
                string living_expenses = string.Format("{0:#,0}", row.Value.living_expenses);
                string preliving_expenses = string.Format("{0:#,0}", row.Value.preliving_expenses);
                string 선납잔액 = string.Format("{0:#,0}", row.Value.선납잔액);


                xlWorkSheet.Cells[line, 4] = row.Value.normal_prvele;//일반전기 전월
                xlWorkSheet.Cells[line, 5] = row.Value.normal_curele;//일반전기 금월
                xlWorkSheet.Cells[line, 6] = row.Value.normal_usage;//일반전기 사용량
                xlWorkSheet.Cells[line, 7] = normal_money;//일반전기 금액
                xlWorkSheet.Cells[line, 8] = normal_price;//일반전기 단가

                xlWorkSheet.Cells[line, 9] = row.Value.night_prvele;//심야전기 전월
                xlWorkSheet.Cells[line, 10] = row.Value.night_curele;//심야전기 금월
                xlWorkSheet.Cells[line, 11] = row.Value.night_usage;//심야전기 사용량
                xlWorkSheet.Cells[line, 12] = night_money;//심야전기 금액
                xlWorkSheet.Cells[line, 13] = night_price;//심야전기 단가

                xlWorkSheet.Cells[line, 14] = sum;//식수
                xlWorkSheet.Cells[line, 15] = price;//단가
                xlWorkSheet.Cells[line, 16] = total_price;//금액

                xlWorkSheet.Cells[line, 17] = pipe_heatingcost;//공용수도관열비
                xlWorkSheet.Cells[line, 18] = expendables;//소모품
                xlWorkSheet.Cells[line, 19] = medical_refund;//의료환급금
                xlWorkSheet.Cells[line, 20] = phonebill;//전화요금
                xlWorkSheet.Cells[line, 21] = cablefee;//케이블이용료
                xlWorkSheet.Cells[line, 22] = preliving_expenses;//전월생활비
                xlWorkSheet.Cells[line, 23] = living_expenses;//생활비

                xlWorkSheet.Cells[line, 24] = 선납잔액;


                /*
                int i = line;

                chartRange = xlWorkSheet.get_Range("a" + i, "a" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange.EntireColumn.ColumnWidth = 8;
                chartRange = xlWorkSheet.get_Range("b" + i, "b" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange.EntireColumn.ColumnWidth = 15;
                chartRange = xlWorkSheet.get_Range("c" + i, "c" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("d" + i, "d" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("e" + i, "e" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("f" + i, "f" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("g" + i, "g" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("h" + i, "h" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("i" + i, "i" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("j" + i, "j" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("k" + i, "k" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("l" + i, "l" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("m" + i, "m" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("n" + i, "n" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange = xlWorkSheet.get_Range("o" + i, "o" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                chartRange = xlWorkSheet.get_Range("p" + i, "p" + i);
                if (i < 3) chartRange.Interior.Color = System.Drawing.Color.FromArgb(164, 199, 243);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                chartRange = xlWorkSheet.get_Range("q" + i, "q" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                //chartRange.EntireColumn.ColumnWidth = 15;

                chartRange = xlWorkSheet.get_Range("r" + i, "r" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                chartRange = xlWorkSheet.get_Range("s" + i, "s" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
               // chartRange.EntireColumn.ColumnWidth = 15;

                chartRange = xlWorkSheet.get_Range("t" + i, "t" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                chartRange = xlWorkSheet.get_Range("u" + i, "u" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
               // chartRange.EntireColumn.ColumnWidth = 15;

                chartRange = xlWorkSheet.get_Range("v" + i, "v" + i);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
               // chartRange.EntireColumn.ColumnWidth = 15;
               */


                line++;
            }

            xlWorkSheet.Application.StandardFont = "HY견고딕";

            chartRange = xlWorkSheet.get_Range("q3", "w"+ check);
            chartRange.Interior.Color = System.Drawing.Color.FromArgb(166, 255, 188);

            xlWorkBook.SaveAs(path);
          
            MessageBox.Show("저장되었습니다.");
        }

        private void output_info(out_data data,int dir = 0 ,int interval = 0)
        {
            int fontsize = 14;
            int fontsize2 = 12;

            //
            //add data 
            int bottom_inter = 4;

            xlWorkSheet.Cells[6, 3 + interval] = curMonth + "월 선납 잔액";
            xlWorkSheet.Cells[6, 6 + interval] = data.선납금액;

            xlWorkSheet.Cells[8, 3 + interval] = "구분";
            xlWorkSheet.Cells[8, 4 + interval] = "일반";
            xlWorkSheet.Cells[8, 5 + interval] = "심야";
            xlWorkSheet.Cells[8, 6 + interval] = "합계";

            xlWorkSheet.Cells[9, 3 + interval] = "전월검침";
            xlWorkSheet.Cells[9, 4 + interval] = data.prvmeter_normal;
            xlWorkSheet.Cells[9, 5 + interval] = data.prvmeter_night;

            xlWorkSheet.Cells[10, 3 + interval] = "금월검침";
            xlWorkSheet.Cells[10, 4 + interval] = data.curmeter_normal;
            xlWorkSheet.Cells[10, 5 + interval] = data.curmeter_night;

            xlWorkSheet.Cells[11, 3 + interval] = "사용량";
            xlWorkSheet.Cells[11, 4 + interval] = data.usage_normal;
            xlWorkSheet.Cells[11, 5 + interval] = data.usage_night;

            xlWorkSheet.Cells[12, 3 + interval] = "단가(KW당)";
            xlWorkSheet.Cells[12, 4 + interval] = data.kw_normal;
            xlWorkSheet.Cells[12, 5 + interval] = data.kw_night;

            xlWorkSheet.Cells[13, 3 + interval] = "전기요금";
            xlWorkSheet.Cells[13, 4 + interval] = data.electbill_normal;
            xlWorkSheet.Cells[13, 5 + interval] = data.electbill_night;

            xlWorkSheet.Cells[14, 3 + interval] = "공용수도광열비";
            xlWorkSheet.Cells[14, 6 + interval] = data.public_water == "0" ? "-" : data.public_water; 

            xlWorkSheet.Cells[15, 3 + interval] = "식비";
            xlWorkSheet.Cells[15, 4 + interval] = data.food_expenses;
            xlWorkSheet.Cells[15, 6 + interval] = data.food_expenses_sum;

            xlWorkSheet.Cells[16, 3 + interval] = "소모품";
            xlWorkSheet.Cells[16, 6 + interval] = data.expendables == "0"? "-":data.expendables;

            xlWorkSheet.Cells[17, 3 + interval] = "의료비환급금";
            xlWorkSheet.Cells[17, 6 + interval] = data.medical_expense_refund == "0"? "-":data.medical_expense_refund;

            xlWorkSheet.Cells[18, 3 + interval] = "전화사용료";
            xlWorkSheet.Cells[18, 6 + interval] = data.phone_bill == "0" ? "-": data.phone_bill;

            xlWorkSheet.Cells[19, 3 + interval] = "케이블이용료";
            xlWorkSheet.Cells[19, 6 + interval] = data.cable_expense == "0" ? "-" : data.cable_expense;

            xlWorkSheet.Cells[20, 3 + interval] = "(월)생활비";
            xlWorkSheet.Cells[20, 6 + interval] = data.premonth_living_expenses;

            xlWorkSheet.Cells[21, 3 + interval] = "소계";
            xlWorkSheet.Cells[21, 6 + interval] = data.소계 == "0" ? "-" : data.소계;

            string text = "이<br/><br/><br/>용<br/><br/><br/>내<br/><br/><br/>역";
            string textWithNewLine = text.Replace("<br/>", Environment.NewLine);

            xlWorkSheet.Cells[8, 2 + interval] = textWithNewLine;
            xlWorkSheet.Cells[8, 2 + interval].Style.WrapText = true;

            chartRange = xlWorkSheet.get_Range("b8", "b21");
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 2;

            chartRange = xlWorkSheet.get_Range("j8", "j21");
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 2;


            text = "청<br/>구<br/>내<br/>역";
            textWithNewLine = text.Replace("<br/>", Environment.NewLine);

            xlWorkSheet.Cells[23, 2 + interval] = textWithNewLine;
            xlWorkSheet.Cells[23, 2 + interval].Style.WrapText = true;

            chartRange = xlWorkSheet.get_Range("b23", "b25");
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 2;
            chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            chartRange = xlWorkSheet.get_Range("j23", "j25");
            chartRange.HorizontalAlignment = 3;
            chartRange.VerticalAlignment = 2;
            chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);


            xlWorkSheet.Cells[23, 3 + interval] = curMonth + "월 관리비";
            xlWorkSheet.Cells[23, 6 + interval] = data.관리비;

            int tempMonth = curMonth + 1;
            if (tempMonth > 12)
                tempMonth = 1;

            xlWorkSheet.Cells[24, 3 + interval] = tempMonth + "월 생활비";
            xlWorkSheet.Cells[24, 6 + interval] = data.month_living_expenses;

            xlWorkSheet.Cells[25, 3 + interval] = "청구(미납)금액";
            xlWorkSheet.Cells[25, 6 + interval] = data.청구금액;




            Microsoft.Office.Interop.Excel.Range oRange = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[25 + bottom_inter, 4 + interval];
            float Left = (float)((double)oRange.Left) + 30;
            float Top = (float)((double)oRange.Top);

            string path = System.IO.Directory.GetCurrentDirectory() + "\\logo.jpg";

            xlWorkSheet.Shapes.AddPicture( path, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Left, Top, 150, 43);

            
            for (int i = 8; i < 22; i++)
            {
                if (dir == 0)
                {
                    xlWorkSheet.get_Range("c" + i, "c" + i).Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    xlWorkSheet.get_Range("d" + i, "d" + i).Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    xlWorkSheet.get_Range("e" + i, "e" + i).Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    xlWorkSheet.get_Range("f" + i, "f" + i).Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                    chartRange = xlWorkSheet.get_Range("c" + i, "f" + i);
                }
                else
                {
                    xlWorkSheet.get_Range("k" + i, "k" + i).Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    xlWorkSheet.get_Range("l" + i, "l" + i).Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    xlWorkSheet.get_Range("m" + i, "m" + i).Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    xlWorkSheet.get_Range("n" + i, "n" + i).Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                    chartRange = xlWorkSheet.get_Range("k" + i, "n" + i);
                }


                chartRange.EntireColumn.ColumnWidth = 17;
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                chartRange.Font.Bold = true;
                chartRange.Font.Size = fontsize;
            }

            

            for (int i = 23; i < 26; i++)
            {
                if (dir == 0)
                {
                    xlWorkSheet.get_Range("c" + i, "c" + i).Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    xlWorkSheet.get_Range("d" + i, "d" + i).Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    xlWorkSheet.get_Range("e" + i, "e" + i).Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    xlWorkSheet.get_Range("f" + i, "f" + i).Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                    chartRange = xlWorkSheet.get_Range("c" + i, "f" + i);
                }
                else
                {
                    xlWorkSheet.get_Range("k" + i, "k" + i).Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    xlWorkSheet.get_Range("l" + i, "l" + i).Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    xlWorkSheet.get_Range("m" + i, "m" + i).Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    xlWorkSheet.get_Range("n" + i, "n" + i).Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                    chartRange = xlWorkSheet.get_Range("k" + i, "n" + i);
                }


                chartRange.EntireColumn.ColumnWidth = 17;
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                chartRange.Font.Bold = true;
                chartRange.Font.Size = fontsize;
            }

            if (dir == 0)
            {
                xlWorkSheet.get_Range("b6", "e6").Merge(false);
    
                xlWorkSheet.get_Range("b1", "f2").Merge(false);
                xlWorkSheet.get_Range("b3", "f3").Merge(false);
                xlWorkSheet.get_Range("b4", "f4").Merge(false);
                xlWorkSheet.get_Range("b5", "f5").Merge(false);
                xlWorkSheet.get_Range("f9", "f13").Merge(false);
                xlWorkSheet.get_Range("d14", "e14").Merge(false);
                xlWorkSheet.get_Range("d15", "e15").Merge(false);

                xlWorkSheet.get_Range("d16", "e16").Merge(false);
                xlWorkSheet.get_Range("d17", "e17").Merge(false);
                xlWorkSheet.get_Range("d18", "e18").Merge(false);
                xlWorkSheet.get_Range("d19", "e19").Merge(false);
                xlWorkSheet.get_Range("d20", "e20").Merge(false);
                xlWorkSheet.get_Range("d21", "e21").Merge(false);

                xlWorkSheet.get_Range("b22", "f22").Merge(false);

                xlWorkSheet.get_Range("b23", "b25").Merge(false);
                xlWorkSheet.get_Range("d23", "e23").Merge(false);
                xlWorkSheet.get_Range("d24", "e24").Merge(false);
                xlWorkSheet.get_Range("d25", "e25").Merge(false);

                xlWorkSheet.get_Range("b23", "f25").Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);


                xlWorkSheet.get_Range("b8", "b21").Merge(false);
                xlWorkSheet.get_Range("b8", "b21").Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);


                for (int i = 0; i < 5; i++)
                {
                    xlWorkSheet.get_Range("b" + (26 + i).ToString(), "f" + (26 + i).ToString()).Merge(false);
                }


                chartRange = xlWorkSheet.get_Range("c8", "f8");
                chartRange.Interior.Color = System.Drawing.Color.FromArgb(226, 156, 54);

                chartRange = xlWorkSheet.get_Range("c9", "e9");
                chartRange.Interior.Color = System.Drawing.Color.FromArgb(84, 155, 205);

                chartRange = xlWorkSheet.get_Range("c10", "e10");
                chartRange.Interior.Color = System.Drawing.Color.FromArgb(84, 155, 205);

                chartRange = xlWorkSheet.get_Range("f14", "f21");
                chartRange.Interior.Color = System.Drawing.Color.FromArgb(226, 156, 54);

                chartRange = xlWorkSheet.get_Range("f23", "f25");
                chartRange.Interior.Color = System.Drawing.Color.FromArgb(226, 156, 54);

                chartRange = xlWorkSheet.get_Range("b6", "e6");
                chartRange.Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                chartRange.Font.Bold = true;
                chartRange.Font.Size = fontsize;


                chartRange = xlWorkSheet.get_Range("f6", "f6");
                chartRange.Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                chartRange.Font.Bold = true;
                chartRange.Font.Size = fontsize;

                chartRange = xlWorkSheet.get_Range("a1", "a26");
                chartRange.EntireColumn.ColumnWidth = 1;

                chartRange = xlWorkSheet.get_Range("b1", "b26");
                chartRange.EntireColumn.ColumnWidth = 5;


                chartRange = xlWorkSheet.get_Range("g1", "g26");
                chartRange.EntireColumn.ColumnWidth = 1;

                chartRange = xlWorkSheet.get_Range("f9", "f9");
                chartRange.FormulaR1C1 = "소수점\r\n 이하\r\n 반올림함";
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                chartRange.Font.Bold = true;
                chartRange.Font.Size = fontsize;

                chartRange = xlWorkSheet.get_Range("b1", "f2");
                chartRange.FormulaR1C1 = "영 수 증(고객보관용)";
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                //chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                //chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                chartRange.Font.Size = 24;
                chartRange.Font.Bold = true;

                chartRange = xlWorkSheet.get_Range("b3", "f3");
                chartRange.FormulaR1C1 = "( " + data.hosu + " ) 호 /號室  성명/性名     :" + data.name;
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                //chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                //chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                chartRange.Font.Size = fontsize;
                chartRange.Font.Bold = true;

                string[] dates = data.published_date.Split('-');

                chartRange = xlWorkSheet.get_Range("b4", "f4");
                chartRange.FormulaR1C1 = "발행일: " + dates[0] +  "年" + dates[1] + "月" + dates[2] + "日(" + data.published_month + "월분)";
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                //chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                //chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                chartRange.Font.Size = fontsize;
                chartRange.Font.Bold = true;

                int day = DateTime.DaysInMonth(Convert.ToInt32(dates[0]), Convert.ToInt32(dates[1]));

                chartRange = xlWorkSheet.get_Range("b5", "f5");
                chartRange.FormulaR1C1 = "납부기한: " + dates[1] + "월" + day + "일까지";
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                //chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                //chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                chartRange.Font.Size = fontsize;
                chartRange.Font.Bold = true;

                chartRange = xlWorkSheet.get_Range("b" + (22 + bottom_inter).ToString(), "f" + (22 + bottom_inter).ToString());
                chartRange.FormulaR1C1 = "위와 같이 통지합니다(100원단위 절사)";
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                //chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                //chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                chartRange.Font.Size = fontsize2;
                chartRange.Font.Bold = false;

                chartRange = xlWorkSheet.get_Range("b" + (23 + bottom_inter).ToString(), "f" + (23 + bottom_inter).ToString());
                chartRange.FormulaR1C1 = "수납기관 : 우리은행 효정설악점";
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                //chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                //chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                chartRange.Font.Size = fontsize2;
                chartRange.Font.Bold = false;

                chartRange = xlWorkSheet.get_Range("b" + (24 + bottom_inter).ToString(), "f" + (24 + bottom_inter).ToString());
                chartRange.FormulaR1C1 = "계좌번호 : 우리은행 1005-488-999555";
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                //chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                //chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                chartRange.Font.Size = fontsize2;
                chartRange.Font.Bold = false;

                chartRange = xlWorkSheet.get_Range("a1", "g" + (26 + bottom_inter).ToString());
                chartRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            }
            else
            {
                xlWorkSheet.get_Range("j6", "m6").Merge(false);

                xlWorkSheet.get_Range("i1", "n2").Merge(false);
                xlWorkSheet.get_Range("i3", "n3").Merge(false);
                xlWorkSheet.get_Range("i4", "n4").Merge(false);
                xlWorkSheet.get_Range("i5", "n5").Merge(false);


                xlWorkSheet.get_Range("n9", "n13").Merge(false);
                xlWorkSheet.get_Range("l14", "m14").Merge(false);
                xlWorkSheet.get_Range("l15", "m15").Merge(false);


                xlWorkSheet.get_Range("l16", "m16").Merge(false);
                xlWorkSheet.get_Range("l17", "m17").Merge(false);
                xlWorkSheet.get_Range("l18", "m18").Merge(false);
                xlWorkSheet.get_Range("l19", "m19").Merge(false);
                xlWorkSheet.get_Range("l20", "m20").Merge(false);
                xlWorkSheet.get_Range("l21", "m21").Merge(false);

                xlWorkSheet.get_Range("j22", "n22").Merge(false);

                xlWorkSheet.get_Range("j23", "j25").Merge(false);
                xlWorkSheet.get_Range("l23", "m23").Merge(false);
                xlWorkSheet.get_Range("l24", "m24").Merge(false);
                xlWorkSheet.get_Range("l25", "m25").Merge(false);

                xlWorkSheet.get_Range("j23", "n25").Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);


                xlWorkSheet.get_Range("j8", "j21").Merge(false);
                xlWorkSheet.get_Range("j8", "j21").Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);


                for (int i = 0; i < 5; i++)
                {
                    xlWorkSheet.get_Range("i" + (26 + i).ToString(), "n" + (26 + i).ToString()).Merge(false);
                }

                chartRange = xlWorkSheet.get_Range("k8", "n8");
                chartRange.Interior.Color = System.Drawing.Color.FromArgb(226, 156, 54);

                chartRange = xlWorkSheet.get_Range("k9", "m9");
                chartRange.Interior.Color = System.Drawing.Color.FromArgb(84, 155, 205);

                chartRange = xlWorkSheet.get_Range("k10", "m10");
                chartRange.Interior.Color = System.Drawing.Color.FromArgb(84, 155, 205);

                chartRange = xlWorkSheet.get_Range("n14", "n21");
                chartRange.Interior.Color = System.Drawing.Color.FromArgb(226, 156, 54);

                chartRange = xlWorkSheet.get_Range("n23", "n25");
                chartRange.Interior.Color = System.Drawing.Color.FromArgb(226, 156, 54);

                
                chartRange = xlWorkSheet.get_Range("j6", "m6");
                chartRange.Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                chartRange.Font.Bold = true;
                chartRange.Font.Size = fontsize;


                chartRange = xlWorkSheet.get_Range("n6", "n6");
                chartRange.Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);
                chartRange.Cells.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                chartRange.Font.Bold = true;
                chartRange.Font.Size = fontsize;


                chartRange = xlWorkSheet.get_Range("i1", "i26");
                chartRange.EntireColumn.ColumnWidth = 1;

                chartRange = xlWorkSheet.get_Range("j1", "j26");
                chartRange.EntireColumn.ColumnWidth = 5;


                chartRange = xlWorkSheet.get_Range("o1", "o26");
                chartRange.EntireColumn.ColumnWidth = 1;

                chartRange = xlWorkSheet.get_Range("n1", "n26");
                chartRange.EntireColumn.ColumnWidth = 17;


                chartRange = xlWorkSheet.get_Range("n9", "n9");
                chartRange.FormulaR1C1 = "소수점\r\n 이하\r\n 반올림함";
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                chartRange.Font.Bold = true;
                chartRange.Font.Size = fontsize;

                chartRange = xlWorkSheet.get_Range("i1", "l2");
                chartRange.FormulaR1C1 = "영 수 증(빌리지보관용)";
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                //chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                //chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                chartRange.Font.Size = 24;
                chartRange.Font.Bold = true;

                chartRange = xlWorkSheet.get_Range("i3", "l3");
                chartRange.FormulaR1C1 = "( " + data.hosu + " ) 호 /號室  성명/性名     :" + data.name;
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                //chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                //chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                chartRange.Font.Size = fontsize;
                chartRange.Font.Bold = true;

                string[] dates = data.published_date.Split('-');
                chartRange = xlWorkSheet.get_Range("i4", "l4");
                chartRange.FormulaR1C1 = "발행일: " + dates[0] + "年" + dates[1] + "月" + dates[2] + "日(" + data.published_month + "월분)";
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                //chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                //chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                chartRange.Font.Size = fontsize;
                chartRange.Font.Bold = true;

                int day = DateTime.DaysInMonth(Convert.ToInt32(dates[0]), Convert.ToInt32(dates[1]));

                chartRange = xlWorkSheet.get_Range("i5", "l5");
                chartRange.FormulaR1C1 = "납부기한: " + dates[1] + "월" + day + "일까지";
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                //chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                //chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                chartRange.Font.Size = fontsize;
                chartRange.Font.Bold = true;

                
                chartRange = xlWorkSheet.get_Range("i" + (22 + bottom_inter).ToString(), "l" + (22 + bottom_inter).ToString());
                chartRange.FormulaR1C1 = "위와 같이 통지합니다(100원단위 절사)";
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                //chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                //chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                chartRange.Font.Size = fontsize2;
                chartRange.Font.Bold = false;

                chartRange = xlWorkSheet.get_Range("i" + (23 + bottom_inter).ToString(), "l" + (23 + bottom_inter).ToString());
                chartRange.FormulaR1C1 = "수납기관 : 우리은행 효정설악점";
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                //chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                //chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                chartRange.Font.Size = fontsize2;
                chartRange.Font.Bold = false;

                chartRange = xlWorkSheet.get_Range("i" + (24 + bottom_inter).ToString(), "l" + (24 + bottom_inter).ToString());
                chartRange.FormulaR1C1 = "계좌번호 : 우리은행 1005-488-999555";
                chartRange.HorizontalAlignment = 3;
                chartRange.VerticalAlignment = 2;
                //chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                //chartRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                chartRange.Font.Size = fontsize2;
                chartRange.Font.Bold = false;

                chartRange = xlWorkSheet.get_Range("i1", "o" + (26 + bottom_inter).ToString());
                chartRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            }

            xlWorkSheet.Cells.Rows.RowHeight = 26;
            chartRange = xlWorkSheet.get_Range("g1", "g25");
            chartRange.EntireColumn.ColumnWidth = 1;
            chartRange = xlWorkSheet.get_Range("h1", "h25");
            chartRange.EntireColumn.ColumnWidth = 1;
            xlWorkSheet.Cells[7, 1].RowHeight = 10;

            chartRange = xlWorkSheet.get_Range("f14", "f25");
            chartRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

            chartRange = xlWorkSheet.get_Range("n14", "n25");
            chartRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

            /*

            chartRange = xlWorkSheet.get_Range("b4", "e4");
            chartRange.Font.Bold = true;
            chartRange = xlWorkSheet.get_Range("b9", "e9");
            chartRange.Font.Bold = true;
            */

        }

        public static string ConvertToExcelPrinterFriendlyName(string printerName)
        {
            var key = Registry.CurrentUser;
            var subkey = key.OpenSubKey(@"Software\Microsoft\Windows NT\CurrentVersion\Devices");

            var value = subkey.GetValue(printerName);
            if (value == null) throw new Exception(string.Format("Device not found: {0}", printerName));

            var portName = value.ToString().Substring(9);  //strip away the winspool, 

            return string.Format("{0}에 있는 {1}", portName , printerName); ;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            dataMap.Clear();
            meals_DB();
            electricity_DB();
            common_DB();
            moneyreceived_DB();

            if (dataMap.Count() == 0)
            {
                MessageBox.Show("ERROR " + curYear + "년 " + curMonth + "월 데이터가 없습니다.");

                return;
            }

            if (MessageBox.Show("선택하신 " + curYear.ToString() + "년 " + curMonth.ToString() + "월 영수증을 출력 합니까?", "영수증", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                PrintDialog dialog = new PrintDialog();
                dialog.AllowPrintToFile = true;
                dialog.AllowCurrentPage = true;
                dialog.AllowSomePages = true;
                dialog.AllowSelection = true;
                dialog.UseEXDialog = true;
                dialog.PrinterSettings.Duplex = Duplex.Simplex;
                dialog.PrinterSettings.FromPage = 0;
                dialog.PrinterSettings.ToPage = 8;
                dialog.PrinterSettings.PrintRange = PrintRange.SomePages;

                //Call ShowDialog  
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    xlWorkSheet.Application.StandardFont = "HY견고딕";
                    Cursor.Current = Cursors.WaitCursor;

                    var excelPrinterName = ConvertToExcelPrinterFriendlyName(dialog.PrinterSettings.PrinterName);
                    xlApp.ActivePrinter = excelPrinterName;

                    int totalCount = 0;
                    foreach (DataGridViewRow r in dataGridView1.Rows)
                    {
                        if (r.Cells["CK"].Value != null && (bool)r.Cells["CK"].Value == true)
                        {
                            totalCount++;
                        }
                    }

                    int num = 0, nPage = 0;
                    foreach (KeyValuePair<int, data> row in dataMap)
                    {
                        num++;
                        if (dataGridView1.Rows[num - 1].Cells["CK"].Value == null)
                        {
                            continue;
                        }

                        int tempYear = curYear;
                        int tempMonth = curMonth + 1;
                        if (tempMonth == 13)
                        {
                            tempYear++;
                            tempMonth = 1;
                        }

                        string name = row.Value.name;
                        string room = row.Value.room.ToString();
                        string date = tempYear + "-" + tempMonth + "-" + curDay;

                        int 소계 = row.Value.pipe_heatingcost + row.Value.total_price + row.Value.expendables + row.Value.phonebill +
                            row.Value.preliving_expenses + row.Value.cablefee + row.Value.normal_money + row.Value.night_money;
                        소계 = (소계 / 1000);
                        소계 *= 1000;

                        int 관리비 = row.Value.pipe_heatingcost + row.Value.total_price + row.Value.expendables + row.Value.phonebill +
                            row.Value.cablefee + row.Value.normal_money + row.Value.night_money;
                        관리비 = (관리비 / 1000);
                        관리비 *= 1000;

                        int 청구금액 = 관리비 + row.Value.living_expenses + row.Value.선납잔액;
                        if (청구금액 < 0)
                            청구금액 = 0;
                        else
                        {
                            if (관리비 + row.Value.living_expenses < 청구금액)
                            {
                                청구금액 = 관리비 + row.Value.living_expenses;
                            }
                        }
                        

                        nPage++;
                        print_num.Text = nPage + "/" + totalCount;

                        if (name == null)
                            continue;

                        output_info(new out_data(
                            name,
                            room,               //호수
                            date,               //발행일
                            curMonth.ToString(),//월분
                            string.Format("{0:#,0}", row.Value.normal_prvele),   //전월검침 일반
                            string.Format("{0:#,0}", row.Value.night_prvele),    //전월검침 심야
                            string.Format("{0:#,0}", row.Value.normal_curele),   //금월검침 일반
                            string.Format("{0:#,0}", row.Value.night_curele),    //금월검침 심야
                            string.Format("{0:#,0}", row.Value.normal_usage),    //사용량 일반
                            string.Format("{0:#,0}", row.Value.night_usage),     //사용량 심야
                            string.Format("{0:#,0}", row.Value.normal_price),    //단가 일반
                            string.Format("{0:#,0}", row.Value.night_price),     //단가 심야
                            string.Format("{0:#,0}", row.Value.normal_money),    //전기요금 일반
                            string.Format("{0:#,0}", row.Value.night_money),     //전기요금 심야
                            "",   //전기요금 합계
                            string.Format("{0:#,0}", row.Value.pipe_heatingcost),//공용수도광열비
                            string.Format("{0:#,0}", row.Value.sum),             //식비 
                            string.Format("{0:#,0}", row.Value.total_price),     //식비 합계
                            string.Format("{0:#,0}", row.Value.expendables),     //소모품
                            string.Format("{0:#,0}", row.Value.medical_refund),  //의료비환급금
                            string.Format("{0:#,0}", row.Value.phonebill),       //전화사용료
                            string.Format("{0:#,0}", row.Value.preliving_expenses), //전달생활비
                            string.Format("{0:#,0}", row.Value.living_expenses), //생활비

                            string.Format("{0:#,0}", row.Value.cablefee),        //케이블이용료
                            string.Format("{0:#,0}", 소계), 
                            string.Format("{0:#,0}", 관리비),
                            string.Format("{0:#,0}", 청구금액),
                            string.Format("{0:#,0}", -row.Value.선납잔액)
                            ));

                        output_info(new out_data(
                           name,
                           room,               //호수
                           date,               //발행일
                           curMonth.ToString(),//월분
                           string.Format("{0:#,0}", row.Value.normal_prvele),   //전월검침 일반
                            string.Format("{0:#,0}", row.Value.night_prvele),    //전월검침 심야
                            string.Format("{0:#,0}", row.Value.normal_curele),   //금월검침 일반
                            string.Format("{0:#,0}", row.Value.night_curele),    //금월검침 심야
                            string.Format("{0:#,0}", row.Value.normal_usage),    //사용량 일반
                            string.Format("{0:#,0}", row.Value.night_usage),     //사용량 심야
                            string.Format("{0:#,0}", row.Value.normal_price),    //단가 일반
                            string.Format("{0:#,0}", row.Value.night_price),     //단가 심야
                            string.Format("{0:#,0}", row.Value.normal_money),    //전기요금 일반
                            string.Format("{0:#,0}", row.Value.night_money),     //전기요금 심야
                            "",   //전기요금 합계
                            string.Format("{0:#,0}", row.Value.pipe_heatingcost),//공용수도광열비
                            string.Format("{0:#,0}", row.Value.sum),             //식비 
                            string.Format("{0:#,0}", row.Value.total_price),     //식비 합계
                            string.Format("{0:#,0}", row.Value.expendables),     //소모품
                            string.Format("{0:#,0}", row.Value.medical_refund),  //의료비환급금
                            string.Format("{0:#,0}", row.Value.phonebill),       //전화사용료
                            string.Format("{0:#,0}", row.Value.preliving_expenses), //전달생활비
                            string.Format("{0:#,0}", row.Value.living_expenses), //생활비
                            string.Format("{0:#,0}", row.Value.cablefee),        //케이블이용료
                            string.Format("{0:#,0}", 소계),
                            string.Format("{0:#,0}", 관리비),
                            string.Format("{0:#,0}", 청구금액),
                            string.Format("{0:#,0}", -row.Value.선납잔액)
                           ), 1, 8);


                        xlWorkSheet.Application.StandardFont = "HY견고딕";



                        //Set margins
                        xlWorkSheet.PageSetup.TopMargin = 0;
                        xlWorkSheet.PageSetup.BottomMargin = 0;
                        xlWorkSheet.PageSetup.LeftMargin = 0;
                        xlWorkSheet.PageSetup.RightMargin = 0;

                        //Set Center on page
                        xlWorkSheet.PageSetup.CenterHorizontally = true;
                        xlWorkSheet.PageSetup.CenterVertically = false;


                        xlWorkSheet.PageSetup.FitToPagesWide = 1;
                        xlWorkSheet.PageSetup.FitToPagesTall = 100;

                        xlWorkSheet.PageSetup.Zoom = 72;





                        xlWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

                        xlWorkSheet.PrintOut();

                        //break;
                    }






                    //xlWorkSheet.PrintOut();
                    //dlg.Document.Print();


                    //xlWorkBook.SaveAs("d:\\csharp.net-informations.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    //xlWorkBook.Close(true, misValue, misValue);
                    //xlApp.Quit();

                    //releaseObject(xlApp);
                    //releaseObject(xlWorkBook);
                    //releaseObject(xlWorkSheet);

                    Cursor.Current = Cursors.Default;

                    //xlWorkBook.Close(null, null, null);                 // close your workbook
                    //xlApp.Quit();                                   // exit excel application

                    MessageBox.Show("File created !");
                }
            }
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        static void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                   // Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

        

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            if (cb.SelectedIndex > -1)
            {
                curMonth = int.Parse(cb.SelectedItem.ToString());
            }
        }

        void meals_DB()
        {
            string query = "select * from VA_meals where vm_year = " + curYear + " AND vm_month = " + curMonth;

            MySqlConnection conn = null;
            string strConn = "Server=222.101.9.218;Database=village_admin;Uid=viluser;Pwd=viluser@csit;";
            using (conn = new MySqlConnection(strConn))
            {
                conn.Open();

                using (MySqlCommand command = new MySqlCommand(query, conn))
                {
                    //command.Parameters.AddWithValue("@val", val);

                    //ExecuteReader를 이용하여
                    //연결 모드로 데이타 가져오기
                    MySqlDataReader rdr = command.ExecuteReader();
                    while (rdr.Read())
                    {
                        int room = int.Parse(rdr["vm_room"].ToString());
                        string name = rdr["vm_name"].ToString();
                        int sum  = int.Parse(rdr["vm_sum"].ToString());
                        int price = int.Parse(rdr["vm_price"].ToString());
                        int total_price = int.Parse(rdr["vm_totalprice"].ToString());

                        data data;
                        if (dataMap.TryGetValue(room ,out data))
                        {
                            data.sum += sum;
                            data.name += ("/" + name);
                            data.total_price += total_price;
                        }
                        else
                        {
                            data = new data();

                            data.sum = sum;
                            data.name = name;
                            data.price = price;
                            data.total_price = total_price;
                        }

                        data.room = room;


                        dataMap[room] = data;
                    }


                }

                conn.Close();
            }
        }

        void electricity_DB()
        {
            string query = "select * from VA_electricity where ve_year = " + curYear + " AND ve_month = " + curMonth;

            MySqlConnection conn = null;
            string strConn = "Server=222.101.9.218;Database=village_admin;Uid=viluser;Pwd=viluser@csit;";
            using (conn = new MySqlConnection(strConn))
            {
                conn.Open();

                using (MySqlCommand command = new MySqlCommand(query, conn))
                {
                    //command.Parameters.AddWithValue("@val", val);

                    //ExecuteReader를 이용하여
                    //연결 모드로 데이타 가져오기
                    MySqlDataReader rdr = command.ExecuteReader();
                    while (rdr.Read())
                    {
                        int room = int.Parse(rdr["ve_room"].ToString());
                        float normal_premonth = float.Parse(rdr["ve_normal_premonth"].ToString());
                        float normal_curmonth = float.Parse(rdr["ve_normal_curmonth"].ToString());
                        float normal_usage = float.Parse(rdr["ve_normal_usage"].ToString());
                        int normal_money = int.Parse(rdr["ve_normal_money"].ToString());
                        float normal_price = float.Parse(rdr["ve_normal_price"].ToString());

                        float night_premonth = float.Parse(rdr["ve_night_premonth"].ToString());
                        float night_curmonth = float.Parse(rdr["ve_night_curmonth"].ToString());
                        float night_usage = float.Parse(rdr["ve_night_usage"].ToString());
                        int night_money = int.Parse(rdr["ve_night_money"].ToString());
                        float night_price = float.Parse(rdr["ve_night_price"].ToString());

                        data data;
                        dataMap.TryGetValue(room, out data);
                        if(data == null)
                        {
                            data = new data();
                            data.room = room;
                        }
                        
                        data.normal_prvele = normal_premonth;
                        data.normal_curele = normal_curmonth;
                        data.normal_usage = normal_usage;
                        data.normal_money = normal_money;
                        data.normal_price = normal_price;

                        data.night_prvele = night_premonth;
                        data.night_curele = night_curmonth;
                        data.night_usage = night_usage;
                        data.night_money = night_money;
                        data.night_price = night_price;

                        dataMap[room] = data;
                    }


                }

                conn.Close();
            }
        }

        void common_DB()
        {
            string query = "select * from VA_common where vc_year = " + curYear + " AND vc_month = " + curMonth;

            MySqlConnection conn = null;
            string strConn = "Server=222.101.9.218;Database=village_admin;Uid=viluser;Pwd=viluser@csit;";
            using (conn = new MySqlConnection(strConn))
            {
                conn.Open();

                using (MySqlCommand command = new MySqlCommand(query, conn))
                {
                    //command.Parameters.AddWithValue("@val", val);

                    //ExecuteReader를 이용하여
                    //연결 모드로 데이타 가져오기
                    MySqlDataReader rdr = command.ExecuteReader();
                    while (rdr.Read())
                    {
                        int room = int.Parse(rdr["vc_room"].ToString());
                        int pipe_heatingcost = int.Parse(rdr["vc_pipe_heatingcost"].ToString());
                        int expendables = int.Parse(rdr["vc_expendables"].ToString());
                        int medical_refund = int.Parse(rdr["vc_medical_refund"].ToString());
                        int phonebill = int.Parse(rdr["vc_phonebill"].ToString());
                        int cablefee = int.Parse(rdr["vc_cablefee"].ToString());

                        int living_expenses = 0;
                        if (rdr["vc_living_expenses"].ToString() != "")
                            living_expenses = int.Parse(rdr["vc_living_expenses"].ToString());

                        int preliving_expenses = 0;
                        if (rdr["vc_preliving_expenses"].ToString() != "")
                            preliving_expenses = int.Parse(rdr["vc_preliving_expenses"].ToString());

                        data data;
                        dataMap.TryGetValue(room, out data);
                        if (data == null)
                        {
                            data = new data();
                            data.room = room;
                        }

                        data.pipe_heatingcost = pipe_heatingcost;
                        data.expendables = expendables;
                        data.medical_refund = medical_refund;
                        data.phonebill = phonebill;
                        data.cablefee = cablefee;

                        data.living_expenses = living_expenses;
                        data.preliving_expenses = preliving_expenses;

                        dataMap[room] = data;
                    }


                }

                conn.Close();
            }
        }

        void moneyreceived_DB()
        {
            string query = "select * from VA_moneyreceived where vm_year = " + curYear + " AND vm_month = " + curMonth;

            MySqlConnection conn = null;
            string strConn = "Server=222.101.9.218;Database=village_admin;Uid=viluser;Pwd=viluser@csit;";
            using (conn = new MySqlConnection(strConn))
            {
                conn.Open();

                using (MySqlCommand command = new MySqlCommand(query, conn))
                {
                    //command.Parameters.AddWithValue("@val", val);

                    //ExecuteReader를 이용하여
                    //연결 모드로 데이타 가져오기
                    MySqlDataReader rdr = command.ExecuteReader();
                    while (rdr.Read())
                    {
                        int room = int.Parse(rdr["vm_room"].ToString());
                        int money = int.Parse(rdr["vm_money"].ToString());

                        
                        data data;
                        dataMap.TryGetValue(room, out data);
                        if (data == null)
                        {
                            data = new data();
                            data.room = room;
                        }

                        data.선납잔액 = money;

                        dataMap[room] = data;
                    }


                }

                conn.Close();
            }
        }

        void pre_common_DB()
        {
            int tempYear = curYear;
            int tempMonth = curMonth - 1;
            if (tempMonth == 0)
            {
                tempMonth = 12;
                tempYear = curYear - 1;
            }

            string query = "select * from VA_common where vc_year = " + tempYear + " AND vc_month = " + tempMonth;

            MySqlConnection conn = null;
            string strConn = "Server=222.101.9.218;Database=village_admin;Uid=viluser;Pwd=viluser@csit;";
            using (conn = new MySqlConnection(strConn))
            {
                conn.Open();

                using (MySqlCommand command = new MySqlCommand(query, conn))
                {
                    //command.Parameters.AddWithValue("@val", val);

                    //ExecuteReader를 이용하여
                    //연결 모드로 데이타 가져오기
                    MySqlDataReader rdr = command.ExecuteReader();
                    while (rdr.Read())
                    {
                        int room = int.Parse(rdr["vc_room"].ToString());
                        int pipe_heatingcost = int.Parse(rdr["vc_pipe_heatingcost"].ToString());
                        int expendables = int.Parse(rdr["vc_expendables"].ToString());
                        int medical_refund = int.Parse(rdr["vc_medical_refund"].ToString());
                        int phonebill = int.Parse(rdr["vc_phonebill"].ToString());
                        int cablefee = int.Parse(rdr["vc_cablefee"].ToString());
                        int living_expenses = int.Parse(rdr["vc_living_expenses"].ToString());

                        data data;
                        if (dataMap.TryGetValue(room, out data))
                        {
                            if (pipe_heatingcost != 0 && data.pipe_heatingcost == 0)
                                data.pipe_heatingcost = pipe_heatingcost;

                            if (expendables != 0 && data.expendables == 0)
                              data.expendables = expendables;

                            if (medical_refund != 0 && data.medical_refund == 0)
                                data.medical_refund = medical_refund;

                            if (phonebill != 0 && data.phonebill == 0)
                                data.phonebill = phonebill;

                            if (cablefee != 0 && data.cablefee == 0)
                                data.cablefee = cablefee;

                            if (living_expenses != 0 && data.preliving_expenses == 0)
                                data.preliving_expenses = living_expenses;
                            
                            if (living_expenses != 0 && data.living_expenses == 0)
                                data.living_expenses = living_expenses;
                        }

                        dataMap[room] = data;
                    }


                }

                conn.Close();
            }
        }

        void delete_data()
        {
            string query = "DELETE FROM VA_common where vc_year = @vc_year AND vc_month = @vc_month";

            MySqlConnection conn = null;

            string strConn = "Server=222.101.9.218;Database=village_admin;Uid=viluser;Pwd=viluser@csit;";
            using (conn = new MySqlConnection(strConn))
            {
                conn.Open();

                using (MySqlCommand command = new MySqlCommand(query, conn))
                {
                    command.Parameters.AddWithValue("@vc_year", curYear);
                    command.Parameters.AddWithValue("@vc_month", curMonth);

                    //conn.Open();
                    int result = command.ExecuteNonQuery();

                    // Check Error
                    if (result < 0)
                        Console.WriteLine("Error inserting data into Database!");
                }

                conn.Close();
            }
        }

        void 미수금_선납금delete_data()
        {
            string query = "DELETE FROM VA_moneyreceived where vm_year = @vm_year AND vm_month = @vm_month";

            MySqlConnection conn = null;

            string strConn = "Server=222.101.9.218;Database=village_admin;Uid=viluser;Pwd=viluser@csit;";
            using (conn = new MySqlConnection(strConn))
            {
                conn.Open();

                using (MySqlCommand command = new MySqlCommand(query, conn))
                {
                    command.Parameters.AddWithValue("@vm_year", curYear);
                    command.Parameters.AddWithValue("@vm_month", curMonth);

                    //conn.Open();
                    int result = command.ExecuteNonQuery();

                    // Check Error
                    if (result < 0)
                        Console.WriteLine("Error inserting data into Database!");
                }

                conn.Close();
            }
        }

        void insert_data(data data)
        {
            string query = "INSERT INTO VA_common (vc_year,vc_month,vc_room,vc_pipe_heatingcost,vc_expendables,vc_medical_refund,vc_phonebill,vc_cablefee,vc_living_expenses,vc_preliving_expenses) VALUES (@vc_year,@vc_month,@vc_room,@vc_pipe_heatingcost,@vc_expendables,@vc_medical_refund,@vc_phonebill,@vc_cablefee,@vc_living_expenses,@vc_preliving_expenses)";

            MySqlConnection conn = null;

            string strConn = "Server=222.101.9.218;Database=village_admin;Uid=viluser;Pwd=viluser@csit;";
            using (conn = new MySqlConnection(strConn))
            {
                conn.Open();

                using (MySqlCommand command = new MySqlCommand(query, conn))
                {
                    command.Parameters.AddWithValue("@vc_year", curYear);
                    command.Parameters.AddWithValue("@vc_month", curMonth);

                    command.Parameters.AddWithValue("@vc_room", data.room);
                    command.Parameters.AddWithValue("@vc_pipe_heatingcost", data.pipe_heatingcost);
                    command.Parameters.AddWithValue("@vc_expendables", data.expendables);
                    command.Parameters.AddWithValue("@vc_medical_refund", data.medical_refund);
                    command.Parameters.AddWithValue("@vc_phonebill", data.phonebill);
                    command.Parameters.AddWithValue("@vc_cablefee", data.cablefee);
                    command.Parameters.AddWithValue("@vc_living_expenses", data.living_expenses);
                    command.Parameters.AddWithValue("@vc_preliving_expenses", data.preliving_expenses);

                    //conn.Open();
                    int result = command.ExecuteNonQuery();

                    // Check Error
                    if (result < 0)
                        Console.WriteLine("Error inserting data into Database!");
                }

                conn.Close();
            }
        }

        void 미수금_선납금insert_data(미수금_선납금data data)
        {
            string query = "INSERT INTO VA_moneyreceived (vm_year,vm_month,vm_room,vm_money) VALUES (@vm_year,@vm_month,@vm_room,@vm_money)";

            MySqlConnection conn = null;

            string strConn = "Server=222.101.9.218;Database=village_admin;Uid=viluser;Pwd=viluser@csit;";
            using (conn = new MySqlConnection(strConn))
            {
                conn.Open();

                using (MySqlCommand command = new MySqlCommand(query, conn))
                {
                    command.Parameters.AddWithValue("@vm_year", curYear);
                    command.Parameters.AddWithValue("@vm_month", curMonth);

                    command.Parameters.AddWithValue("@vm_room", data.room);
                    command.Parameters.AddWithValue("@vm_money", data.미수금_선납금);

                    //conn.Open();
                    int result = command.ExecuteNonQuery();

                    // Check Error
                    if (result < 0)
                        Console.WriteLine("Error inserting data into Database!");
                }

                conn.Close();
            }
        }

        static string get_선납금(string year ,string month ,string room)
        {
            string strConn = "Server=222.101.9.218;Database=village_admin;Uid=viluser;Pwd=viluser@csit;";
            string query = "SELECT * FROM VA_moneyreceived where vm_year = '" + year + "'" + "and vm_month='" + month + "'" + "and vm_room ='" + room + "'";
            string result = "";

            using (MySqlConnection conn = new MySqlConnection(strConn))
            {
                conn.Open();

                using (MySqlCommand command = new MySqlCommand(query, conn))
                {
                    //command.Parameters.AddWithValue("@val", val);

                    //ExecuteReader를 이용하여
                    //연결 모드로 데이타 가져오기
                    MySqlDataReader rdr = command.ExecuteReader();
                    rdr.Read();

                    if (rdr.HasRows == false)
                        return null;

                    result = rdr["vm_money"].ToString();


                    rdr.Close();
                }
            }

            return result;
        }

        private void Load_Click_1(object sender, EventArgs e)
        {
            dataMap.Clear();
            meals_DB();
            electricity_DB();
            common_DB();
            moneyreceived_DB();

          

            if (dataMap.Count() == 0)
            {
                MessageBox.Show("ERROR " + curYear + "년 " + curMonth + "월 데이터가 없습니다.");

                return;
            }

            dataGridView1.DataSource = null;
            
            DataTable table = new DataTable();

            // column을 추가합니다.
            table.Columns.Add("번호", typeof(string));
            table.Columns.Add("이름", typeof(string));
            table.Columns.Add("호실", typeof(string));

            table.Columns.Add("전기 일반 전월검침", typeof(string));
            table.Columns.Add("전기 일반 금월검침", typeof(string));
            table.Columns.Add("전기 일반 사용량", typeof(string));
            table.Columns.Add("전기 일반 요금", typeof(string));
            table.Columns.Add("전기 일반 단가(KW당)", typeof(string));
            table.Columns.Add("전기 심야 전월검침", typeof(string));
            table.Columns.Add("전기 심야 금월검침", typeof(string));
            table.Columns.Add("전기 심야 사용량", typeof(string));
            table.Columns.Add("전기 심야 요금", typeof(string));
            table.Columns.Add("전기 심야 단가(KW당)", typeof(string));

            table.Columns.Add("식비 식수", typeof(string));
            table.Columns.Add("식비 단가", typeof(string));
            table.Columns.Add("식비 금액", typeof(string));

            
            table.Columns.Add("공용수도광열비", typeof(string));
            table.Columns.Add("소모품", typeof(string));
            table.Columns.Add("의료비환급금", typeof(string));
            table.Columns.Add("전화사용료", typeof(string));
            table.Columns.Add("케이블이용료", typeof(string));
            table.Columns.Add("전월생활비", typeof(string));
            table.Columns.Add("생활비", typeof(string));

            table.Columns.Add("미수금/선납금", typeof(string));



            int num = 1;
            foreach (KeyValuePair< int, data> row in dataMap)
            {
                string normal_money = string.Format("{0:#,0}", row.Value.normal_money);
                string normal_price = string.Format("{0:#,0}", row.Value.normal_price);

                string night_money = string.Format("{0:#,0}", row.Value.night_money);
                string night_price = string.Format("{0:#,0}", row.Value.night_price);

                string sum = string.Format("{0:#,0}", row.Value.sum);
                string price = string.Format("{0:#,0}", row.Value.price);
                string total_price = string.Format("{0:#,0}", row.Value.total_price);


                string pipe_heatingcost = string.Format("{0:#,0}", row.Value.pipe_heatingcost);
                string expendables = string.Format("{0:#,0}", row.Value.expendables);
                string medical_refund = string.Format("{0:#,0}", row.Value.medical_refund);
                string phonebill = string.Format("{0:#,0}", row.Value.phonebill);
                string cablefee = string.Format("{0:#,0}", row.Value.cablefee);
                string preliving_expenses = string.Format("{0:#,0}", row.Value.preliving_expenses);
                string living_expenses = string.Format("{0:#,0}", row.Value.living_expenses);

                string 선납잔액 = string.Format("{0:#,0}", row.Value.선납잔액);


                table.Rows.Add(num, row.Value.name, row.Value.room, row.Value.normal_prvele, row.Value.normal_curele, row.Value.normal_usage, normal_money, normal_price,
                    row.Value.night_prvele, row.Value.night_curele, row.Value.night_usage, night_money, night_price, sum, price, total_price, pipe_heatingcost,
                    expendables, medical_refund, phonebill, cablefee, preliving_expenses ,living_expenses, 선납잔액);
                num++;
            }



            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue;
            dataGridView1.DataSource = table;
         
            for (int i = 0; i < 4; i++)
            {
                dataGridView1.Columns[i].Frozen = true;
            }

            int column = 0;
            foreach (DataGridViewColumn dc in dataGridView1.Columns)
            {
                if(column > 0)
                    dc.ReadOnly = true;

                dc.Visible = true;
                column++;
            }

            for (int i = 0; i < table.Columns.Count; i++)
            {
                if ( i < 2 )
                    dataGridView1.Columns[i].Width = 60;
                else if (i == 2)
                    dataGridView1.Columns[i].Width = 150;
                else if (i == 3)
                    dataGridView1.Columns[i].Width = 60;

                else if (i < 5)
                    dataGridView1.Columns[i].Width = 130;
                else 
                    dataGridView1.Columns[i].Width = 150;
                //dataGridView1.Columns[i].ReadOnly = true;
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            /*textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
            textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            textBox4.Text = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();
            */

            string str = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            if (e.ColumnIndex == 2)  // 3번째 칼럼이 선택되면....
            {
                MessageBox.Show((e.RowIndex + 1) + "  Row  " + (e.ColumnIndex + 1) + "  Column button clicked ");
            }
        }


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex < 0 || e.RowIndex < 0)
                return;

            string str = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue.ToString();
            return;
        }

        private void comboBox_year_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            if (cb.SelectedIndex > -1)
            {
                curYear = int.Parse(cb.SelectedItem.ToString());
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            dataMap.Clear();
            meals_DB();
            electricity_DB();
            common_DB();
            moneyreceived_DB();

            if (dataMap.Count > 0)
             pre_common_DB();

            if (dataMap.Count() == 0)
            {
                MessageBox.Show("ERROR " + curYear + "년 " + curMonth + "월 데이터가 없습니다.");

                return;
            }

            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.FileName = "빌리지 관리비 업로드_" + curYear + "_" + curMonth + ".xlsx"; //초기 파일명을 지정할 때 사용한다.
            saveFileDialog.Filter = "Excel|*.xlsx";
            saveFileDialog.Title = "Save an Excel File";

            Nullable<bool> result = saveFileDialog.ShowDialog();
            if (result == true)
            {
                Cursor.Current = Cursors.WaitCursor;
                file_export(saveFileDialog.FileName);
                Cursor.Current = Cursors.Default;
            }

            xlWorkBook.Close(0);
            xlApp.Quit();
        }

        
        private void DBUpload_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("선택하신 " + curYear.ToString() + "년 " + curMonth.ToString() + "월 데이터를 업로드 합니다", "업로드", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                dataList.Clear();

                Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();
                openFileDlg.DefaultExt = "xlsx";
                openFileDlg.Filter = "Excel Files(*.xls)| *.xlsx";

                Nullable<bool> result = openFileDlg.ShowDialog();
                if (result == true)
                {
                    foreach (string filename in openFileDlg.FileNames)
                    {
                        Excel.Application excelApp = null;
                        Excel.Workbook wb = null;
                        Excel.Worksheet ws = null;
                        try
                        {
                            excelApp = new Excel.Application();
                            wb = excelApp.Workbooks.Open(filename);
                            // path 대신 문자열도 가능합니다
                            // 예. Open(@"D:\test\test.xslx");
                            ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;
                            // 첫번째 Worksheet를 선택합니다.
                            Excel.Range rng = ws.UsedRange;   // '여기'
                                                              // 현재 Worksheet에서 사용된 셀 전체를 선택합니다.
                            filddata = rng.Value;

                            for (int r = 3; r <= filddata.GetLength(0); r++)
                            {
                                data datas = new data();

                                for (int c = 1; c <= filddata.GetLength(1); c++)
                                {
                                    if (filddata[r, c] == null)
                                    {
                                        continue;
                                    }

                                    object buffer = filddata[r, c];

                                    try
                                    {
                                        switch (c)
                                        {
                                            case 3:
                                                datas.room = int.Parse(buffer.ToString());
                                                break;
                                            case 17://공용수도관열비
                                                datas.pipe_heatingcost = int.Parse(buffer.ToString());
                                                break;
                                            case 18://소모품
                                                datas.expendables = int.Parse(buffer.ToString());
                                                break;
                                            case 19://의료비환급금
                                                datas.medical_refund = int.Parse(buffer.ToString());
                                                break;
                                            case 20://전화요금
                                                datas.phonebill = int.Parse(buffer.ToString());
                                                break;
                                            case 21://케이블이용료
                                                datas.cablefee = int.Parse(buffer.ToString());
                                                break;
                                            case 22://전월생활비
                                                datas.preliving_expenses = int.Parse(buffer.ToString());
                                                break;
                                            case 23://생활비
                                                datas.living_expenses = int.Parse(buffer.ToString());
                                                break;
                                        }
                                    }
                                    catch (Exception Ex)
                                    {
                                        MessageBox.Show("오류 테이블 데이터가 잘못되었습니다. 확인하세요~~");
                                        return;
                                    }
                                }


                                dataList.Add(datas);
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                        finally
                        {
                            delete_data();

                            foreach (data data in dataList)
                            {
                                insert_data(data);
                            }

                            wb.Close(null, null, null);                 // close your workbook
                            excelApp.Quit();                                   // exit excel application

                            MessageBox.Show("업로드가 완료되었습니다~~");
                        }
                    }
                }
            }
        }

        private void gvSheetList_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.ColumnIndex == 0 && e.RowIndex == -1)
            {
                e.PaintBackground(e.ClipBounds, false);

                Point pt = e.CellBounds.Location;  // where you want the bitmap in the cell

                int nChkBoxWidth = 15;
                int nChkBoxHeight = 15;
                int offsetx = (e.CellBounds.Width - nChkBoxWidth) / 2;
                int offsety = (e.CellBounds.Height - nChkBoxHeight) / 2;

                pt.X += offsetx;
                pt.Y += offsety;

                CheckBox cb = new CheckBox();
                cb.Size = new Size(nChkBoxWidth, nChkBoxHeight);
                cb.Location = pt;
                cb.CheckedChanged += new EventHandler(gvSheetListCheckBox_CheckedChanged);

                ((DataGridView)sender).Controls.Add(cb);

                e.Handled = true;
            }
        }

        private void gvSheetListCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow r in dataGridView1.Rows)
            {
                r.Cells["CK"].Value = ((CheckBox)sender).Checked;
            }
        }

        private void 미수금_선납금_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("선택하신 " + curYear.ToString() + "년 " + curMonth.ToString() + "월 미수금,선납금을 업로드 합니다", "업로드", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                미수금_선납금dataList.Clear();

                Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();
                openFileDlg.DefaultExt = "xlsx";
                openFileDlg.Filter = "Excel Files(*.xls)| *.xlsx";

                Nullable<bool> result = openFileDlg.ShowDialog();
                if (result == true)
                {
                    foreach (string filename in openFileDlg.FileNames)
                    {
                        Excel.Application excelApp = null;
                        Excel.Workbook wb = null;
                        Excel.Worksheet ws = null;
                        try
                        {
                            excelApp = new Excel.Application();
                            wb = excelApp.Workbooks.Open(filename);
                            // path 대신 문자열도 가능합니다
                            // 예. Open(@"D:\test\test.xslx");
                            ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;
                            // 첫번째 Worksheet를 선택합니다.
                            Excel.Range rng = ws.UsedRange;   // '여기'
                                                              // 현재 Worksheet에서 사용된 셀 전체를 선택합니다.
                            filddata = rng.Value;

                            for (int r = 2; r <= filddata.GetLength(0); r++)
                            {
                                미수금_선납금data datas = new 미수금_선납금data();

                                for (int c = 2; c <= filddata.GetLength(1); c++)
                                {
                                    if (filddata[r, c] == null)
                                    {
                                        continue;
                                    }

                                    object buffer = filddata[r, c];

                                    try
                                    {
                                        switch (c)
                                        {
                                            case 2:
                                                string[] words = buffer.ToString().Split('(');
                                                if (words.Length > 1 && words[0] == "")
                                                {
                                                    words = words[1].ToString().Split(')');
                                                    datas.room = int.Parse(words[0]);
                                                }
                                                break;
                                            case 7:
                                                if (datas.room > 0)
                                                {
                                                    datas.미수금_선납금 = int.Parse(buffer.ToString());
                                                }
                                                break;

                                        }
                                    }
                                    catch (Exception Ex)
                                    {
                                        MessageBox.Show("오류 테이블 데이터가 잘못되었습니다. 확인하세요~~");
                                        return;
                                    }
                                }

                                if (datas.room > 0)
                                    미수금_선납금dataList.Add(datas);
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                        finally
                        {
                            미수금_선납금delete_data();

                            foreach (미수금_선납금data data in 미수금_선납금dataList)
                            {
                                미수금_선납금insert_data(data);
                            }

                            wb.Close(null, null, null);                 // close your workbook
                            excelApp.Quit();                                   // exit excel application

                            MessageBox.Show("업로드가 완료되었습니다~~");
                        }
                    }
                }
            }
        }

        private void comboBox_day_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            if (cb.SelectedIndex > -1)
            {
                curDay = int.Parse(cb.SelectedItem.ToString());
            }
        }
    }
}
