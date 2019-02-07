// С учетом капиллярных сил
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace csharpApplication
{
    class Density
    {
        public const double _compresFactorWater0 = 0.00004;
        public const double _compresFactorOil0 = 0.0001;
        public const double _compresFactor = 0.00005;
        public const double _densityWater0 = 1.05;
        public const double _densityOil0 = 0.75;

        public const double _porosity = 0.12;


    }

    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public class Main
    {
        #region переменные
        private string now = DateTime.Now.ToString();
        private Double mmO;
        private Double MO = 0;
        private Double M_oldO = 0;
        private Double mmW;
        private Double MW = 0;
        private Double M_oldW = 0;
        private Double Q_w;
        private Double Q_o;
        private Double Q_w_d;
        private Double Q_o_d;
        private Double Qwater = 0;
        //private Double _porosity;//пористость
        private Double _viscosityOil;//вязкость нефть
        private Double _viscosityWater;//вязкость вода
        //private Double _diameter=3;//диаметр
        private Double _L;
        private Double t;//

        bool q = true;

        //private Double j=1;
        private Double du;
        private Double dt_max;
        private Double dt;
        private Double dt_s;
        private Double dt_pr;
        private Double dt_start = 100;//начальный шаг по времени
        private Double kt = 1.5;// коэффициент увеличения шага по времени
        private Double Tt = 3628800;//Общее время расчета//0;//00;//2592000;//5184000;//1296000;//3024000;//93312000;//46656000;//23328000;//15552000;//432000;//18 мес, 36
        private Double dtm;
        private const double _densityOilSt = 0.81;
        private const double _densityWaterSt = 1.3;
        const Int32 _n = 30;//15;//количество ячеек по пространственной координате
        private Double[] _pressureOilNew = new Double[_n + 1];
        private Double[] _pressureWaterNew = new Double[_n + 1];
        private Double[] _pressureOilOld = new Double[_n + 1];
        private Double[] dp = new Double[_n + 1];// nadbavka

        private bool t_f = true;

        private Double[] dx = new Double[_n + 1];
        //private Double[] dtm = new Double[_n + 1];
        private Double[] u = new Double[_n + 1];
        private Double[] y = new Double[_n + 1];
        private Double[] dy = new Double[_n + 1];
        private Double[] x = new Double[_n + 1];
        private Double[] dy_kv = new Double[_n + 1];

        private Double Rc;
        private Double rs = 10; //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        private Double h = 1000;
        private Double[] dr = new Double[_n + 1];
        private Double[] r = new Double[_n + 1];

        private Double[] _saturationWaterOld = new Double[_n + 1];
        private Double[] _saturationWaterNew = new Double[_n + 1];


        private Double _pressure0;//
        private Double _pressure1;//

        private Double _saturationWater0;

        private Double _permeability;//абсолютная проницаемость
        private Double _saturationOilOst;//остаточная нефтенасыщенность
        private Double _saturationWaterOst;//остаточная водонасыщенность
        private Double Q; // дебит (левая граница) 
        private Double Qsum;

        //переменные для прогонки
        private Double[] A = new Double[_n + 1];
        private Double[] AA = new Double[_n + 1];
        private Double[] C = new Double[_n + 1];
        private Double[] CC = new Double[_n + 1];
        private Double[] C2 = new Double[_n + 1];

        private Double[] B = new Double[_n + 1];
        private Double[] BB = new Double[_n + 1];
        private Double[] F = new Double[_n + 1];
        private Double[] F2 = new Double[_n + 1];
        private Double[] FF2 = new Double[_n + 1];
        private Double[] F3 = new Double[_n + 1];

        private Double[] alfa = new Double[_n + 1];
        private Double[] betta = new Double[_n + 1];
        //

        private Double[] K1_oil = new Double[_n + 1];
        private Double[] K1_water = new Double[_n + 1];
        private Double[] K2_oil = new Double[_n + 1];
        private Double[] K2_water = new Double[_n + 1];
        private Double[] denW_pr = new Double[_n + 1];
        private Double[] denO_pr = new Double[_n + 1];
        private Double[] denW_pr_ob = new Double[_n + 1];
        private Double[] denO_pr_ob = new Double[_n + 1];


        private Double[] A1 = new Double[_n + 1];
        private Double[] B1 = new Double[_n + 1];
        private Double[] C1 = new Double[_n + 1];
        private Double[] F1 = new Double[_n + 1];

        private Double[] uuu = new Double[_n + 1];
        private Double[] f_pl = new Double[_n + 1];
        private Double[] f_mn = new Double[_n + 1];
        private Double[] psi_pl = new Double[_n + 1];
        private Double[] psi_mn = new Double[_n + 1];
        private Double[] UU = new Double[_n + 1];
        private Double[] UUU = new Double[_n + 1];
        private Double[] Ff = new Double[_n + 1];
        private Double[] Psi = new Double[_n + 1];
        private Double[] Pk = new Double[_n + 1];
        private Double jj;
        private Double[] t_m = new Double[_n + 1];
        private Double[] f_pr = new Double[_n + 1];
        private Double[] kn = new Double[_n + 1];
        private Double[] kw = new Double[_n + 1];
        private Double[] Vn = new Double[_n + 1];
        private Double[] Vw = new Double[_n + 1];
        private Double[] PermOil = new Double[_n + 1];
        private Double[] PermWater = new Double[_n + 1];
        private Double[] PressCap = new Double[_n + 1];
        private Double[] PressCapNew = new Double[_n + 1];
        private Double[] row = new Double[_n + 1];
        private Double[] F_F = new double[_n + 1];
        private Double[] F_Psi = new double[_n + 1];
        private Double[] F_pr = new double[_n + 1];
        private Double[] Pcap_pr = new double[_n + 1];
        private Double[] Obvodn = new double[_n + 1];
        private Double[] Uw = new double[_n + 1];
        private Double[] Un = new double[_n + 1];
        private Double[] U_new = new double[_n + 1];
        private Double[] ObvodnNew = new double[_n + 1];
        private Double[] Velosity = new double[_n + 1];//i+1/2
        private Double[] VelisityOil = new double[_n + 1];//i+1/2
        private Double[] VelosityWater = new double[_n + 1];//i+1/2
        private Double[] Vnew = new double[_n + 1];
        private Double[] KKn = new double[_n + 1];
        private Double[] dPn = new double[_n + 1];
        private Double[] KKw = new double[_n + 1];
        private Double[] dPw = new double[_n + 1];
        private Double[] KKn_vel = new double[_n + 1];
        private Double[] KKw_vel = new double[_n + 1];
        private Double[] dPn_vel = new double[_n + 1];
        private Double[] dPw_vel = new double[_n + 1];
        private Double[] Skor_u_n = new double[_n + 1];
        private Double[] Skor_u_w = new double[_n + 1];
        private Double[] Skor_u = new double[_n + 1];
        private Double[] Vel_u_n = new double[_n + 1];
        private Double[] Vel_u_w = new double[_n + 1];
        private Double[] Vel_u = new double[_n + 1];
        private Double[] VelosityNew = new double[_n + 1];//i+1/2
        private Double[] VelisityOilNew = new double[_n + 1];//i+1/2
        private Double[] VelosityWaterNew = new double[_n + 1];//i+1/2
        private Int32 iter;

        private Double[] Sk0_Oil_1 = new double[2];
        private Double[] Sk0_Water_1 = new double[2];
        private Double[] Sk0_1 = new double[2];
        private Double[] Sk0_Oil_2 = new double[2];
        private Double[] Sk0_Water_2 = new double[2];
        private Double[] Sk0_2 = new double[2];
        private double ps;

        private Double[] m = new double[_n + 1];
        private Double[] m_old = new double[_n + 1];

        private Double[] RR = new double[_n + 1];

        private double dt_max1;


        private Double[] p_iter_old = new double[_n + 1];
        private Double[] p_iter_new = new double[_n + 1];

        public Double[] densityWater = new double[_n + 1];
        public Double[] densityOil = new double[_n + 1];
        
        public Double[] densityWaterOld = new double[_n + 1];
        public Double[] densityOilOld = new double[_n + 1];

        // headers для Excel для времени
        public Dictionary<int, string> hoursTime = new Dictionary<int, string>() {
            { 0, "нач" },
            { 1, "1 час" },
            { 3, "3 часа" },
            { 6, "6 часов" },
            { 12, "12 часов" },
            { 24, "1 сутки" },
            { 24*3, "3 сутки" },
            { 24*5, "5 суток" },
            { 24*8, "8 сутки" },
            { 24*9+23, "9+ суток" },
            { 24*10, "10 суток" },
            //{ 24*10 + 3, "10 суток 3 часа" },
            //{ 24*10 + 6, "10 суток 6 часа" },
            //{ 24*10 + 12, "10 суток 12 часа" },
            //{ 24*11, "11 суток" },
            { 24*12, "12 суток" },
            //{ 24*13, "13 суток" },
            //{ 24*14, "14 суток" },
            { 24*15, "15 суток" },
            { 24*18, "18 суток" },
            { 24*20, "20 суток" },
            { 24*22, "22 суток" },
            { 24*25, "25 суток" },
            { 24*28, "28 суток" },
            { 24*29 + 23, "29+ суток" },
            { 24*30, "30 суток" },
            { 24*30 + 3, "30 суток 3 часа" },
            { 24*30 + 6, "30 суток 6 часов" },
            { 24*30 + 12, "30 суток 12 часов" },
            { 24*31, "31 суток" },
            { 24*32, "32 суток" },
            { 24*33, "33 суток" },
            { 24*34, "34 суток" },
            { 24*35, "35 суток" },
            { 24*38, "38 суток" },
            { 24*40, "40 суток" },
            { 24*42, "42 суток" },
            { 24*45, "45 суток" },
            { 24*48, "48 суток" },
            { 24*50, "50 суток" },
            { 24*52, "52 суток" },
            { 24*55, "55 суток" },
            { 24*58, "58 суток" },
            //{ 24*59, "59+ суток" },
            { 24*60, "60 суток" },
            { 24*65, "65 суток" },
            { 24*70, "70 суток" },
            { 24*80, "80 суток" },
            { 24*90, "90 суток" },
            { 24*100, "100 суток" },
            { 24*120, "120 суток" },
            { 24*130, "130 суток" },
            { 24*150, "150 суток" },
            { 24*170, "170 суток" },
            { 24*200, "200 суток" },
            { 24*220, "220 суток" },
            { 24*250, "250 суток" },
            { 24*280, "280 суток" },
            { 24*310, "310 суток" },
            { 24*350, "350 суток" },
            { 24*380, "380 суток" },
            { 24*400, "400 суток" },
            { 24*419 + 23, "419+ суток" }

        //    { 24*120, "120 суток" },
        //    { 24*180, "180 суток" },
        //    { 24*210, "210 суток" },
        //    { 24*240, "240 суток" },
        //    { 24*270, "270 суток" },
        //    { 24*360, "360 суток" },
        //    { 24*450, "450 суток" },
        //    { 24*640, "640 суток" },
        //    { 24*730, "730 суток" },
        //    { 24*820, "820 суток" },
        //    { 24*910, "910 суток" },
        //    { 24*1000, "1000 суток" },
        //    { 24*1080, "1080 суток" }
        };

        #endregion

        public Main()
        {
        }

        //Работа с excel

        public List<string> string_fromArray_double(double[] array)
        {
            List<string> result = new List<string>();
            for (int i = 0; i < array.Length; i++)
                result.Add(array.ToString());
            return result;
        }

        public List<float> getTimeXls(string name) {
            FileInfo file = new FileInfo(name);
            if (file.Exists)
            {
                Console.WriteLine("Excel exist!");
            }
            else
            {
                Console.WriteLine(string.Format("Excel '{0}' does not exist! Shutdown", name));
                return new List<float>();
            }
            var inputTimeXls = new ExcelPackage(file);
            var timeSheet = inputTimeXls.Workbook.Worksheets.FirstOrDefault();
            if (timeSheet == null)
            {
                Console.WriteLine("Worksheet does not exist! Shutdown");
                return new List<float>();
            }
            int row = 1;
            bool error = false;
            var result = new List<float>();

            while(!error) {
                var cellValue = timeSheet.Cells[row, 1].Value;
                float res = 0;
                if (cellValue != null)
                {
                    bool isParsed = float.TryParse(cellValue.ToString(), out res);
                    if (isParsed) {
                        result.Add(res);
                    }
                    else {
                        error = true;
                        Console.WriteLine(string.Format("Row {0} is not parsed", row));
                    }
                }
                else {
                    error = true;
                    Console.WriteLine(string.Format("Row {0} is null", row));
                }
                row++;
            }
            Console.WriteLine("Excel read success");
            return result;
        }

        public void StartProcess()
        {
            Dictionary<String, double[]> dataToExcel = new Dictionary<string, double[]>();
            List<Dictionary<String, double[]>> datas = new List<Dictionary<String, double[]>>();

            Density dens = new Density();

            //обработка и проверка исходных данных
            String addresseeFilePath;
            CreateNewFile(out addresseeFilePath);

            // Если ничего не выбрано, ничего не делаем
            if (string.IsNullOrEmpty(addresseeFilePath))
                return;

            var streamReader = new StreamReader(addresseeFilePath);//создание объекта для чтения из файла
            String firstText = streamReader.ReadToEnd();//чтение всего файла
            String secondText = firstText.Replace(" ", String.Empty);
            streamReader.Close();//закрытие файла

            var valuesList = secondText.Replace("m=", "").Replace("muOil=", "").Replace("muWater=", "").Replace("L=", "").Replace("p0=", "").Replace("p1=", "").Replace("k=", "").Replace("Sno=", "").Replace("Swo=", "").Replace("Cwater=", "").Replace("Coil=", "").Replace("C=", "").Replace("roWater=", "").Replace("roOil=", "").Replace("Q=", "").Split(new char[] { ';', ' ' }).ToList();//

            var ValueArray = new List<Double>();

            foreach (var value in valuesList)
            {
                if (!String.IsNullOrWhiteSpace(value))
                {
                    Double valueDouble;
                    if (Double.TryParse(value, out valueDouble))
                    {
                        ValueArray.Add(valueDouble);
                    }
                }
            }

            #region присваиваем значения из файла переменным
            //_porosity = ValueArray[0];
            _viscosityOil = ValueArray[0];
            _viscosityWater = ValueArray[1];
            Rc = ValueArray[2];//L
            _pressure0 = ValueArray[3];
            _pressure1 = ValueArray[4];
            _permeability = ValueArray[5];
            _saturationOilOst = ValueArray[6];
            _saturationWaterOst = ValueArray[7];
            Q = ValueArray[8];
            Qsum = ValueArray[9];

            #endregion
            string _saturationWaterNew_filePath = get_filePath();

            _saturationWater0 = 0.422; // НУ! сейчас 30 суток!!!!!

            _saturationWaterOld[0] = 1 - _saturationOilOst;

            t = 0;
            #region коррдината x
            //x[0] = 1;
            ////du = (Math.Log((_L+1)/(_L+1)) - Math.Log(x[0] / (_L + 1))) / (_n );
            //u[0] = Math.Log(x[0] / (_L + 1));
            //u[_n] = Math.Log((_L + 1) / (_L + 1));
            //du = (u[_n] - u[0]) / _n;
            ////x[_n] = _L + 1;
            //dy[0] = (_L + 1) * Math.Exp(u[0] + du / 2) - (_L + 1) * Math.Exp(u[0]);
            //for (int i = 1; i <= _n; i++)
            //{
            //    u[i] = u[i - 1] + du;
            //    dy[i] = (_L + 1) * Math.Exp(u[i] + du / 2) - (_L + 1) * Math.Exp(u[i - 1] + du / 2);

            //    x[i] = (_L + 1) * Math.Exp(u[i]);
            //    dx[i] = x[i] - x[i - 1];
            //}
            //dx[0] = dx[1];
            ////dy[0] = dy[1];
            #endregion
            r[0] = rs;
            //du = (Math.Log((_L+1)/(_L+1)) - Math.Log(x[0] / (_L + 1))) / (_n );
            u[0] = Math.Log(r[0] / Rc);
            double u_n = 0;
            u_n = Math.Log(Rc / Rc);
            du = (u_n - u[0]) / (_n + 0.5);
            //x[_n] = _L + 1;
            dy[0] = Rc * Math.Exp(u[0] + du / 2) - Rc * Math.Exp(u[0]);
            dy_kv[0] = (Rc * Math.Exp(u[0] + du / 2) * Rc * Math.Exp(u[0] + du / 2) - Rc * Math.Exp(u[0]) * Rc * Math.Exp(u[0])) / 2;
            // dy_kv[0] = Rc * Math.Exp(u[0] + du / 2) * Rc * Math.Exp(u[0] + du / 2) - 2 * Rc * Math.Exp(u[0] + du / 2) * Rc * Math.Exp(u[0]) + Rc * Math.Exp(u[0]) * Rc * Math.Exp(u[0]);
            for (int i = 1; i <= _n; i++)
            {
                u[i] = u[i - 1] + du;
                dy[i] = Rc * Math.Exp(u[i] + du / 2) - Rc * Math.Exp(u[i - 1] + du / 2);
                dy_kv[i] = (Rc * Math.Exp(u[i] + du / 2) * Rc * Math.Exp(u[i] + du / 2) - Rc * Math.Exp(u[i - 1] + du / 2) * Rc * Math.Exp(u[i - 1] + du / 2)) / 2;
                // dy_kv[i] = Rc * Math.Exp(u[i] + du / 2) * Rc * Math.Exp(u[i] + du / 2) - 2 * Rc * Math.Exp(u[i] + du / 2) * Rc * Math.Exp(u[i - 1] + du / 2) + Rc * Math.Exp(u[i - 1] + du / 2) * Rc * Math.Exp(u[i - 1] + du / 2);

                r[i] = Rc * Math.Exp(u[i]);
                dr[i] = r[i] - r[i - 1];
            }
            dy[_n] = Rc * Math.Exp(u[_n] + du / 2) - Rc * Math.Exp(u[_n - 1] + du / 2);
            //dy[_n + 1] = dy[_n]; 
            //dx[0] = dx[1];

            //string text2 = "";
            //for (Int32 p = 0; p < dr.GetLength(0); p++)
            //{
            //    text2 = text2 + dr[p].ToString() + " " + " \r";
            //}
            //save_toFile("dr = ", _saturationWaterNew_filePath);
            //save_toFile(text2, _saturationWaterNew_filePath);
            dataToExcel.Add("dr", (double[])dr.Clone());
            //string text22 = "";
            //for (Int32 p = 0; p < r.GetLength(0); p++)
            //{
            //    text22 = text22 + r[p].ToString() + " " + " \r";
            //}
            //save_toFile("r = ", _saturationWaterNew_filePath);
            //save_toFile(text22, _saturationWaterNew_filePath);
            dataToExcel.Add("u", (double[])dy.Clone());

            //string text222 = "";
            double[] logR = new double[dr.GetLength(0)];
            for (Int32 p = 0; p < dr.GetLength(0); p++)
            {
                logR[p] = Math.Log(r[p]);
                //text222 = text222 + Math.Log(r[p]).ToString() + " " + " \r";
            }
            //save_toFile("lnr = ", _saturationWaterNew_filePath);
            //save_toFile(text222, _saturationWaterNew_filePath);
            //  dataToExcel.Add("lnr", (double[])logR.Clone());
            // bool bol = true;
            var timeArray = getTimeXls("input_time.xlsx");
            var timesExist = timeArray.Count != 0;
            int timeIndex = 0;
            while ((t < Tt && !timesExist) ||
                   (timesExist && timeIndex <= timeArray.Count))
            {
                var dataExcel = new Dictionary<String, double[]>();
                dataExcel.Add("dt_enter", new double[] { dt });
                // Density dens = new Density();
                //// t=0
                #region t=0
                //Расчет
                if (t == 0)
                {
                    jj = 0;
                    if (timesExist) {
                        dt = timeArray[0];
                    }
                    else {
						dt = dt_start;
                    }
					_pressureOilNew[0] = _pressure1;//!!!_pressure1
					//усл без ГРП
					_saturationWaterOld[0] = _saturationWater0;// 1 - _saturationOilOst;0
					_saturationWaterNew[0] = _saturationWater0;// 1 - _saturationOilOst;


                    for (Int32 i = 1; i <= _n; i++)
                    {
                        _pressureOilNew[i] = _pressure0;
                        _saturationWaterNew[i] = _saturationWater0;
                        _saturationWaterOld[i] = _saturationWaterNew[i];
                    }

                    //#region Для ГРП
                    //for (Int32 i = 1; i <= _n; i++)
                    //{
                    //    _pressureOilNew[i] = _pressure0;
                    //}
                    //for (Int32 i = 0; i <= 5; i++)
                    //{
                    //    _saturationWaterNew[i] = 1 - _saturationOilOst;
                    //    _saturationWaterOld[i] = _saturationWaterNew[i];
                    //}
                    //for (Int32 i = 6; i <= _n; i++)
                    //{
                    //    _saturationWaterNew[i] = _saturationWater0;
                    //    _saturationWaterOld[i] = _saturationWaterNew[i];
                    //}
                    //#endregion
                    //
                    for (Int32 i = 0; i <= _n; i++)
                    {
                        PressCap[i] = 0.937 * (Math.Exp(-25 * _saturationWaterOld[i]) - Math.Exp(-25 * (1 - _saturationOilOst))) / (Math.Exp(-25 * _saturationWaterOst) - Math.Exp(-25 * (1 - _saturationOilOst)));

                        _pressureWaterNew[i] = _pressureOilNew[i] - PressCap[i];
                    }

                    for (Int32 i = 0; i <= _n; i++)
                    {
                        //m[i] = _porosity * (1 + _compresFactor * (_pressureOilNew[i] - _pressure0));
                        m[i] = Density._porosity * (1 + Density._compresFactor * (_pressureOilNew[i] - _pressure0));

                        densityWater[i] = Density._densityWater0 * (1 + Density._compresFactorWater0 * (_pressureWaterNew[i] - _pressure0));// dens.GetDensityWater(_pressureWaterNew[i], _pressure1); 

                        //row [i] = Density._densityWater0 * (1 + Density._compresFactorWater0 * (_pressureWaterNew[i] - _pressure0));// через ф-цию надо
                        densityOil[i] = Density._densityOil0 * (1 + Density._compresFactorOil0 * (_pressureOilNew[i] - _pressure0));// dens.GetDensityOil(_pressureOilNew[i], _pressure0); //Density._densityOil0 * (1 + Density._compresFactorOil0 * (_pressureOilNew[i] - _pressure0));// через ф-цию надо
                        // PressCap[i] = 0.937 * (Math.Exp(-25 * _saturationWaterOld[i]) - Math.Exp(-25 * (1 - _saturationOilOst))) / (Math.Exp(-25 * _saturationWaterOst) - Math.Exp(-25 * (1 - _saturationOilOst)));
                        PermWater[i] = 0.05 * Math.Pow((_saturationWaterOld[i] - _saturationWaterOst), 1.2) / Math.Pow((1 - _saturationOilOst - _saturationWaterOst), 1.2);
                        PermOil[i] = Math.Pow((1 - _saturationWaterOld[i] - _saturationOilOst), 2.2) / Math.Pow((1 - _saturationWaterOst - _saturationOilOst), 2.2);
                        F_F[i] = PermWater[i] / (PermWater[i] + PermOil[i] * _viscosityWater / _viscosityOil);
                        F_Psi[i] = PermOil[i] * PermWater[i] / (PermWater[i] + _viscosityWater * PermOil[i] / _viscosityOil);
                        F_pr[i] = 0.4114489718 * Math.Pow((_saturationWaterOld[i] - _saturationWaterOst), 0.2) / (0.3428741432 * Math.Pow((_saturationWaterOld[i] - _saturationWaterOst), 1.2) + 34.11683017 * (_viscosityWater / _viscosityOil) * (Math.Pow((1 - _saturationWaterOld[i] - _saturationOilOst), 2.2))) - 0.3428741432 * Math.Pow((_saturationWaterOld[i] - _saturationWaterOst), 1.2) * (0.4114489718 * Math.Pow((_saturationWaterOld[i] - _saturationWaterOst), 0.2) - 75.05702637 * (_viscosityWater / _viscosityOil) * (Math.Pow((1 - _saturationWaterOld[i] - _saturationOilOst), 1.2))) / (Math.Pow((0.3428741432 * Math.Pow((_saturationWaterOld[i] - _saturationWaterOst), 1.2) + 34.11683017 * (_viscosityWater / _viscosityOil) * (Math.Pow((1 - _saturationWaterOld[i] - _saturationOilOst), 2.2))), 2));
                        Pcap_pr[i] = -8.779958120 * 100000 * Math.Exp(-25 * _saturationWaterOld[i]);

                    }


                    for (Int32 l = 0; l < _n; l++)
                    {
                        VelisityOil[l] = -1 * _permeability * (PermOil[l + 1] / _viscosityOil) * (_pressureOilNew[l + 1] - _pressureOilNew[l]) / dr[l + 1];
                        VelosityWater[l] = -1 * _permeability * (PermWater[l + 1] / _viscosityWater) * (_pressureOilNew[l + 1] - _pressureOilNew[l]) / dr[l + 1] - (-1 * _permeability * PermWater[l + 1]) / _viscosityWater * (PressCap[l + 1] - PressCap[l]) / dr[l + 1];
                        Velosity[l] = VelisityOil[l] + VelosityWater[l];

                    }


                    for (Int32 i = 0; i <= _n; i++)
                    {
                        ObvodnNew[i] = VelosityWater[i] / Velosity[i]; //Vw[i] / UUU[i];
                    }

                    double U = Velosity[0];




                    if (!timesExist) // Если нет массива времен из excel
                    {
                        dt_max1 = 1 / 2 * (Density._porosity / (Math.Abs(U))) * (1 / F_pr[1]) * RR[0];

                        if (F_Psi[0] > F_Psi[1])
                        {
                            ps = F_Psi[0];
                        }
                        else
                        {
                            ps = F_Psi[1];
                        }

                        // dt_max = 0.6728 * Density._porosity * _viscosityOil * dr[1] * dr[1] * dy_kv[0] / (Math.Abs(Pcap_pr[0]) * r[0] * _permeability * ps); //новое 0.6728
                        RR[0] = Rc * Rc * (Math.Exp(2 * (u[0] + du / 2)) - Math.Exp(2 * (u[0])));

                        dt_max = 0.3364 * Density._porosity * _viscosityOil * du * du * RR[0] / (Math.Abs(Pcap_pr[0]) * _permeability * ps); //новое 0.6728


                        for (Int32 j = 1; j < _n; j++)
                        {
                            if (F_Psi[j] > F_Psi[j + 1])
                            {
                                ps = F_Psi[j];
                            }
                            else
                            {
                                ps = F_Psi[j + 1];
                            }
                            RR[j] = Rc * Rc * (Math.Exp(2 * (u[j] + du / 2)) - Math.Exp(2 * (u[j] - du / 2)));

                            //dtm = 0.6728 * Density._porosity * _viscosityOil * dr[j] * dr[j] * dy_kv[j] / (Math.Abs(Pcap_pr[j]) * r[j] * _permeability * ps); //0.6728
                            dtm = 0.3364 * Density._porosity * _viscosityOil * du * du * RR[j] / (Math.Abs(Pcap_pr[j]) * _permeability * ps); //новое 0.6728


                            if (dt_max > dtm)
                            {
                                dt_max = dtm;
                            }
                        }
                        if (dt_max1 < dt_max)
                        {
                            dt_max = dt_max1;
                        }

                        //сравниваем dt с dtmax
                        if (dt_max < dt)
                        {
                            dt = dt_max;//!!!!*10

                        }
                    }
                    //if (dt > 1000)
                    //{
                    //    dt = dt/2;
                    //}

                }
                #endregion
                //// t!=0

                if (t != 0)
                {
                    if (_saturationWaterOld[0] < (1 - _saturationOilOst))
                    {
                        //!!!!!
                        if (t >= 864000)// 2592000)//1 мес = 2592000//8 мес=20736000//864000
                        {

                            Qsum = 0;// 113.6;//28.95; //

                        }

                        for (Int32 m = 0; m <= _n; m++)
                        {
                            PermWater[m] = 0.05 * Math.Pow((_saturationWaterOld[m] - _saturationWaterOst), 1.2) / Math.Pow((1 - _saturationOilOst - _saturationWaterOst), 1.2);
                            PermOil[m] = Math.Pow((1 - _saturationWaterOld[m] - _saturationOilOst), 2.2) / Math.Pow((1 - _saturationWaterOst - _saturationOilOst), 2.2);
                            PressCap[m] = 0.937 * (Math.Exp(-25 * _saturationWaterNew[m]) - Math.Exp(-25 * (1 - _saturationOilOst))) / (Math.Exp(-25 * _saturationWaterOst) - Math.Exp(-25 * (1 - _saturationOilOst)));
                            F_F[m] = PermWater[m] / (PermWater[m] + PermOil[m] * _viscosityWater / _viscosityOil);
                            F_Psi[m] = PermOil[m] * PermWater[m] / (PermWater[m] + _viscosityWater * PermOil[m] / _viscosityOil);
                            F_pr[m] = 0.4114489718 * Math.Pow((_saturationWaterOld[m] - _saturationWaterOst), 0.2) / (0.3428741432 * Math.Pow((_saturationWaterOld[m] - _saturationWaterOst), 1.2) + 34.11683017 * (_viscosityWater / _viscosityOil) * (Math.Pow((1 - _saturationWaterOld[m] - _saturationOilOst), 2.2))) - 0.3428741432 * Math.Pow((_saturationWaterOld[m] - _saturationWaterOst), 1.2) * (0.4114489718 * Math.Pow((_saturationWaterOld[m] - _saturationWaterOst), 0.2) - 75.05702637 * (_viscosityWater / _viscosityOil) * (Math.Pow((1 - _saturationWaterOld[m] - _saturationOilOst), 1.2))) / (Math.Pow((0.3428741432 * Math.Pow((_saturationWaterOld[m] - _saturationWaterOst), 1.2) + 34.11683017 * (_viscosityWater / _viscosityOil) * (Math.Pow((1 - _saturationWaterOld[m] - _saturationOilOst), 2.2))), 2));
                            Pcap_pr[m] = -8.779958120 * 100000 * Math.Exp(-25 * _saturationWaterOld[m]);

                        }

                        iter = 0;
                        bool dt_p = true;
                        // boolean it=
                        while (iter < 20 && dt_p == true)//12)
                        {
                            // /iter=0
                            #region iter=0
                            if (iter == 0)
                            {
                                for (Int32 i = 0; i < _n + 1; i++)
                                {
                                    p_iter_old[i] = _pressureOilNew[i];
                                    _pressureOilOld[i] = _pressureOilNew[i];
                                    p_iter_new[i] = p_iter_old[i];
                                }
                                iter++;
                            }
                            #endregion
                            // /iter>0
                            #region iter>0
                            if (iter != 0)
                            {
                                if (t >= 864000)//2592000)//  1 мес = 2592000//8 мес=20736000//864000
                                {

                                    Qsum = 0;// 113.6;//28.95; //

                                }

                                //if (Qsum == 0 && q == true)
                                //{
                                //    dt = 1;
                                //    q = false;
                                //}
                                for (Int32 i = 0; i <= _n; i++)/////////////////////////
                                {
                                    m_old[i] = Density._porosity * (1 + Density._compresFactor * (_pressureOilOld[i] - _pressure0));
                                    densityWaterOld[i] = Density._densityWater0 * (1 + Density._compresFactorWater0 * (_pressureOilOld[i] - PressCap[i] - _pressure0));
                                    densityOilOld[i] = Density._densityOil0 * (1 + Density._compresFactorOil0 * (_pressureOilOld[i] - _pressure0));//заменить методами

                                    m[i] = Density._porosity * (1 + Density._compresFactor * (p_iter_old[i] - _pressure0));
                                    densityWater[i] = Density._densityWater0 * (1 + Density._compresFactorWater0 * (p_iter_old[i] - PressCap[i] - _pressure0));//1
                                    densityOil[i] = Density._densityOil0 * (1 + Density._compresFactorOil0 * (p_iter_old[i] - _pressure0));//заменить методами //1

                                    K1_oil[i] = _permeability * PermOil[i] / _viscosityOil;
                                    K1_water[i] = _permeability * PermWater[i] / _viscosityWater;
                                    denW_pr[i] = Density._compresFactorWater0 * Density._densityWater0; //0
                                    denO_pr[i] = Density._compresFactorOil0 * Density._densityOil0; //0
                                    denW_pr_ob[i] = -1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[i] * densityWater[i]);
                                    denO_pr_ob[i] = -1 * Density._compresFactorOil0 * Density._densityOil0 / (densityOil[i] * densityOil[i]);



                                }
                                //коэф прогонки

                                #region прогонка

                                A[0] = 0;
                                AA[0] = 0;

                                double b0 = _permeability * PermOil[1] / _viscosityOil * 1 / du * (densityOil[1] / densityOil[0] + 1);
                                double b1 = _permeability * PermOil[1] / _viscosityOil * p_iter_old[1] / du * (1 / densityOil[0] * Density._compresFactorOil0 * Density._densityOil0);
                                double b2 = _permeability * PermOil[1] / _viscosityOil * p_iter_old[0] / du * (1 / densityOil[0] * Density._densityOil0 * Density._compresFactorOil0);
                                double b3 = _permeability * PermWater[1] / _viscosityWater * 1 / du * (densityWater[1] / densityWater[0] + 1);
                                double b4 = _permeability * PermWater[1] / _viscosityWater * p_iter_old[1] / du * (1 / densityWater[0] * Density._compresFactorWater0 * Density._densityWater0);
                                double b5 = _permeability * PermWater[1] / _viscosityWater * PressCap[1] / du * (1 / densityWater[0] * Density._densityWater0 * Density._compresFactorWater0);
                                double b6 = _permeability * PermWater[1] / _viscosityWater * p_iter_old[0] / du * (1 / densityWater[0] * Density._compresFactorWater0 * Density._densityWater0);
                                double b7 = _permeability * PermWater[1] / _viscosityWater * PressCap[0] / du * 0.5 * (1 / densityWater[0] * Density._densityWater0 * Density._compresFactorWater0);

                                B[0] = -1 * dt * Math.PI * h * (b0 + b1 - b2) - dt * Math.PI * h * (b3 + b4 - b5 - b6 + b7);// +qq1 - qq2; // +q3;//+q3

                                BB[0] = -1 * dt * Math.PI * h * (b0 + b1 - b2) - dt * Math.PI * h * (b3 + b4 - b5 - b6 + b7);

                                RR[0] = Rc * Rc * (Math.Exp(2 * (u[0] + du / 2)) - Math.Exp(2 * (u[0])));
                                double c0 = RR[0] * (1 - _saturationWaterOld[0]) * Density._porosity * Density._compresFactor;
                                double c1 = RR[0] * (1 - _saturationWaterOld[0]) * densityOilOld[0] * m_old[0] * (-1 * Density._compresFactorOil0 * Density._densityOil0 / (densityOil[0] * densityOil[0]));
                                double c2 = (_permeability * PermOil[1] / _viscosityOil * p_iter_old[1] / du * densityOil[1] * (-1 * Density._densityOil0 * Density._compresFactorOil0 / (densityOil[0] * densityOil[0])));
                                double c3 = _permeability * PermOil[1] / _viscosityOil * 1 / du * (densityOil[1] / densityOil[0] + 1);
                                double c4 = _permeability * PermOil[1] / _viscosityOil * p_iter_old[0] / du * densityOil[1] * (-1 * Density._compresFactorOil0 * Density._densityOil0 / (densityOil[0] * densityOil[0]));
                                double c6 = RR[0] * _saturationWaterOld[0] * densityWaterOld[0] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[0] * densityWater[0])) * m_old[0];
                                double c5 = RR[0] * _saturationWaterOld[0] * Density._porosity * Density._compresFactor;
                                double c7 = (_permeability * PermWater[1] / _viscosityWater * p_iter_old[1] / du * densityWater[1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[0] * densityWater[0])));
                                double c8 = (_permeability * PermWater[1] / _viscosityWater * PressCap[1] / du * densityWater[1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[0] * densityWater[0])));
                                double c9 = (_permeability * PermWater[1] / _viscosityWater * 1 / du * (1 + densityWater[1] / densityWater[0]));
                                double c10 = (_permeability * PermWater[1] / _viscosityWater * p_iter_old[0] / du * densityWater[1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[0] * densityWater[0])));
                                double c11 = (_permeability * PermWater[1] / _viscosityWater * PressCap[0] / du * densityWater[1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[0] * densityWater[0])));
                                //double c12 = du * _densityOilSt * Q * (-1 * Density._compresFactorOil0 * Density._densityOil0 / (densityOil[0] * densityOil[0])) / (2 * Math.PI * h);
                                double c13 = (-1 * Density._compresFactorOil0 * Density._densityOil0 / (densityOil[0] * densityOil[0]));

                                double c19 = dt * 2 * Math.PI * h * _permeability * PermWater[1] / _viscosityWater * 1 / du;

                                C[0] = c0 * Math.PI * h - c1 * Math.PI * h - (dt * Math.PI * h * c2 - dt * Math.PI * h * c3 - dt * Math.PI * h * c4) + Math.PI * h * c5 - Math.PI * h * c6 - (dt * Math.PI * h * c7 - dt * Math.PI * h * c8 - dt * Math.PI * h * c9 - dt * Math.PI * h * c10 + dt * Math.PI * h * c11);// +qq3 - qq4 + qq5;// +c16 - c19; //c17 - c18 - c19;//+ c16 + dt * _densityWaterSt / densityWater[0] * q4 + dt * _densityWaterSt * q4_0;//+c17; //+ c15;//+ c12

                                CC[0] = c0 * Math.PI * h - c1 * Math.PI * h - (dt * Math.PI * h * c2 - dt * Math.PI * h * c3 - dt * Math.PI * h * c4) + Math.PI * h * c5 - Math.PI * h * c6 - (dt * Math.PI * h * c7 - dt * Math.PI * h * c8 - dt * Math.PI * h * c9 - dt * Math.PI * h * c10 + dt * Math.PI * h * c11);// +c16; 


                                //F2[1] = -1 * ((1 - _saturationWaterOld[1]) * (m[1] / dt - (densityOilOld[1] / densityOil[1]) * m_old[1] / dt) - 1 / dy[1] * (_permeability * PermOil[2] / _viscosityOil * (p_iter_old[2] - p_iter_old[1]) / dx[2] * 0.5 * (densityOil[2] / densityOil[1] + 1) - _permeability * PermOil[1] / _viscosityOil * (p_iter_old[1] - p_iter_old[0]) / dx[1] * 0.5 * (1 + densityOil[0] / densityOil[1])) + _saturationWaterOld[1] * (m_old[1] / dt - densityWaterOld[1] / densityWater[1] * m_old[1] / dt) - 1 / dy[1] * (_permeability * PermWater[2] / _viscosityWater * (p_iter_old[2] - PressCap[2] - p_iter_old[1] + PressCap[1]) / dx[2] * 0.5 * (densityWater[2] / densityWater[1] + 1) - _permeability * PermWater[1] / _viscosityWater * (p_iter_old[1] - PressCap[1] - p_iter_old[0] + PressCap[0]) / dx[1] * 0.5 * (1 + densityWater[0] / densityWater[1])));// -_pressure1 * A[1];
                                double f0 = Math.PI * h * RR[0] * (1 - _saturationWaterOld[0]) * (m[0] - (densityOilOld[0] / densityOil[0]) * m_old[0]);
                                double f1 = dt * Math.PI * h * (_permeability * PermOil[1] / _viscosityOil * (p_iter_old[1] - p_iter_old[0]) / du * (densityOil[1] / densityOil[0] + 1));
                                double f2 = Math.PI * h * RR[0] * _saturationWaterOld[0] * (m[0] - densityWaterOld[0] / densityWater[0] * m_old[0]);
                                double f3 = dt * Math.PI * h * (_permeability * PermWater[1] / _viscosityWater * (p_iter_old[1] - PressCap[1] - p_iter_old[0] + PressCap[0]) / du * (densityWater[1] / densityWater[0] + 1));
                                double f4 = dt * _densityOilSt / densityOil[0] * Q;
                                double f5 = dt * 2 * Math.PI * h * (_permeability * PermWater[1] / _viscosityWater) * (p_iter_old[1] - PressCap[1] - p_iter_old[0] + PressCap[0]) / du;
                                double qq6 = dt * Qsum;// -dt * densityWater[1] * 2 * Math.PI * h * (_permeability * PermWater[0] / _viscosityWater) * (p_iter_old[1] - PressCap[1] - p_iter_old[0] + PressCap[0]) / du;

                                F2[0] = -1 * (f0 - f1 + f2 - f3 + qq6); //+ qq7);

                                FF2[0] = -1 * (f0 - f1 + f2 - f3 + f4);

                                for (Int32 l = 1; l < _n; l++)
                                {
                                    A1[l] = -1 / dy[l] * (-K1_oil[l] * _pressureOilNew[l] / du * 0.5 * 1 / densityOil[l] * denO_pr[l] + K1_oil[l] * 1 / du * 0.5 + K1_oil[l] * 1 / du * 0.5 * 1 / densityOil[l] * (densityOil[l - 1] + _pressureOilNew[l - 1] * denO_pr[l - 1])) - 1 / dy[l] * (-K1_water[l] * _pressureOilNew[l] / du * 0.5 * 1 / densityWater[l] * denW_pr[l - 1] + K1_water[l] * PressCap[l] / du * 0.5 * 1 / densityWater[l] * denW_pr[l - 1] + K1_water[l] * 1 / du * 0.5 * (1 + densityWater[l - 1] / densityWater[l]) + K1_water[l] * _pressureOilNew[l - 1] / du * 0.5 * 1 / densityWater[l] * denW_pr[l - 1] - K1_water[l] * PressCap[l - 1] / du * 0.5 - K1_water[l] * PressCap[l - 1] / du * 1 / densityWater[l] * denW_pr[l - 1]);

                                    double al0 = dt * Math.PI * h * _permeability * PermOil[l] / _viscosityOil * (p_iter_old[l] / du * Density._compresFactorOil0 * Density._densityWater0 * 1 / densityOil[l]);
                                    double al1 = dt * Math.PI * h * _permeability * PermOil[l] / _viscosityOil * 1 / du * (densityOil[l - 1] / densityOil[l] + 1);
                                    double al2 = dt * Math.PI * h * _permeability * PermOil[l] / _viscosityOil * p_iter_old[l - 1] / du * Density._compresFactorOil0 * Density._densityOil0 * 1 / densityOil[l];
                                    double al3 = dt * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * (p_iter_old[l] - PressCap[l]) / du * Density._compresFactorWater0 * Density._densityWater0 * 1 / densityWater[l];
                                    double al4 = dt * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * 1 / du * (1 + densityWater[l - 1] / densityWater[l]);
                                    double al5 = dt * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * (p_iter_old[l - 1] - PressCap[l - 1]) / du * Density._compresFactorWater0 * Density._densityWater0 * 1 / densityWater[l];
                                    //- PressCap[l] / du * Density._compresFactorWater0 * Density._densityWater0 * 1 / densityWater[l] - _pressureOilNew[l - 1] / du * Density._compresFactorWater0 * Density._densityWater0 * 1 / densityWater[l] - (1 + densityWater[l - 1] / densityWater[l]) * 1 / du + PressCap[l - 1] / du * Density._densityWater0 * Density._compresFactorWater0 * 1 / densityWater[l]);
                                    A[l] = al0 - al1 - al2 + al3 - al4 - al5;
                                    AA[l] = A[l];
                                    //
                                    C2[l] = r[l] * r[l] * (1 - _saturationWaterOld[l]) * 1 / dt * Density._porosity * Density._compresFactor - r[l] * r[l] * (1 - _saturationWaterOld[l]) * densityOil[l] * m_old[l] / dt * (-1 * Density._compresFactorOil0 * Density._densityOil0 / (densityOil[l] * densityOil[l])) - 1 / dy[l] * (_permeability * PermOil[l + 1] / _viscosityOil * _pressureOilNew[l + 1] / du * 0.5 * densityOil[l + 1] * (-1 * Density._densityOil0 * Density._compresFactorOil0 / (densityOil[l] * densityOil[l])) - _permeability * PermOil[l + 1] / _viscosityOil * 1 / du * 0.5 * (densityOil[l + 1] / densityOil[l] + 1) - _permeability * PermOil[l + 1] / _viscosityOil * _pressureOilNew[l] / du * 0.5 * densityOil[l + 1] * (-1 * Density._compresFactorOil0 * Density._densityOil0 / (densityOil[l] * densityOil[l])) - _permeability * PermOil[l] / _viscosityOil * 1 / du * 0.5 * (1 + densityOil[l - 1] / densityOil[l]) - _permeability * PermOil[l] / _viscosityOil * _pressureOilNew[l] / du * 0.5 * densityOil[l - 1] * (-1 * Density._densityOil0 * Density._compresFactorOil0 / (densityOil[l] * densityOil[l])) + _permeability * PermOil[l] / _viscosityOil * _pressureOilNew[l - 1] / du * 0.5 * densityOil[l - 1] * (Density._compresFactorOil0 * Density._densityOil0 / (densityOil[l] * densityOil[l]))) + r[l] * r[l] * _saturationWaterOld[l] * (1 / dt) * Density._porosity * Density._compresFactor - r[l] * r[l] * _saturationWaterOld[l] * densityWaterOld[l] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l])) * m_old[l] / dt - 1 / dy[l] * (_permeability * PermWater[l + 1] / _viscosityWater * _pressureOilNew[l + 1] / du * 0.5 * densityWater[l + 1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[l] * densityWater[l])) - _permeability * PermWater[l + 1] / _viscosityWater * PressCap[l + 1] / du * 0.5 * densityWater[l + 1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l])) - _permeability * PermWater[l + 1] / _viscosityWater * 1 / du * 0.5 * (1 + densityWater[l + 1] / densityWater[l]) - _permeability * PermWater[l + 1] / _viscosityWater * _pressureOilNew[l] / du * 0.5 * densityWater[l + 1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[l] * densityWater[l])) + _permeability * PermWater[l + 1] / _viscosityWater * PressCap[l] / du * 0.5 * densityWater[l + 1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l])) - _permeability * PermWater[l] / _viscosityWater * 1 / du * 0.5 * (1 + densityWater[l - 1] / densityWater[l]) - _permeability * PermWater[l] / _viscosityWater * _pressureOilNew[l] / du * 0.5 * densityWater[l - 1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[l] * densityWater[l])) + _permeability * PermWater[l] / _viscosityWater * PressCap[l] / du * 0.5 * densityWater[l - 1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l])) + _permeability * PermWater[l] / _viscosityWater * _pressureOilNew[l - 1] / du * 0.5 * densityWater[l - 1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[l] * densityWater[l])) - _permeability * PermWater[l] / _viscosityWater * PressCap[l - 1] / du * 0.5 * densityWater[l - 1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l])));
                                    /// dy[1] * (_permeability * PermWater[2] * (densityWater[2] + densityWater[1]) / (2 * densityWater[1] * _viscosityWater * dx[2]) + _permeability * PermWater[1] * (densityWater[1] + dens._densityWater0) / (2 * densityWater[1] * _viscosityWater * dx[1]) + _permeability * PermOil[2] * (densityOil[2] + densityOil[1]) / (2 * densityOil[1] * _viscosityOil * dx[2]) + _permeability * PermOil[1] * (densityOil[1] + dens._densityOil0) / (2 * densityOil[1] * _viscosityOil * dx[1]));
                                    RR[l] = Rc * Rc * (Math.Exp(2 * (u[l] + du / 2)) - Math.Exp(2 * (u[l] - du / 2)));

                                    double cl0 = Math.PI * h * (1 - _saturationWaterOld[l]) * Density._porosity * Density._compresFactor * RR[l];
                                    double cl1 = Math.PI * h * (1 - _saturationWaterOld[l]) * densityOilOld[l] * m_old[l] * (-1 * Density._compresFactorOil0 * Density._densityOil0 / (densityOil[l] * densityOil[l])) * RR[l];
                                    double cl2 = dt * Math.PI * h * _permeability * PermOil[l + 1] / _viscosityOil * p_iter_old[l + 1] / du * densityOil[l + 1] * (-1 * Density._densityOil0 * Density._compresFactorOil0 / (densityOil[l] * densityOil[l]));
                                    double cl3 = dt * Math.PI * h * _permeability * PermOil[l + 1] / _viscosityOil * 1 / du * (densityOil[l + 1] / densityOil[l] + 1);
                                    double cl4 = dt * Math.PI * h * _permeability * PermOil[l + 1] / _viscosityOil * p_iter_old[l] / du * densityOil[l + 1] * (-1 * Density._compresFactorOil0 * Density._densityOil0 / (densityOil[l] * densityOil[l]));
                                    double cl5 = dt * Math.PI * h * _permeability * PermOil[l] / _viscosityOil * 1 / du * (1 + densityOil[l - 1] / densityOil[l]);
                                    double cl6 = dt * Math.PI * h * _permeability * PermOil[l] / _viscosityOil * p_iter_old[l] / du * densityOil[l - 1] * (-1 * Density._densityOil0 * Density._compresFactorOil0 / (densityOil[l] * densityOil[l]));
                                    double cl7 = dt * Math.PI * h * _permeability * PermOil[l] / _viscosityOil * p_iter_old[l - 1] / du * densityOil[l - 1] * (-1 * Density._densityOil0 * Density._compresFactorOil0 / (densityOil[l] * densityOil[l]));// (Density._compresFactorOil0 * Density._densityOil0 / (densityOil[l] * densityOil[l]));
                                    double cl8 = Math.PI * h * _saturationWaterOld[l] * Density._porosity * Density._compresFactor * RR[l];
                                    double cl9 = Math.PI * h * _saturationWaterOld[l] * densityWaterOld[l] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l])) * m_old[l] * RR[l];
                                    double cl10 = dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * p_iter_old[l + 1] / du * densityWater[l + 1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[l] * densityWater[l]));
                                    double cl11 = dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * PressCap[l + 1] / du * densityWater[l + 1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l]));
                                    double cl12 = dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * 1 / du * (1 + densityWater[l + 1] / densityWater[l]);
                                    double cl13 = dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * p_iter_old[l] / du * densityWater[l + 1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[l] * densityWater[l]));
                                    double cl14 = dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * PressCap[l] / du * densityWater[l + 1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l]));
                                    double cl15 = dt * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * 1 / du * (1 + densityWater[l - 1] / densityWater[l]);
                                    double cl16 = dt * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * p_iter_old[l] / du * densityWater[l - 1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[l] * densityWater[l]));
                                    double cl17 = dt * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * PressCap[l] / du * densityWater[l - 1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l]));
                                    double cl18 = dt * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * p_iter_old[l - 1] / du * densityWater[l - 1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[l] * densityWater[l]));
                                    double cl19 = dt * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * PressCap[l - 1] / du * densityWater[l - 1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l]));

                                    C[l] = cl0 - cl1 - cl2 + cl3 + cl4 + cl5 + cl6 - cl7 + cl8 - cl9 - cl10 + cl11 + cl12 + cl13 - cl14 + cl15 + cl16 - cl17 - cl18 + cl19;// -cl20;
                                    CC[l] = C[l];

                                    B[l] = -1 * (dt * Math.PI * h * _permeability * PermOil[l + 1] / _viscosityOil * 1 / du * (densityOil[l + 1] / densityOil[l] + 1) + dt * Math.PI * h * _permeability * PermOil[l + 1] / _viscosityOil * _pressureOilNew[l + 1] / du * (1 / densityOil[l] * Density._compresFactorOil0 * Density._densityOil0) - dt * Math.PI * h * _permeability * PermOil[l + 1] / _viscosityOil * _pressureOilNew[l] / du * 1 / densityOil[l] * Density._densityOil0 * Density._compresFactorOil0) - (dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * 1 / du * (densityWater[l + 1] / densityWater[l] + 1) + dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * _pressureOilNew[l + 1] / du * (1 / densityWater[l] * Density._compresFactorWater0 * Density._densityWater0) - dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * PressCap[l + 1] / du * (1 / densityWater[l] * Density._densityWater0 * Density._compresFactorWater0) - dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * _pressureOilNew[l] / du * (1 / densityWater[l] * Density._compresFactorWater0 * Density._densityWater0) + dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * PressCap[l] / du * (1 / densityWater[l] * Density._densityWater0 * Density._compresFactorWater0));

                                    BB[l] = B[l];


                                    // F[l] = -1 * ((1 - _saturationWaterOld[l]) * (m[l] / dt - (dens.GetDensityOil(_pressureOilOld[l], _pressure1) / dens.GetDensityOil(p_iter_old[l], _pressure1)) * m_old[l] / dt) - 1 / dy[l] * (_permeability * PermWater[l + 1] / _viscosityOil * (p_iter_old[l + 1] - p_iter_old[l]) / dx[l + 1] * 0.5 * (dens.GetDensityOil(p_iter_old[l + 1], _pressure1) / dens.GetDensityOil(p_iter_old[l], _pressure1) + 1) - _permeability * PermOil[l] / _viscosityOil * (p_iter_old[l] - p_iter_old[l - 1]) / dx[l] * 0.5 * (1 + dens.GetDensityOil(p_iter_old[l - 1], _pressure1) / dens.GetDensityOil(p_iter_old[l], _pressure1))) + r[l] * r[l] * _saturationWaterOld[l] * (m_old[l] / dt - dens.GetDensityWater(_pressureOilOld[l], _pressure1) / dens.GetDensityWater(p_iter_old[l], _pressure1) * m_old[l] / dt) - 1 / dy[l] * (_permeability * PermWater[l + 1] / _viscosityWater * (p_iter_old[l + 1] - PressCap[l + 1] - p_iter_old[l] + PressCap[l]) / dx[l + 1] * 0.5 * (dens.GetDensityWater(p_iter_old[l + 1], _pressure1) / dens.GetDensityWater(p_iter_old[l], _pressure1) + 1) - _permeability * PermWater[l] / _viscosityWater * (p_iter_old[l] - PressCap[l] - p_iter_old[l - 1] + PressCap[l - 1]) / dx[l] * 0.5 * (1 + dens.GetDensityWater(p_iter_old[l - 1], _pressure1) / dens.GetDensityWater(p_iter_old[l], _pressure1))));

                                    F3[l] = -1 * r[l] * r[l] * ((1 - _saturationWaterOld[l]) * (m[l] / dt - (densityOilOld[l] / densityOil[l]) * m_old[l] / dt) - 1 / dy[l] * (_permeability * PermOil[l + 1] / _viscosityOil * (p_iter_old[l + 1] - p_iter_old[l]) / du * 0.5 * (densityOil[l + 1] / densityOil[l] + 1) - _permeability * PermOil[l] / _viscosityOil * (p_iter_old[l] - p_iter_old[l - 1]) / du * 0.5 * (1 + densityOil[l - 1] / densityOil[l])) + r[l] * r[l] * _saturationWaterOld[l] * (m_old[l] / dt - densityWaterOld[l] / densityWater[l] * m_old[l] / dt) - 1 / dy[l] * (_permeability * PermWater[l + 1] / _viscosityWater * (p_iter_old[l + 1] - PressCap[l + 1] - p_iter_old[l] + PressCap[l]) / du * 0.5 * (densityWater[l + 1] / densityWater[l] + 1) - _permeability * PermWater[l] / _viscosityWater * (p_iter_old[l] - PressCap[l] - p_iter_old[l - 1] + PressCap[l - 1]) / du * 0.5 * (1 + densityWater[l - 1] / densityWater[l])));// - _pressure1 * A[1];

                                    double fl0 = Math.PI * h * (1 - _saturationWaterOld[l]) * (m[l] - (densityOilOld[l] / densityOil[l]) * m_old[l]) * RR[l];
                                    double fl1 = dt * Math.PI * h * _permeability * PermOil[l + 1] / _viscosityOil * (p_iter_old[l + 1] - p_iter_old[l]) / du * (densityOil[l + 1] / densityOil[l] + 1);
                                    double fl2 = dt * Math.PI * h * _permeability * PermOil[l] / _viscosityOil * (p_iter_old[l] - p_iter_old[l - 1]) / du * (1 + densityOil[l - 1] / densityOil[l]);
                                    double fl3 = Math.PI * h * _saturationWaterOld[l] * (m[l] - densityWaterOld[l] / densityWater[l] * m_old[l]) * RR[l];
                                    double fl4 = dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * (p_iter_old[l + 1] - PressCap[l + 1] - p_iter_old[l] + PressCap[l]) / du * (densityWater[l + 1] / densityWater[l] + 1);
                                    double fl5 = dt * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * (p_iter_old[l] - PressCap[l] - p_iter_old[l - 1] + PressCap[l - 1]) / du * (1 + densityWater[l - 1] / densityWater[l]);
                                    F2[l] = -1 * (fl0 - (fl1 - fl2) + fl3 - (fl4 - fl5));// - dt * _densityOilSt / densityOil[l] * Q);
                                    FF2[l] = F2[l];
                                }


                                //правая граница
                                double an1 = dt * Math.PI * h * (1 / densityOil[_n]) * (_permeability * PermOil[_n] / _viscosityOil) * (p_iter_old[_n] / du) * denO_pr[_n - 1];
                                double an2 = dt * Math.PI * h * (_permeability * PermOil[_n] / _viscosityOil) * 1 / du * (densityOil[_n - 1] / densityOil[_n] + 1);

                                double an4 = dt * Math.PI * h * (1 / densityOil[_n]) * (_permeability * PermOil[_n] / _viscosityOil) * (p_iter_old[_n - 1] / du) * denO_pr[_n - 1];
                                double an5 = dt * Math.PI * h * (1 / densityWater[_n]) * (_permeability * PermWater[_n] / _viscosityWater) * (p_iter_old[_n] / du) * denW_pr[_n - 1];
                                double an6 = dt * Math.PI * h * (1 / densityWater[_n]) * (_permeability * PermWater[_n] / _viscosityWater) * (PressCap[_n] / du) * denW_pr[_n - 1];
                                double an7 = dt * Math.PI * h * (1 / densityWater[_n]) * (_permeability * PermWater[_n] / _viscosityWater) * (1 / du) * (densityWater[_n - 1] / densityWater[_n] + 1);

                                double an9 = dt * Math.PI * h * (1 / densityWater[_n]) * (_permeability * PermWater[_n] / _viscosityWater) * (p_iter_old[_n - 1] / du) * denW_pr[_n - 1];
                                double an10 = dt * Math.PI * h * (1 / densityWater[_n]) * (_permeability * PermWater[_n] / _viscosityWater) * (PressCap[_n - 1] / du) * denW_pr[_n - 1];
                                A[_n] = an1 - an2 - an4 + an5 - an6 - an7 - an9 + an10;
                                AA[_n] = A[_n];

                                B[_n] = 0;
                                BB[_n] = B[_n];

                                RR[_n] = Rc * Rc * (Math.Exp(2 * (u[_n])) - Math.Exp(2 * (u[_n] - du / 2)));
                                double cn0 = (_permeability * PermOil[_n] / _viscosityOil);
                                double cn00 = (_permeability * PermWater[_n] / _viscosityWater);
                                double cn1 = Math.PI * h * (1 - _saturationWaterOld[_n]) * Density._porosity * Density._compresFactor * RR[_n];// - 
                                double cn2 = Math.PI * h * denO_pr_ob[_n] * (1 - _saturationWaterOld[_n]) * (densityOilOld[_n] * m_old[_n]) * RR[_n];
                                double cn3 = dt * Math.PI * h * denO_pr_ob[_n] * (_permeability * PermOil[_n] / _viscosityOil) * (p_iter_old[_n] / du) * (densityOil[_n - 1]);
                                double cn4 = dt * Math.PI * h * (1 / densityOil[_n]) * (_permeability * PermOil[_n] / _viscosityOil) * (1 / du) * (densityOil[_n - 1]);
                                double cn5 = dt * Math.PI * h * (_permeability * PermOil[_n] / _viscosityOil) * (1 / du);
                                double cn6 = dt * Math.PI * h * denO_pr_ob[_n] * cn0 * (p_iter_old[_n - 1] / du) * (densityOil[_n - 1]);

                                double cn8 = Math.PI * h * _saturationWaterOld[_n] * Density._porosity * Density._compresFactor * RR[_n];
                                double cn9 = Math.PI * h * denW_pr_ob[_n] * _saturationWaterOld[_n] * (densityWaterOld[_n] * m_old[_n]) * RR[_n];
                                double cn10 = dt * Math.PI * h * denW_pr_ob[_n] * cn00 * (p_iter_old[_n] / du) * (densityWater[_n - 1]);
                                double cn11 = dt * Math.PI * h * (1 / densityWater[_n]) * cn00 * (1 / du) * (densityWater[_n - 1]);
                                double cn12 = dt * Math.PI * h * cn00 * (1 / du);
                                double cn13 = dt * Math.PI * h * cn00 * (denW_pr_ob[_n]) * (PressCap[_n] / du) * (densityWater[_n - 1]);
                                double cn14 = dt * Math.PI * h * cn00 * denW_pr_ob[_n] * (p_iter_old[_n - 1] / du) * (densityWater[_n - 1]);
                                double cn15 = dt * Math.PI * h * denW_pr_ob[_n] * cn00 * (PressCap[_n - 1] / du) * (densityWater[_n - 1]);
                                C[_n] = cn1 - cn2 + cn3 + cn4 + cn5 - cn6 + cn8 - cn9 + cn10 + cn11 + cn12 - cn13 - cn14 + cn15;// -dt * _densityOilSt * Q * denO_pr_ob[_n];
                                CC[_n] = C[_n];

                                double fn1 = Math.PI * h * (1 - _saturationWaterOld[_n]) * (m[_n] - (densityOilOld[_n] / densityOil[_n]) * m_old[_n]) * RR[_n];
                                double fn2 = dt * Math.PI * h * (_permeability * PermOil[_n] / _viscosityOil) * (p_iter_old[_n] - p_iter_old[_n - 1]) / du * (densityOil[_n - 1] / densityOil[_n] + 1);
                                double fn3 = Math.PI * h * _saturationWaterOld[_n] * (m[_n] - densityWaterOld[_n] / densityWater[_n] * m_old[_n]) * RR[_n];
                                double fn4 = dt * Math.PI * h * (_permeability * PermWater[_n] / _viscosityWater) * (p_iter_old[_n] - PressCap[_n] - p_iter_old[_n - 1] + PressCap[_n - 1]) / du * (densityWater[_n - 1] / densityWater[_n] + 1);
                                F2[_n] = -1 * (fn1 + fn2 + fn3 + fn4);// - dt * _densityOilSt / densityOil[_n] * Q);// проверить знаки !!!+fn2+fn4
                                FF2[_n] = F2[_n];

                                #endregion прогонка
                                // alfa, betta
                                #region alfa, betta
                                alfa[1] = (-1 * B[0]) / C[0];
                                betta[1] = F2[0] / C[0];

                                for (Int32 k = 1; k < _n; k++)
                                {
                                    alfa[k + 1] = (-1 * B[k]) / (A[k] * alfa[k] + C[k]);
                                    betta[k + 1] = (F2[k] - A[k] * betta[k]) / (A[k] * alfa[k] + C[k]);

                                }
                                #endregion

                                dp[_n] = (F2[_n] - A[_n] * betta[_n]) / (C[_n] + A[_n] * alfa[_n]);


                                for (Int32 g = _n - 1; g >= 0; g--)
                                {
                                    dp[g] = alfa[g + 1] * dp[g + 1] + betta[g + 1];

                                }

                                for (Int32 k = 0; k <= _n; k++)
                                {
                                    p_iter_new[k] = p_iter_old[k] + dp[k];

                                }

                                #region // проверка dt
                                //dt_max1 = (Density._porosity / Math.Abs(Velosity[0])) * (1 / F_pr[1]) * dy[0];
                                //dt_max1 = (Density._porosity / (r[0] * Math.Abs(Velosity[0]))) * (1 / F_pr[1]) * dy_kv[0];
                                dt_max1 = 1 / 2 * (Density._porosity / (Math.Abs(Velosity[0]))) * (1 / F_pr[1]) * RR[0];



                                for (Int32 j = 1; j < _n; j++)// определяем dtmax как min по j
                                {

                                    // dtm = (Density._porosity / Math.Abs(Velosity[j])) * (1 / F_pr[j + 1]) * dy[j];
                                    dtm = (Density._porosity / (r[j] * Math.Abs(Velosity[j]))) * (1 / F_pr[j + 1]) * dy_kv[j];//!!!!!


                                    if (dt_max1 > dtm)
                                    {
                                        dt_max1 = dtm;
                                    }

                                }
                                if (F_Psi[0] > F_Psi[1])
                                {
                                    ps = F_Psi[0];
                                }
                                else
                                {
                                    ps = F_Psi[1];
                                }

                                RR[0] = Rc * Rc * (Math.Exp(2 * (u[0] + du / 2)) - Math.Exp(2 * (u[0])));

                                dt_max = 0.3364 * Density._porosity * _viscosityOil * du * du * RR[0] / (Math.Abs(Pcap_pr[0]) * _permeability * ps); //новое 0.6728


                                for (Int32 j = 1; j < _n; j++)
                                {
                                    if (F_Psi[j] > F_Psi[j + 1])
                                    {
                                        ps = F_Psi[j];
                                    }
                                    else
                                    {
                                        ps = F_Psi[j + 1];
                                    }
                                    RR[j] = Rc * Rc * (Math.Exp(2 * (u[j] + du / 2)) - Math.Exp(2 * (u[j] - du / 2)));

                                    //dtm = 0.6728 * Density._porosity * _viscosityOil * dr[j] * dr[j] * dy_kv[j] / (Math.Abs(Pcap_pr[j]) * r[j] * _permeability * ps); //0.6728
                                    dtm = 0.3364 * Density._porosity * _viscosityOil * du * du * RR[j] / (Math.Abs(Pcap_pr[0]) * _permeability * ps); //новое 0.6728

                                    if (dt_max > dtm)
                                    {
                                        dt_max = dtm;
                                    }
                                }
                                if (dt_max1 < dt_max)
                                {
                                    dt_max = dt_max1;
                                }
                                double dt2 = dt;
                                //сравниваем dt с dtmax
                                if (dt_max < dt2)
                                {
                                    dt2 = dt_max;//* 10
                                                 //bol = true;
                                }
                                if (dt2 < dt)
                                {
                                    dt = dt2;
                                    iter = 0;
                                    dt_p = true;
                                }
                                else
                                {
                                    dt_p = false;
                                }
                                //if (dt > 1000)
                                //{
                                //    dt = dt/2;
                                //}

                                //if (t > 864000 && dt > 1000)//2592000 && dt > 1000)//
                                //{
                                //    dt = dt / 10;
                                //}

                                #endregion

                                // norma
                                #region
                                double SumKvOld = 0;
                                double SumKvNew = 0;
                                for (Int32 k = 0; k <= _n; k++)
                                {
                                    SumKvOld = SumKvOld + p_iter_old[k] * p_iter_old[k];
                                    SumKvNew = SumKvNew + p_iter_new[k] * p_iter_new[k];
                                }

                                double norm;
                                norm = Math.Abs(Math.Sqrt(SumKvOld) - Math.Sqrt(SumKvNew));
                                for (Int32 k = 0; k <= _n; k++)
                                {
                                    p_iter_old[k] = p_iter_new[k];
                                }

                                if (norm < 0.001)
                                {
                                    for (Int32 k = 0; k <= _n; k++)
                                    {
                                        _pressureOilNew[k] = p_iter_new[k];
                                    }

                                    break;
                                }
                                else
                                {
                                    iter++;
                                }
                                #endregion
                            }
                            #endregion
                            //iter++;


                            for (Int32 k = 0; k <= _n; k++)
                            {
                                p_iter_old[k] = p_iter_new[k];

                            }
                        }
                        dt_pr = dt;
                        dataExcel.Add("iter1", new double[] { iter });
                        //dataExcel.Add("dt_pres", new double[] { dt });

                        for (Int32 k = 0; k <= _n; k++)
                        {
                            _pressureOilOld[k] = _pressureOilNew[k];

                        }
                        for (Int32 k = 0; k <= _n; k++)
                        {
                            _pressureOilNew[k] = p_iter_new[k];

                        }
                        for (Int32 k = 0; k < _n + 1; k++)
                        {
                            _pressureWaterNew[k] = _pressureOilNew[k] - PressCap[k];
                        }


                        for (Int32 i = 0; i <= _n; i++)
                        {
                            m_old[i] = Density._porosity * (1 + Density._compresFactor * (_pressureOilOld[i] - _pressure1));
                            densityWaterOld[i] = Density._densityWater0 * (1 + Density._compresFactorWater0 * (_pressureOilOld[i] - PressCap[i] - _pressure1));
                            densityOilOld[i] = Density._densityOil0 * (1 + Density._compresFactorOil0 * (_pressureOilOld[i] - _pressure0));//заменить методами

                            m[i] = Density._porosity * (1 + Density._compresFactor * (_pressureOilNew[i] - _pressure1));
                            densityWater[i] = Density._densityWater0 * (1 + Density._compresFactorWater0 * (_pressureOilNew[i] - PressCap[i] - _pressure1));
                            densityOil[i] = Density._densityOil0 * (1 + Density._compresFactorOil0 * (_pressureOilNew[i] - _pressure1));//заменить методами
                        }

                        for (Int32 l = 1; l < _n; l++)
                        {
                            VelisityOil[l] = -1 * _permeability * (PermOil[l + 1] / _viscosityOil) * (_pressureOilNew[l + 1] - _pressureOilNew[l]) / dr[l + 1];
                            VelosityWater[l] = -1 * _permeability * (PermWater[l + 1] / _viscosityWater) * (_pressureOilNew[l + 1] - _pressureOilNew[l]) / dr[l + 1] - (-1 * _permeability * PermWater[l + 1]) / _viscosityWater * (PressCap[l + 1] - PressCap[l]) / dr[l + 1];
                            Velosity[l] = VelisityOil[l] + VelosityWater[l];
                        }

                        for (Int32 i = 0; i <= _n; i++)
                        {
                            ObvodnNew[i] = VelosityWater[i] / Velosity[i];// Vw[i] / UUU[i];
                        }


                        #region ШАГ
                        if (!timesExist)  // Если нет Excel
                        {
                            //dt_max1 = (Density._porosity / Math.Abs(Velosity[0])) * (1 / F_pr[1]) * dy[0];
                            //dt_max1 = (Density._porosity / (r[0] * Math.Abs(Velosity[0]))) * (1 / F_pr[1]) * dy_kv[0];
                            dt_max1 = 1 / 2 * (Density._porosity / (Math.Abs(Velosity[0]))) * (1 / F_pr[1]) * RR[0];


                            for (Int32 j = 1; j < _n; j++)// определяем dtmax как min по j
                            {

                                // dtm = (Density._porosity / Math.Abs(Velosity[j])) * (1 / F_pr[j + 1]) * dy[j];
                                dtm = (Density._porosity / (r[j] * Math.Abs(Velosity[j]))) * (1 / F_pr[j + 1]) * dy_kv[j];


                                if (dt_max1 > dtm)
                                {
                                    dt_max1 = dtm;
                                }

                            }
                            if (F_Psi[0] > F_Psi[1])
                            {
                                ps = F_Psi[0];
                            }
                            else
                            {
                                ps = F_Psi[1];
                            }

                            RR[0] = Rc * Rc * (Math.Exp(2 * (u[0] + du / 2)) - Math.Exp(2 * (u[0])));

                            dt_max = 0.3364 * Density._porosity * _viscosityOil * du * du * RR[0] / (Math.Abs(Pcap_pr[0]) * _permeability * ps); //новое 0.6728


                            for (Int32 j = 1; j < _n; j++)
                            {
                                if (F_Psi[j] > F_Psi[j + 1])
                                {
                                    ps = F_Psi[j];
                                }
                                else
                                {
                                    ps = F_Psi[j + 1];
                                }
                                RR[j] = Rc * Rc * (Math.Exp(2 * (u[j] + du / 2)) - Math.Exp(2 * (u[j] - du / 2)));

                                //dtm = 0.6728 * Density._porosity * _viscosityOil * dr[j] * dr[j] * dy_kv[j] / (Math.Abs(Pcap_pr[j]) * r[j] * _permeability * ps); //0.6728
                                dtm = 0.3364 * Density._porosity * _viscosityOil * du * du * RR[j] / (Math.Abs(Pcap_pr[0]) * _permeability * ps); //новое 0.6728

                                if (dt_max > dtm)
                                {
                                    dt_max = dtm;
                                }
                            }
                            if (dt_max1 < dt_max)
                            {
                                dt_max = dt_max1;
                            }

                            //сравниваем dt с dtmax
                            if (dt_max < dt)
                            {
                                dt = dt_max;//* 10
                                            //bol = true;
                            }
                        }
                        //if (dt > 1000)
                        //{
                        //    dt = dt/2;
                        //}

                        //if (t > 864000 && dt > 1000)//2592000 && dt > 1000)//
                        //{
                        //    dt = dt / 10;
                        //}
                        #endregion
                        dt_s = dt;
                        //dataExcel.Add("dt_sat", new double[] { dt });

                        #region new насыщенность


                        if (_saturationWaterOld[0] < (1 - _saturationOilOst) || Qsum == 0)
                        {


                            // double g1 = _saturationWaterOld[0] * (-1 * (m_old[0] * densityWaterOld[0]) / (m[0] * densityWater[0]) + 1);
                            double g1 = _saturationWaterOld[0] * (1 - (m_old[0] * densityWaterOld[0]) / (m[0] * densityWater[0]));

                            //double g2 = dt / (RR[0] * m[0] * densityWater[0]) * (_permeability / _viscosityWater) * PermWater[1] * (_pressureWaterNew[1] - _pressureWaterNew[0]) / du * (densityWater[1] + densityWater[0]);
                            double g2 = dt / (RR[0] * m[0]) * (_permeability / _viscosityWater) * PermWater[1] * (_pressureWaterNew[1] - _pressureWaterNew[0]) / du * (1 + densityWater[1] / densityWater[0]);

                            _saturationWaterNew[0] = _saturationWaterOld[0] - g1 + g2;

                            if (_saturationWaterNew[0] > (1 - _saturationOilOst))
                            {
                                _saturationWaterNew[0] = 1 - _saturationOilOst;
                                Q_w = 2 * Math.PI * h * (_permeability * PermWater[1] / _viscosityWater) * (_pressureOilNew[1] - PressCap[1] - _pressureOilNew[0] + PressCap[0]) / du;

                                Q_w_d = 2 * Math.PI * h * (_permeability * PermWater[_n] / _viscosityWater) * (_pressureOilNew[_n] - PressCap[_n] - _pressureOilNew[0] + PressCap[0]) / (Math.Log(Rc / rs));


                                Q_o_d = 2 * Math.PI * h * (_permeability * PermOil[_n] / _viscosityOil) * (_pressureOilNew[_n] - _pressureOilNew[0]) / (Math.Log(Rc / rs));

                            }
                        }
                        if (_saturationWaterOld[0] >= (1 - _saturationOilOst) && Qsum != 0)
                        {
                            _saturationWaterNew[0] = 1 - _saturationOilOst;// _saturationWaterOld[0];

                        }

                        Double[] gg1 = new Double[_n + 1];
                        Double[] gg2 = new Double[_n + 1];
                        Double[] gg3 = new Double[_n + 1];
                        Double[] gg4 = new Double[_n + 1];

                        for (Int32 i = 1; i < _n; i++)
                        {
                            //  gg1[i] = _saturationWaterOld[i] * (-1 * (m_old[i] * densityWaterOld[i]) / (m[i] * densityWater[i]) + 1);
                            gg1[i] = _saturationWaterOld[i] * (1 - (m_old[i] * densityWaterOld[i]) / (m[i] * densityWater[i]));

                            gg2[i] = dt / (RR[i] * m[i] * densityWater[i]);
                            gg3[i] = (_permeability / _viscosityWater) * PermWater[i + 1] * (_pressureWaterNew[i + 1] - _pressureWaterNew[i]) / du * (densityWater[i + 1] + densityWater[i]);
                            gg4[i] = (_permeability / _viscosityWater) * PermWater[i] * (_pressureWaterNew[i] - _pressureWaterNew[i - 1]) / du * (densityWater[i - 1] + densityWater[i]);
                            _saturationWaterNew[i] = _saturationWaterOld[i] - gg1[i] + gg2[i] * (gg3[i] - gg4[i]); //_saturationWaterOld[i] * (1 - (m_old[i] * densityWaterOld[i]) / (m[i] * densityWater[i])) + dt / (RR[i] * m[i]) * ((_permeability / _viscosityOil) * PermWater[i + 1] * (_pressureWaterNew[i + 1] - _pressureWaterNew[i]) / du * (densityWater[i + 1] / densityWater[i] + 1) - (_permeability / _viscosityWater) * PermWater[i] * (_pressureWaterNew[i] - _pressureWaterNew[i - 1]) / du * (1 + densityWater[i - 1] / densityWater[i]));


                        }
                        _saturationWaterNew[_n] = _saturationWaterOld[_n] - _saturationWaterOld[_n] * (-1 * (m_old[_n] * densityWaterOld[_n]) / (m[_n] * densityWater[_n]) + 1) + dt / (RR[_n] * m[_n] * densityWater[_n]) * (-1 * (_permeability / _viscosityWater) * PermWater[_n] * (_pressureWaterNew[_n] - _pressureWaterNew[_n - 1]) / du * (densityWater[_n - 1] + densityWater[_n]));


                        #endregion
                        //
                        for (Int32 l = 0; l <= _n; l++)
                        {
                            PermWater[l] = 0.05 * Math.Pow((_saturationWaterNew[l] - _saturationWaterOst), 1.2) / Math.Pow((1 - _saturationOilOst - _saturationWaterOst), 1.2);
                            PermOil[l] = Math.Pow((1 - _saturationWaterNew[l] - _saturationOilOst), 2.2) / Math.Pow((1 - _saturationWaterOst - _saturationOilOst), 2.2);

                        }


                        for (Int32 i = 0; i <= _n; i++)
                        {
                            MO = MO + (densityOil[i] * m[i] * Math.PI * h * Rc * Rc * (Math.Exp(2 * (u[i] + du / 2)) - Math.Exp(2 * (u[i] - du / 2))));
                            M_oldO = M_oldO + (densityOilOld[i] * m_old[i] * Math.PI * h * Rc * Rc * (Math.Exp(2 * (u[i] + du / 2)) - Math.Exp(2 * (u[i] - du / 2))));
                            MW = MW + (densityWater[i] * m[i] * Math.PI * h * Rc * Rc * (Math.Exp(2 * (u[i] + du / 2)) - Math.Exp(2 * (u[i] - du / 2))));
                            M_oldW = M_oldW + (densityWaterOld[i] * m_old[i] * Math.PI * h * Rc * Rc * (Math.Exp(2 * (u[i] + du / 2)) - Math.Exp(2 * (u[i] - du / 2))));

                        }
                        mmO = (MO - M_oldO - Q * dt) / M_oldO;
                        mmW = (MW - M_oldW - (2 * Math.PI * h * (_permeability * PermWater[1] / _viscosityWater) * (_pressureOilNew[1] - PressCap[1] - _pressureOilNew[0] + PressCap[0]) / du) * dt) / M_oldW;

                        Q_w = 0; // 2 * Math.PI * h * (_permeability * PermWater[1] / _viscosityWater) * (_pressureOilNew[1] - PressCap[1] - _pressureOilNew[0] + PressCap[0]) / du;
                        Q_o = 2 * Math.PI * h * (_permeability * PermOil[1] / _viscosityOil) * (_pressureOilNew[1] - _pressureOilNew[0]) / du;


                        Q_w_d = 0; // 2 * Math.PI * h * (_permeability * PermWater[_n] / _viscosityWater) * (_pressureOilNew[_n] - PressCap[_n] - _pressureOilNew[0] + PressCap[0]) / (Math.Log(Rc / rs));
                        Q_o_d = 2 * Math.PI * h * (_permeability * PermOil[_n] / _viscosityOil) * (_pressureOilNew[_n] - _pressureOilNew[0]) / (Math.Log(Rc / rs));

                    }
                    // else!!!!!!!!!!!!!!!!!!!!!!!!
                    else
                    {
                        //месяц!!

                        if (t >= 864000)//2592000)// 1 мес = 2592000//8 мес=20736000//864000
                        {

                            Qsum = 0;// 113.6;//28.95; //

                        }

                        //if (Qsum == 0 && q == true)
                        //{
                        //    dt = 1;
                        //    q = false;
                        //}

                        for (Int32 m = 0; m <= _n; m++)
                        {
                            PermWater[m] = 0.05 * Math.Pow((_saturationWaterOld[m] - _saturationWaterOst), 1.2) / Math.Pow((1 - _saturationOilOst - _saturationWaterOst), 1.2);
                            PermOil[m] = Math.Pow((1 - _saturationWaterOld[m] - _saturationOilOst), 2.2) / Math.Pow((1 - _saturationWaterOst - _saturationOilOst), 2.2);
                            PressCap[m] = 0.937 * (Math.Exp(-25 * _saturationWaterNew[m]) - Math.Exp(-25 * (1 - _saturationOilOst))) / (Math.Exp(-25 * _saturationWaterOst) - Math.Exp(-25 * (1 - _saturationOilOst)));
                            F_F[m] = PermWater[m] / (PermWater[m] + PermOil[m] * _viscosityWater / _viscosityOil);
                            F_Psi[m] = PermOil[m] * PermWater[m] / (PermWater[m] + _viscosityWater * PermOil[m] / _viscosityOil);
                            F_pr[m] = 0.4114489718 * Math.Pow((_saturationWaterOld[m] - _saturationWaterOst), 0.2) / (0.3428741432 * Math.Pow((_saturationWaterOld[m] - _saturationWaterOst), 1.2) + 34.11683017 * (_viscosityWater / _viscosityOil) * (Math.Pow((1 - _saturationWaterOld[m] - _saturationOilOst), 2.2))) - 0.3428741432 * Math.Pow((_saturationWaterOld[m] - _saturationWaterOst), 1.2) * (0.4114489718 * Math.Pow((_saturationWaterOld[m] - _saturationWaterOst), 0.2) - 75.05702637 * (_viscosityWater / _viscosityOil) * (Math.Pow((1 - _saturationWaterOld[m] - _saturationOilOst), 1.2))) / (Math.Pow((0.3428741432 * Math.Pow((_saturationWaterOld[m] - _saturationWaterOst), 1.2) + 34.11683017 * (_viscosityWater / _viscosityOil) * (Math.Pow((1 - _saturationWaterOld[m] - _saturationOilOst), 2.2))), 2));
                            Pcap_pr[m] = -8.779958120 * 100000 * Math.Exp(-25 * _saturationWaterOld[m]);

                        }

                        iter = 0;
                        bool dt_p = true;
                        // boolean it=
                        while (iter < 20 && dt_p == true)//12)
                        {
                            // /iter=0
                            #region iter=0
                            if (iter == 0)
                            {
                                for (Int32 i = 0; i < _n + 1; i++)
                                {
                                    p_iter_old[i] = _pressureOilNew[i];
                                    _pressureOilOld[i] = _pressureOilNew[i];
                                    p_iter_new[i] = p_iter_old[i];
                                }
                                iter++;
                            }
                            #endregion
                            // /iter>0
                            #region iter>0
                            if (iter != 0)
                            {


                                for (Int32 i = 0; i <= _n; i++)/////////////////////////
                                {
                                    m_old[i] = Density._porosity * (1 + Density._compresFactor * (_pressureOilOld[i] - _pressure0));
                                    densityWaterOld[i] = Density._densityWater0 * (1 + Density._compresFactorWater0 * (_pressureOilOld[i] - PressCap[i] - _pressure0));
                                    densityOilOld[i] = Density._densityOil0 * (1 + Density._compresFactorOil0 * (_pressureOilOld[i] - _pressure0));//заменить методами

                                    m[i] = Density._porosity * (1 + Density._compresFactor * (p_iter_old[i] - _pressure0));
                                    densityWater[i] = Density._densityWater0 * (1 + Density._compresFactorWater0 * (p_iter_old[i] - PressCap[i] - _pressure0));//1
                                    densityOil[i] = Density._densityOil0 * (1 + Density._compresFactorOil0 * (p_iter_old[i] - _pressure0));//заменить методами //1

                                    K1_oil[i] = _permeability * PermOil[i] / _viscosityOil;
                                    K1_water[i] = _permeability * PermWater[i] / _viscosityWater;
                                    denW_pr[i] = Density._compresFactorWater0 * Density._densityWater0; //0
                                    denO_pr[i] = Density._compresFactorOil0 * Density._densityOil0; //0
                                    denW_pr_ob[i] = -1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[i] * densityWater[i]);
                                    denO_pr_ob[i] = -1 * Density._compresFactorOil0 * Density._densityOil0 / (densityOil[i] * densityOil[i]);


                                }
                                //коэф прогонки

                                A[0] = 0;
                                AA[0] = 0;

                                double b0 = _permeability * PermOil[1] / _viscosityOil * 1 / du * (densityOil[1] / densityOil[0] + 1);
                                double b1 = _permeability * PermOil[1] / _viscosityOil * p_iter_old[1] / du * (1 / densityOil[0] * Density._compresFactorOil0 * Density._densityOil0);
                                double b2 = _permeability * PermOil[1] / _viscosityOil * p_iter_old[0] / du * (1 / densityOil[0] * Density._densityOil0 * Density._compresFactorOil0);
                                double b3 = _permeability * PermWater[1] / _viscosityWater * 1 / du * (densityWater[1] / densityWater[0] + 1);
                                double b4 = _permeability * PermWater[1] / _viscosityWater * p_iter_old[1] / du * (1 / densityWater[0] * Density._compresFactorWater0 * Density._densityWater0);
                                double b5 = _permeability * PermWater[1] / _viscosityWater * PressCap[1] / du * (1 / densityWater[0] * Density._densityWater0 * Density._compresFactorWater0);
                                double b6 = _permeability * PermWater[1] / _viscosityWater * p_iter_old[0] / du * (1 / densityWater[0] * Density._compresFactorWater0 * Density._densityWater0);
                                double b7 = _permeability * PermWater[1] / _viscosityWater * PressCap[0] / du * 0.5 * (1 / densityWater[0] * Density._densityWater0 * Density._compresFactorWater0);
                                double qq1 = dt * densityWater[0] * 2 * Math.PI * h * _permeability * PermWater[1] / _viscosityWater * 1 / du;
                                double qq2 = dt * 2 * Math.PI * h * _permeability * PermWater[1] / _viscosityWater * 1 / du;

                                if (t >= 864000)//2592000)//
                                {
                                    B[0] = -1 * dt * Math.PI * h * (b0 + b1 - b2) - dt * Math.PI * h * (b3 + b4 - b5 - b6 + b7);
                                }
                                else
                                {
                                    B[0] = -1 * dt * Math.PI * h * (b0 + b1 - b2) - dt * Math.PI * h * (b3 + b4 - b5 - b6 + b7) + qq1 - qq2; // +q3;//+q3
                                }

                                BB[0] = -1 * dt * Math.PI * h * (b0 + b1 - b2) - dt * Math.PI * h * (b3 + b4 - b5 - b6 + b7);

                                RR[0] = Rc * Rc * (Math.Exp(2 * (u[0] + du / 2)) - Math.Exp(2 * (u[0])));
                                double c0 = RR[0] * (1 - _saturationWaterOld[0]) * Density._porosity * Density._compresFactor;
                                double c1 = RR[0] * (1 - _saturationWaterOld[0]) * densityOilOld[0] * m_old[0] * (-1 * Density._compresFactorOil0 * Density._densityOil0 / (densityOil[0] * densityOil[0]));
                                double c2 = (_permeability * PermOil[1] / _viscosityOil * p_iter_old[1] / du * densityOil[1] * (-1 * Density._densityOil0 * Density._compresFactorOil0 / (densityOil[0] * densityOil[0])));
                                double c3 = _permeability * PermOil[1] / _viscosityOil * 1 / du * (densityOil[1] / densityOil[0] + 1);
                                double c4 = _permeability * PermOil[1] / _viscosityOil * p_iter_old[0] / du * densityOil[1] * (-1 * Density._compresFactorOil0 * Density._densityOil0 / (densityOil[0] * densityOil[0]));
                                double c6 = RR[0] * _saturationWaterOld[0] * densityWaterOld[0] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[0] * densityWater[0])) * m_old[0];
                                double c5 = RR[0] * _saturationWaterOld[0] * Density._porosity * Density._compresFactor;
                                double c7 = (_permeability * PermWater[1] / _viscosityWater * p_iter_old[1] / du * densityWater[1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[0] * densityWater[0])));
                                double c8 = (_permeability * PermWater[1] / _viscosityWater * PressCap[1] / du * densityWater[1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[0] * densityWater[0])));
                                double c9 = (_permeability * PermWater[1] / _viscosityWater * 1 / du * (1 + densityWater[1] / densityWater[0]));
                                double c10 = (_permeability * PermWater[1] / _viscosityWater * p_iter_old[0] / du * densityWater[1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[0] * densityWater[0])));
                                double c11 = (_permeability * PermWater[1] / _viscosityWater * PressCap[0] / du * densityWater[1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[0] * densityWater[0])));

                                double c13 = (-1 * Density._compresFactorOil0 * Density._densityOil0 / (densityOil[0] * densityOil[0]));

                                double c19 = dt * 2 * Math.PI * h * _permeability * PermWater[1] / _viscosityWater * 1 / du;
                                double qq3 = dt * 2 * Math.PI * h * _permeability * PermWater[1] / _viscosityWater * (p_iter_old[1] - PressCap[1] - p_iter_old[0] + PressCap[0]) / du;
                                double qq4 = dt * 2 * densityWater[0] * Math.PI * h * _permeability * PermWater[1] / _viscosityWater * 1 / du;
                                double qq5 = dt * 2 * Math.PI * h * _permeability * PermWater[1] / _viscosityWater * 1 / du;

                                if (t >= 864000)//2592000)//
                                {
                                    C[0] = c0 * Math.PI * h - c1 * Math.PI * h - (dt * Math.PI * h * c2 - dt * Math.PI * h * c3 - dt * Math.PI * h * c4) + Math.PI * h * c5 - Math.PI * h * c6 - (dt * Math.PI * h * c7 - dt * Math.PI * h * c8 - dt * Math.PI * h * c9 - dt * Math.PI * h * c10 + dt * Math.PI * h * c11);
                                }

                                else
                                {
                                    C[0] = c0 * Math.PI * h - c1 * Math.PI * h - (dt * Math.PI * h * c2 - dt * Math.PI * h * c3 - dt * Math.PI * h * c4) + Math.PI * h * c5 - Math.PI * h * c6 - (dt * Math.PI * h * c7 - dt * Math.PI * h * c8 - dt * Math.PI * h * c9 - dt * Math.PI * h * c10 + dt * Math.PI * h * c11) + qq3 - qq4 + qq5;// +c16 - c19; //c17 - c18 - c19;//+ c16 + dt * _densityWaterSt / densityWater[0] * q4 + dt * _densityWaterSt * q4_0;//+c17; //+ c15;//+ c12
                                }


                                CC[0] = c0 * Math.PI * h - c1 * Math.PI * h - (dt * Math.PI * h * c2 - dt * Math.PI * h * c3 - dt * Math.PI * h * c4) + Math.PI * h * c5 - Math.PI * h * c6 - (dt * Math.PI * h * c7 - dt * Math.PI * h * c8 - dt * Math.PI * h * c9 - dt * Math.PI * h * c10 + dt * Math.PI * h * c11);// +c16; 

                                double f0 = Math.PI * h * RR[0] * (1 - _saturationWaterOld[0]) * (m[0] - (densityOilOld[0] / densityOil[0]) * m_old[0]);
                                double f1 = dt * Math.PI * h * (_permeability * PermOil[1] / _viscosityOil * (p_iter_old[1] - p_iter_old[0]) / du * (densityOil[1] / densityOil[0] + 1));
                                double f2 = Math.PI * h * RR[0] * _saturationWaterOld[0] * (m[0] - densityWaterOld[0] / densityWater[0] * m_old[0]);
                                double f3 = dt * Math.PI * h * (_permeability * PermWater[1] / _viscosityWater * (p_iter_old[1] - PressCap[1] - p_iter_old[0] + PressCap[0]) / du * (densityWater[1] / densityWater[0] + 1));
                                double f4 = dt * _densityOilSt / densityOil[0] * Q;
                                double f5 = dt * 2 * Math.PI * h * (_permeability * PermWater[1] / _viscosityWater) * (p_iter_old[1] - PressCap[1] - p_iter_old[0] + PressCap[0]) / du;
                                double qq6 = dt * Qsum - dt * densityWater[1] * 2 * Math.PI * h * (_permeability * PermWater[0] / _viscosityWater) * (p_iter_old[1] - PressCap[1] - p_iter_old[0] + PressCap[0]) / du;
                                double qq7 = dt * 2 * Math.PI * h * (_permeability * PermWater[1] / _viscosityWater) * (p_iter_old[1] - PressCap[1] - p_iter_old[0] + PressCap[0]) / du;

                                if (t >= 864000)//2592000)//
                                {
                                    F2[0] = -1 * (f0 - f1 + f2 - f3);
                                }
                                else
                                {
                                    F2[0] = -1 * (f0 - f1 + f2 - f3 + qq6 + qq7);
                                }


                                FF2[0] = -1 * (f0 - f1 + f2 - f3 + f4);

                                for (Int32 l = 1; l < _n; l++)
                                {
                                    A1[l] = -1 / dy[l] * (-K1_oil[l] * _pressureOilNew[l] / du * 0.5 * 1 / densityOil[l] * denO_pr[l] + K1_oil[l] * 1 / du * 0.5 + K1_oil[l] * 1 / du * 0.5 * 1 / densityOil[l] * (densityOil[l - 1] + _pressureOilNew[l - 1] * denO_pr[l - 1])) - 1 / dy[l] * (-K1_water[l] * _pressureOilNew[l] / du * 0.5 * 1 / densityWater[l] * denW_pr[l - 1] + K1_water[l] * PressCap[l] / du * 0.5 * 1 / densityWater[l] * denW_pr[l - 1] + K1_water[l] * 1 / du * 0.5 * (1 + densityWater[l - 1] / densityWater[l]) + K1_water[l] * _pressureOilNew[l - 1] / du * 0.5 * 1 / densityWater[l] * denW_pr[l - 1] - K1_water[l] * PressCap[l - 1] / du * 0.5 - K1_water[l] * PressCap[l - 1] / du * 1 / densityWater[l] * denW_pr[l - 1]);

                                    double al0 = dt * Math.PI * h * _permeability * PermOil[l] / _viscosityOil * (p_iter_old[l] / du * Density._compresFactorOil0 * Density._densityWater0 * 1 / densityOil[l]);
                                    double al1 = dt * Math.PI * h * _permeability * PermOil[l] / _viscosityOil * 1 / du * (densityOil[l - 1] / densityOil[l] + 1);
                                    double al2 = dt * Math.PI * h * _permeability * PermOil[l] / _viscosityOil * p_iter_old[l - 1] / du * Density._compresFactorOil0 * Density._densityOil0 * 1 / densityOil[l];
                                    double al3 = dt * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * (p_iter_old[l] - PressCap[l]) / du * Density._compresFactorWater0 * Density._densityWater0 * 1 / densityWater[l];
                                    double al4 = dt * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * 1 / du * (1 + densityWater[l - 1] / densityWater[l]);
                                    double al5 = dt * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * (p_iter_old[l - 1] - PressCap[l - 1]) / du * Density._compresFactorWater0 * Density._densityWater0 * 1 / densityWater[l];

                                    A[l] = al0 - al1 - al2 + al3 - al4 - al5;
                                    AA[l] = A[l];


                                    C2[l] = r[l] * r[l] * (1 - _saturationWaterOld[l]) * 1 / dt * Density._porosity * Density._compresFactor - r[l] * r[l] * (1 - _saturationWaterOld[l]) * densityOil[l] * m_old[l] / dt * (-1 * Density._compresFactorOil0 * Density._densityOil0 / (densityOil[l] * densityOil[l])) - 1 / dy[l] * (_permeability * PermOil[l + 1] / _viscosityOil * _pressureOilNew[l + 1] / du * 0.5 * densityOil[l + 1] * (-1 * Density._densityOil0 * Density._compresFactorOil0 / (densityOil[l] * densityOil[l])) - _permeability * PermOil[l + 1] / _viscosityOil * 1 / du * 0.5 * (densityOil[l + 1] / densityOil[l] + 1) - _permeability * PermOil[l + 1] / _viscosityOil * _pressureOilNew[l] / du * 0.5 * densityOil[l + 1] * (-1 * Density._compresFactorOil0 * Density._densityOil0 / (densityOil[l] * densityOil[l])) - _permeability * PermOil[l] / _viscosityOil * 1 / du * 0.5 * (1 + densityOil[l - 1] / densityOil[l]) - _permeability * PermOil[l] / _viscosityOil * _pressureOilNew[l] / du * 0.5 * densityOil[l - 1] * (-1 * Density._densityOil0 * Density._compresFactorOil0 / (densityOil[l] * densityOil[l])) + _permeability * PermOil[l] / _viscosityOil * _pressureOilNew[l - 1] / du * 0.5 * densityOil[l - 1] * (Density._compresFactorOil0 * Density._densityOil0 / (densityOil[l] * densityOil[l]))) + r[l] * r[l] * _saturationWaterOld[l] * (1 / dt) * Density._porosity * Density._compresFactor - r[l] * r[l] * _saturationWaterOld[l] * densityWaterOld[l] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l])) * m_old[l] / dt - 1 / dy[l] * (_permeability * PermWater[l + 1] / _viscosityWater * _pressureOilNew[l + 1] / du * 0.5 * densityWater[l + 1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[l] * densityWater[l])) - _permeability * PermWater[l + 1] / _viscosityWater * PressCap[l + 1] / du * 0.5 * densityWater[l + 1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l])) - _permeability * PermWater[l + 1] / _viscosityWater * 1 / du * 0.5 * (1 + densityWater[l + 1] / densityWater[l]) - _permeability * PermWater[l + 1] / _viscosityWater * _pressureOilNew[l] / du * 0.5 * densityWater[l + 1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[l] * densityWater[l])) + _permeability * PermWater[l + 1] / _viscosityWater * PressCap[l] / du * 0.5 * densityWater[l + 1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l])) - _permeability * PermWater[l] / _viscosityWater * 1 / du * 0.5 * (1 + densityWater[l - 1] / densityWater[l]) - _permeability * PermWater[l] / _viscosityWater * _pressureOilNew[l] / du * 0.5 * densityWater[l - 1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[l] * densityWater[l])) + _permeability * PermWater[l] / _viscosityWater * PressCap[l] / du * 0.5 * densityWater[l - 1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l])) + _permeability * PermWater[l] / _viscosityWater * _pressureOilNew[l - 1] / du * 0.5 * densityWater[l - 1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[l] * densityWater[l])) - _permeability * PermWater[l] / _viscosityWater * PressCap[l - 1] / du * 0.5 * densityWater[l - 1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l])));
                                    /// dy[1] * (_permeability * PermWater[2] * (densityWater[2] + densityWater[1]) / (2 * densityWater[1] * _viscosityWater * dx[2]) + _permeability * PermWater[1] * (densityWater[1] + dens._densityWater0) / (2 * densityWater[1] * _viscosityWater * dx[1]) + _permeability * PermOil[2] * (densityOil[2] + densityOil[1]) / (2 * densityOil[1] * _viscosityOil * dx[2]) + _permeability * PermOil[1] * (densityOil[1] + dens._densityOil0) / (2 * densityOil[1] * _viscosityOil * dx[1]));
                                    RR[l] = Rc * Rc * (Math.Exp(2 * (u[l] + du / 2)) - Math.Exp(2 * (u[l] - du / 2)));

                                    double cl0 = Math.PI * h * (1 - _saturationWaterOld[l]) * Density._porosity * Density._compresFactor * RR[l];
                                    double cl1 = Math.PI * h * (1 - _saturationWaterOld[l]) * densityOilOld[l] * m_old[l] * (-1 * Density._compresFactorOil0 * Density._densityOil0 / (densityOil[l] * densityOil[l])) * RR[l];
                                    double cl2 = dt * Math.PI * h * _permeability * PermOil[l + 1] / _viscosityOil * _pressureOilNew[l + 1] / du * densityOil[l + 1] * (-1 * Density._densityOil0 * Density._compresFactorOil0 / (densityOil[l] * densityOil[l]));
                                    double cl3 = dt * Math.PI * h * _permeability * PermOil[l + 1] / _viscosityOil * 1 / du * (densityOil[l + 1] / densityOil[l] + 1);
                                    double cl4 = dt * Math.PI * h * _permeability * PermOil[l + 1] / _viscosityOil * _pressureOilNew[l] / du * densityOil[l + 1] * (-1 * Density._compresFactorOil0 * Density._densityOil0 / (densityOil[l] * densityOil[l]));
                                    double cl5 = dt * Math.PI * h * _permeability * PermOil[l] / _viscosityOil * 1 / du * (1 + densityOil[l - 1] / densityOil[l]);
                                    double cl6 = dt * Math.PI * h * _permeability * PermOil[l] / _viscosityOil * _pressureOilNew[l] / du * densityOil[l - 1] * (-1 * Density._densityOil0 * Density._compresFactorOil0 / (densityOil[l] * densityOil[l]));
                                    double cl7 = dt * Math.PI * h * _permeability * PermOil[l] / _viscosityOil * _pressureOilNew[l - 1] / du * densityOil[l - 1] * (-1 * Density._densityOil0 * Density._compresFactorOil0 / (densityOil[l] * densityOil[l]));// (Density._compresFactorOil0 * Density._densityOil0 / (densityOil[l] * densityOil[l]));
                                    double cl8 = Math.PI * h * _saturationWaterOld[l] * Density._porosity * Density._compresFactor * RR[l];
                                    double cl9 = Math.PI * h * _saturationWaterOld[l] * densityWaterOld[l] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l])) * m_old[l] * RR[l];
                                    double cl10 = dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * _pressureOilNew[l + 1] / du * densityWater[l + 1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[l] * densityWater[l]));
                                    double cl11 = dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * PressCap[l + 1] / du * densityWater[l + 1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l]));
                                    double cl12 = dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * 1 / du * (1 + densityWater[l + 1] / densityWater[l]);
                                    double cl13 = dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * _pressureOilNew[l] / du * densityWater[l + 1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[l] * densityWater[l]));
                                    double cl14 = dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * PressCap[l] / du * densityWater[l + 1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l]));
                                    double cl15 = dt * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * 1 / du * (1 + densityWater[l - 1] / densityWater[l]);
                                    double cl16 = dt * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * _pressureOilNew[l] / du * densityWater[l - 1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[l] * densityWater[l]));
                                    double cl17 = dt * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * PressCap[l] / du * densityWater[l - 1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l]));
                                    double cl18 = dt * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * _pressureOilNew[l - 1] / du * densityWater[l - 1] * (-1 * Density._densityWater0 * Density._compresFactorWater0 / (densityWater[l] * densityWater[l]));
                                    double cl19 = dt * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * PressCap[l - 1] / du * densityWater[l - 1] * (-1 * Density._compresFactorWater0 * Density._densityWater0 / (densityWater[l] * densityWater[l]));

                                    C[l] = cl0 - cl1 - cl2 + cl3 + cl4 + cl5 + cl6 - cl7 + cl8 - cl9 - cl10 + cl11 + cl12 + cl13 - cl14 + cl15 + cl16 - cl17 - cl18 + cl19;// -cl20 + cl21;
                                    CC[l] = C[l];

                                    B[l] = -1 * (dt * Math.PI * h * _permeability * PermOil[l + 1] / _viscosityOil * 1 / du * (densityOil[l + 1] / densityOil[l] + 1) + dt * Math.PI * h * _permeability * PermOil[l + 1] / _viscosityOil * _pressureOilNew[l + 1] / du * (1 / densityOil[l] * Density._compresFactorOil0 * Density._densityOil0) - dt * Math.PI * h * _permeability * PermOil[l + 1] / _viscosityOil * _pressureOilNew[l] / du * 1 / densityOil[l] * Density._densityOil0 * Density._compresFactorOil0) - (dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * 1 / du * (densityWater[l + 1] / densityWater[l] + 1) + dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * _pressureOilNew[l + 1] / du * (1 / densityWater[l] * Density._compresFactorWater0 * Density._densityWater0) - dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * PressCap[l + 1] / du * (1 / densityWater[l] * Density._densityWater0 * Density._compresFactorWater0) - dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * _pressureOilNew[l] / du * (1 / densityWater[l] * Density._compresFactorWater0 * Density._densityWater0) + dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * PressCap[l] / du * (1 / densityWater[l] * Density._densityWater0 * Density._compresFactorWater0));// - dt * _densityWaterSt / densityWater[l] * 2 * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * 1 / du;

                                    BB[l] = B[l];

                                    // F[l] = -1 * ((1 - _saturationWaterOld[l]) * (m[l] / dt - (dens.GetDensityOil(_pressureOilOld[l], _pressure1) / dens.GetDensityOil(p_iter_old[l], _pressure1)) * m_old[l] / dt) - 1 / dy[l] * (_permeability * PermWater[l + 1] / _viscosityOil * (p_iter_old[l + 1] - p_iter_old[l]) / dx[l + 1] * 0.5 * (dens.GetDensityOil(p_iter_old[l + 1], _pressure1) / dens.GetDensityOil(p_iter_old[l], _pressure1) + 1) - _permeability * PermOil[l] / _viscosityOil * (p_iter_old[l] - p_iter_old[l - 1]) / dx[l] * 0.5 * (1 + dens.GetDensityOil(p_iter_old[l - 1], _pressure1) / dens.GetDensityOil(p_iter_old[l], _pressure1))) + r[l] * r[l] * _saturationWaterOld[l] * (m_old[l] / dt - dens.GetDensityWater(_pressureOilOld[l], _pressure1) / dens.GetDensityWater(p_iter_old[l], _pressure1) * m_old[l] / dt) - 1 / dy[l] * (_permeability * PermWater[l + 1] / _viscosityWater * (p_iter_old[l + 1] - PressCap[l + 1] - p_iter_old[l] + PressCap[l]) / dx[l + 1] * 0.5 * (dens.GetDensityWater(p_iter_old[l + 1], _pressure1) / dens.GetDensityWater(p_iter_old[l], _pressure1) + 1) - _permeability * PermWater[l] / _viscosityWater * (p_iter_old[l] - PressCap[l] - p_iter_old[l - 1] + PressCap[l - 1]) / dx[l] * 0.5 * (1 + dens.GetDensityWater(p_iter_old[l - 1], _pressure1) / dens.GetDensityWater(p_iter_old[l], _pressure1))));

                                    F3[l] = -1 * r[l] * r[l] * ((1 - _saturationWaterOld[l]) * (m[l] / dt - (densityOilOld[l] / densityOil[l]) * m_old[l] / dt) - 1 / dy[l] * (_permeability * PermOil[l + 1] / _viscosityOil * (p_iter_old[l + 1] - p_iter_old[l]) / du * 0.5 * (densityOil[l + 1] / densityOil[l] + 1) - _permeability * PermOil[l] / _viscosityOil * (p_iter_old[l] - p_iter_old[l - 1]) / du * 0.5 * (1 + densityOil[l - 1] / densityOil[l])) + r[l] * r[l] * _saturationWaterOld[l] * (m_old[l] / dt - densityWaterOld[l] / densityWater[l] * m_old[l] / dt) - 1 / dy[l] * (_permeability * PermWater[l + 1] / _viscosityWater * (p_iter_old[l + 1] - PressCap[l + 1] - p_iter_old[l] + PressCap[l]) / du * 0.5 * (densityWater[l + 1] / densityWater[l] + 1) - _permeability * PermWater[l] / _viscosityWater * (p_iter_old[l] - PressCap[l] - p_iter_old[l - 1] + PressCap[l - 1]) / du * 0.5 * (1 + densityWater[l - 1] / densityWater[l])));// - _pressure1 * A[1];

                                    double fl0 = Math.PI * h * (1 - _saturationWaterOld[l]) * (m[l] - (densityOilOld[l] / densityOil[l]) * m_old[l]) * RR[l];
                                    double fl1 = dt * Math.PI * h * _permeability * PermOil[l + 1] / _viscosityOil * (p_iter_old[l + 1] - p_iter_old[l]) / du * (densityOil[l + 1] / densityOil[l] + 1);
                                    double fl2 = dt * Math.PI * h * _permeability * PermOil[l] / _viscosityOil * (p_iter_old[l] - p_iter_old[l - 1]) / du * (1 + densityOil[l - 1] / densityOil[l]);
                                    double fl3 = Math.PI * h * _saturationWaterOld[l] * (m[l] - densityWaterOld[l] / densityWater[l] * m_old[l]) * RR[l];
                                    double fl4 = dt * Math.PI * h * _permeability * PermWater[l + 1] / _viscosityWater * (p_iter_old[l + 1] - PressCap[l + 1] - p_iter_old[l] + PressCap[l]) / du * (densityWater[l + 1] / densityWater[l] + 1);
                                    double fl5 = dt * Math.PI * h * _permeability * PermWater[l] / _viscosityWater * (p_iter_old[l] - PressCap[l] - p_iter_old[l - 1] + PressCap[l - 1]) / du * (1 + densityWater[l - 1] / densityWater[l]);
                                    F2[l] = -1 * (fl0 - (fl1 - fl2) + fl3 - (fl4 - fl5));// - dt * _densityOilSt / densityOil[l] * Q-dt*_densityWaterSt/densityWater[l]*2*Math.PI*h*_permeability*PermWater[l]/_viscosityWater*(_pressureOilNew[l+1]-PressCap[l+1]-_pressureOilNew[l]+PressCap[l])/du);
                                    FF2[l] = F2[l];

                                }


                                //правая граница
                                double an1 = dt * Math.PI * h * (1 / densityOil[_n]) * (_permeability * PermOil[_n] / _viscosityOil) * (_pressureOilNew[_n] / du) * denO_pr[_n - 1];
                                double an2 = dt * Math.PI * h * (_permeability * PermOil[_n] / _viscosityOil) * 1 / du * (densityOil[_n - 1] / densityOil[_n] + 1);
                                // double an3 = (1 / densityOil[_n]) * (_permeability * PermOil[_n] / _viscosityOil) * 1 / du * (densityOil[_n - 1] / 2);
                                double an4 = dt * Math.PI * h * (1 / densityOil[_n]) * (_permeability * PermOil[_n] / _viscosityOil) * (_pressureOilNew[_n - 1] / du) * denO_pr[_n - 1];
                                double an5 = dt * Math.PI * h * (1 / densityWater[_n]) * (_permeability * PermWater[_n] / _viscosityWater) * (_pressureOilNew[_n] / du) * denW_pr[_n - 1];
                                double an6 = dt * Math.PI * h * (1 / densityWater[_n]) * (_permeability * PermWater[_n] / _viscosityWater) * (PressCap[_n] / du) * denW_pr[_n - 1];
                                double an7 = dt * Math.PI * h * (1 / densityWater[_n]) * (_permeability * PermWater[_n] / _viscosityWater) * (1 / du) * (densityWater[_n - 1] / densityWater[_n] + 1);
                                //double an8 = (1 / densityWater[_n]) * (_permeability * PermWater[_n] / _viscosityWater) * (1 / du) * (densityWater[_n - 1] / 2);
                                double an9 = dt * Math.PI * h * (1 / densityWater[_n]) * (_permeability * PermWater[_n] / _viscosityWater) * (_pressureOilNew[_n - 1] / du) * denW_pr[_n - 1];
                                double an10 = dt * Math.PI * h * (1 / densityWater[_n]) * (_permeability * PermWater[_n] / _viscosityWater) * (PressCap[_n - 1] / du) * denW_pr[_n - 1];
                                A[_n] = an1 - an2 - an4 + an5 - an6 - an7 - an9 + an10;
                                AA[_n] = A[_n];

                                B[_n] = 0;
                                BB[_n] = B[_n];

                                RR[_n] = Rc * Rc * (Math.Exp(2 * (u[_n])) - Math.Exp(2 * (u[_n] - du / 2)));
                                double cn0 = (_permeability * PermOil[_n] / _viscosityOil);
                                double cn00 = (_permeability * PermWater[_n] / _viscosityWater);
                                double cn1 = Math.PI * h * (1 - _saturationWaterOld[_n]) * Density._porosity * Density._compresFactor * RR[_n];// - 
                                double cn2 = Math.PI * h * denO_pr_ob[_n] * (1 - _saturationWaterOld[_n]) * (densityOilOld[_n] * m_old[_n]) * RR[_n];
                                double cn3 = dt * Math.PI * h * denO_pr_ob[_n] * (_permeability * PermOil[_n] / _viscosityOil) * (_pressureOilNew[_n] / du) * (densityOil[_n - 1]);
                                double cn4 = dt * Math.PI * h * (1 / densityOil[_n]) * (_permeability * PermOil[_n] / _viscosityOil) * (1 / du) * (densityOil[_n - 1]);
                                double cn5 = dt * Math.PI * h * (_permeability * PermOil[_n] / _viscosityOil) * (1 / du);
                                double cn6 = dt * Math.PI * h * denO_pr_ob[_n] * cn0 * (_pressureOilNew[_n - 1] / du) * (densityOil[_n - 1]);
                                // double cn7 = (1 / densityOil[_n]) * cn0 * (1 / du) * (densityOil[_n - 1]);// -cn7
                                double cn8 = Math.PI * h * _saturationWaterOld[_n] * Density._porosity * Density._compresFactor * RR[_n];
                                double cn9 = Math.PI * h * denW_pr_ob[_n] * _saturationWaterOld[_n] * (densityWaterOld[_n] * m_old[_n]) * RR[_n];
                                double cn10 = dt * Math.PI * h * denW_pr_ob[_n] * cn00 * (_pressureOilNew[_n] / du) * (densityWater[_n - 1]);
                                double cn11 = dt * Math.PI * h * (1 / densityWater[_n]) * cn00 * (1 / du) * (densityWater[_n - 1]);
                                double cn12 = dt * Math.PI * h * cn00 * (1 / du);
                                double cn13 = dt * Math.PI * h * cn00 * (denW_pr_ob[_n]) * (PressCap[_n] / du) * (densityWater[_n - 1]);
                                double cn14 = dt * Math.PI * h * cn00 * denW_pr_ob[_n] * (_pressureOilNew[_n - 1] / du) * (densityWater[_n - 1]);
                                double cn15 = dt * Math.PI * h * denW_pr_ob[_n] * cn00 * (PressCap[_n - 1] / du) * (densityWater[_n - 1]);

                                C[_n] = cn1 - cn2 + cn3 + cn4 + cn5 - cn6 + cn8 - cn9 + cn10 + cn11 + cn12 - cn13 - cn14 + cn15;// -dt * _densityOilSt * Q * denO_pr_ob[_n];
                                CC[_n] = C[_n];

                                //F2[_n - 1] = -1 * ((1 - _saturationWaterOld[_n - 1]) * (m[_n - 1] / dt - (densityOilOld[_n - 1] / densityOil[_n - 1]) * m_old[_n - 1] / dt) - 1 / dy[_n - 1] * (_permeability * PermOil[_n] / _viscosityOil * (p_iter_old[_n] - p_iter_old[_n - 1]) / dx[_n] * 0.5 * (densityOil[_n] / densityOil[_n - 1] + 1) - _permeability * PermOil[_n - 1] / _viscosityOil * (p_iter_old[_n - 1] - p_iter_old[_n - 2]) / dx[_n - 1] * 0.5 * (1 + densityOil[_n - 2] / densityOil[_n - 1])) + _saturationWaterOld[_n - 1] * (m_old[_n - 1] / dt - densityWaterOld[_n - 1] / densityWater[_n - 1] * m_old[_n - 1] / dt) - 1 / dy[_n - 1] * (_permeability * PermWater[_n] / _viscosityWater * (p_iter_old[_n] - PressCap[_n] - p_iter_old[_n - 1] + PressCap[_n - 1]) / dx[_n] * 0.5 * (densityWater[_n] / densityWater[_n - 1] + 1) - _permeability * PermWater[_n - 1] / _viscosityWater * (p_iter_old[_n - 1] - PressCap[_n - 1] - p_iter_old[_n - 2] + PressCap[_n - 2]) / dx[_n - 1] * 0.5 * (1 + densityWater[_n - 2] / densityWater[_n - 1])));// -_pressure0 * B[_n - 1];// - _pressure1 * A[1];
                                double fn1 = Math.PI * h * (1 - _saturationWaterOld[_n]) * (m[_n] - (densityOilOld[_n] / densityOil[_n]) * m_old[_n]) * RR[_n];
                                double fn2 = dt * Math.PI * h * (_permeability * PermOil[_n] / _viscosityOil) * (p_iter_old[_n] - p_iter_old[_n - 1]) / du * (densityOil[_n - 1] / densityOil[_n] + 1);
                                double fn3 = Math.PI * h * _saturationWaterOld[_n] * (m[_n] - densityWaterOld[_n] / densityWater[_n] * m_old[_n]) * RR[_n];
                                double fn4 = dt * Math.PI * h * (_permeability * PermWater[_n] / _viscosityWater) * (p_iter_old[_n] - PressCap[_n] - p_iter_old[_n - 1] + PressCap[_n - 1]) / du * (densityWater[_n - 1] / densityWater[_n] + 1);
                                F2[_n] = -1 * (fn1 + fn2 + fn3 + fn4);// - dt * _densityOilSt / densityOil[_n] * Q);// проверить знаки !!!+fn2+fn4

                                FF2[_n] = F2[_n];

                                #endregion
                                // alfa, betta
                                #region alfa, betta
                                alfa[1] = (-1 * B[0]) / C[0];
                                betta[1] = F2[0] / C[0];

                                //double alf = (-1 * Bb) / Cc;
                                //double bet = Fff / Cc;

                                for (Int32 k = 1; k < _n; k++)
                                {
                                    alfa[k + 1] = (-1 * B[k]) / (A[k] * alfa[k] + C[k]);
                                    betta[k + 1] = (F2[k] - A[k] * betta[k]) / (A[k] * alfa[k] + C[k]);

                                }
                                #endregion

                                dp[_n] = (F2[_n] - A[_n] * betta[_n]) / (C[_n] + A[_n] * alfa[_n]);

                                for (Int32 g = _n - 1; g >= 0; g--)
                                {
                                    dp[g] = alfa[g + 1] * dp[g + 1] + betta[g + 1];

                                }

                                for (Int32 k = 0; k <= _n; k++)
                                {
                                    p_iter_new[k] = p_iter_old[k] + dp[k];

                                }

                                #region // проверка dt
                                if (!timesExist)
                                {
                                    for (Int32 l = 1; l < _n; l++)
                                    {
                                        VelisityOil[l] = -1 * _permeability * (PermOil[l + 1] / _viscosityOil) * (p_iter_new[l + 1] - p_iter_new[l]) / dr[l + 1];
                                        VelosityWater[l] = -1 * _permeability * (PermWater[l + 1] / _viscosityWater) * (p_iter_new[l + 1] - PressCap[l + 1] - p_iter_new[l] + PressCap[l]) / dr[l + 1];// -(-1 * _permeability * PermWater[l + 1]) / _viscosityWater * (PressCap[l + 1] - PressCap[l]) / dr[l + 1];
                                        Velosity[l] = VelisityOil[l] + VelosityWater[l];
                                    }
                                    //dt_max1 = (Density._porosity / Math.Abs(Velosity[0])) * (1 / F_pr[1]) * dy[0];
                                    //dt_max1 = (Density._porosity / (r[0] * Math.Abs(Velosity[0]))) * (1 / F_pr[1]) * dy_kv[0];

                                    dt_max1 = 1 / 2 * (Density._porosity / (Math.Abs(Velosity[0]))) * (1 / F_pr[1]) * RR[0];


                                    for (Int32 j = 1; j < _n; j++)// определяем dtmax как min по j
                                    {

                                        // dtm = (Density._porosity / Math.Abs(Velosity[j])) * (1 / F_pr[j + 1]) * dy[j];
                                        dtm = (Density._porosity / (r[j] * Math.Abs(Velosity[j]))) * (1 / F_pr[j + 1]) * dy_kv[j];


                                        if (dt_max1 > dtm)
                                        {
                                            dt_max1 = dtm;
                                        }

                                    }
                                    if (F_Psi[0] > F_Psi[1])
                                    {
                                        ps = F_Psi[0];
                                    }
                                    else
                                    {
                                        ps = F_Psi[1];
                                    }

                                    RR[0] = Rc * Rc * (Math.Exp(2 * (u[0] + du / 2)) - Math.Exp(2 * (u[0])));

                                    dt_max = 0.3364 * Density._porosity * _viscosityOil * du * du * RR[0] / (Math.Abs(Pcap_pr[0]) * _permeability * ps); //новое 0.6728


                                    for (Int32 j = 1; j < _n; j++)
                                    {
                                        if (F_Psi[j] > F_Psi[j + 1])
                                        {
                                            ps = F_Psi[j];
                                        }
                                        else
                                        {
                                            ps = F_Psi[j + 1];
                                        }
                                        RR[j] = Rc * Rc * (Math.Exp(2 * (u[j] + du / 2)) - Math.Exp(2 * (u[j] - du / 2)));

                                        //dtm = 0.6728 * Density._porosity * _viscosityOil * dr[j] * dr[j] * dy_kv[j] / (Math.Abs(Pcap_pr[j]) * r[j] * _permeability * ps); //0.6728
                                        dtm = 0.3364 * Density._porosity * _viscosityOil * du * du * RR[j] / (Math.Abs(Pcap_pr[0]) * _permeability * ps); //новое 0.6728

                                        if (dt_max > dtm)
                                        {
                                            dt_max = dtm;
                                        }
                                    }
                                    if (dt_max1 < dt_max)
                                    {
                                        dt_max = dt_max1;
                                    }
                                    double dt2 = dt;
                                    //сравниваем dt с dtmax
                                    if (dt_max < dt2)
                                    {
                                        dt2 = dt_max;//* 10
                                                     //bol = true;
                                    }
                                    if (dt2 < dt)
                                    {

                                        dt = dt2;
                                        iter = 0;
                                        dt_p = true;
                                    }
                                    else
                                    {

                                        dt_p = false;
                                    }

                                    //if (dt > 1000)
                                    //{
                                    //    dt = dt/2;
                                    //}
                                }

                                #endregion

                                // norma
                                #region
                                double SumKvOld = 0;
                                double SumKvNew = 0;
                                for (Int32 k = 0; k <= _n; k++)
                                {
                                    SumKvOld = SumKvOld + p_iter_old[k] * p_iter_old[k];
                                    SumKvNew = SumKvNew + p_iter_new[k] * p_iter_new[k];
                                }

                                double norm;
                                norm = Math.Abs(Math.Sqrt(SumKvOld) - Math.Sqrt(SumKvNew));
                                for (Int32 k = 0; k <= _n; k++)
                                {
                                    p_iter_old[k] = p_iter_new[k];
                                }

                                if (norm < 0.001)
                                {
                                    for (Int32 k = 0; k <= _n; k++)
                                    {
                                        _pressureOilNew[k] = p_iter_new[k];
                                    }

                                    break;
                                }
                                else
                                {
                                    iter++;
                                }
                                #endregion
                            }


                            //iter++;


                            for (Int32 k = 0; k <= _n; k++)
                            {
                                p_iter_old[k] = p_iter_new[k];

                            }
                        }
                        dt_pr = dt;
                        dataExcel.Add("iter2", new double[] { iter });
                        //dataExcel.Add("dt_pres", new double[] { dt });
                        for (Int32 k = 0; k <= _n; k++)
                        {
                            _pressureOilOld[k] = _pressureOilNew[k];

                        }


                        for (Int32 k = 0; k <= _n; k++)
                        {
                            _pressureOilNew[k] = p_iter_new[k];

                        }
                        for (Int32 k = 0; k < _n + 1; k++)
                        {
                            _pressureWaterNew[k] = _pressureOilNew[k] - PressCap[k];
                        }


                        for (Int32 i = 0; i <= _n; i++)
                        {
                            m_old[i] = Density._porosity * (1 + Density._compresFactor * (_pressureOilOld[i] - _pressure1));
                            densityWaterOld[i] = Density._densityWater0 * (1 + Density._compresFactorWater0 * (_pressureOilOld[i] - PressCap[i] - _pressure1));
                            densityOilOld[i] = Density._densityOil0 * (1 + Density._compresFactorOil0 * (_pressureOilOld[i] - _pressure1));//заменить методами

                            m[i] = Density._porosity * (1 + Density._compresFactor * (_pressureOilNew[i] - _pressure1));
                            densityWater[i] = Density._densityWater0 * (1 + Density._compresFactorWater0 * (_pressureOilNew[i] - PressCap[i] - _pressure1));
                            densityOil[i] = Density._densityOil0 * (1 + Density._compresFactorOil0 * (_pressureOilNew[i] - _pressure1));//заменить методами
                        }

                        for (Int32 l = 1; l < _n; l++)
                        {
                            VelisityOil[l] = -1 * _permeability * (PermOil[l + 1] / _viscosityOil) * (_pressureOilNew[l + 1] - _pressureOilNew[l]) / dr[l + 1];
                            VelosityWater[l] = -1 * _permeability * (PermWater[l + 1] / _viscosityWater) * (_pressureOilNew[l + 1] - PressCap[l + 1] - _pressureOilNew[l] + PressCap[l]) / dr[l + 1];// -(-1 * _permeability * PermWater[l + 1]) / _viscosityWater * (PressCap[l + 1] - PressCap[l]) / dr[l + 1];
                            Velosity[l] = VelisityOil[l] + VelosityWater[l];
                        }

                        for (Int32 i = 0; i <= _n; i++)
                        {
                            ObvodnNew[i] = VelosityWater[i] / Velosity[i];// Vw[i] / UUU[i];
                        }

                        #region ШАГ
                        if (!timesExist) // Если Excel отсутствует
                        {
                            // dt_max1 = (Density._porosity / (r[0] * Math.Abs(Velosity[0]))) * (1 / F_pr[1]) * dy_kv[0];
                            dt_max1 = 1 / 2 * (Density._porosity / (Math.Abs(Velosity[0]))) * (1 / F_pr[1]) * RR[0];


                            for (Int32 j = 1; j < _n; j++)// определяем dtmax как min по j
                            {
                                dtm = (Density._porosity / (r[j] * Math.Abs(Velosity[j]))) * (1 / F_pr[j + 1]) * dy_kv[j];// !!!


                                if (dt_max1 > dtm)
                                {
                                    dt_max1 = dtm;
                                }

                            }
                            if (F_Psi[0] > F_Psi[1])
                            {
                                ps = F_Psi[0];
                            }
                            else
                            {
                                ps = F_Psi[1];
                            }

                            RR[0] = Rc * Rc * (Math.Exp(2 * (u[0] + du / 2)) - Math.Exp(2 * (u[0])));

                            dt_max = 0.3364 * Density._porosity * _viscosityOil * du * du * RR[0] / (Math.Abs(Pcap_pr[0]) * _permeability * ps); //новое 0.6728


                            for (Int32 j = 1; j < _n; j++)
                            {
                                if (F_Psi[j] > F_Psi[j + 1])
                                {
                                    ps = F_Psi[j];
                                }
                                else
                                {
                                    ps = F_Psi[j + 1];
                                }
                                RR[j] = Rc * Rc * (Math.Exp(2 * (u[j] + du / 2)) - Math.Exp(2 * (u[j] - du / 2)));

                                //dtm = 0.6728 * Density._porosity * _viscosityOil * dr[j] * dr[j] * dy_kv[j] / (Math.Abs(Pcap_pr[j]) * r[j] * _permeability * ps); //0.6728
                                dtm = 0.3364 * Density._porosity * _viscosityOil * du * du * RR[j] / (Math.Abs(Pcap_pr[0]) * _permeability * ps); //новое 0.6728

                                if (dt_max > dtm)
                                {
                                    dt_max = dtm;
                                }
                            }
                            if (dt_max1 < dt_max)
                            {
                                dt_max = dt_max1;
                            }

                            //сравниваем dt с dtmax
                            if (dt_max < dt)
                            {
                                dt = dt_max;//* 10
                                            //bol = true;
                            }
                            //if (dt > 1000)
                            //{
                            //    dt = dt/2;
                            //}
                            //if (t > 864000 && dt > 1000)//2592000 && dt > 1000)//
                            //{
                            //    dt = dt / 10;
                            //}
                        }
                        #endregion
                        dt_s = dt;
                        //dataExcel.Add("dt_sat", new double[] { dt });

                        #region new насыщенность

                        if (_saturationWaterOld[0] < (1 - _saturationOilOst) || Qsum == 0)
                        {

                            //double g1 = _saturationWaterOld[0] * (-1 * (m_old[0] * densityWaterOld[0]) / (m[0] * densityWater[0]) + 1);

                            double g1 = _saturationWaterOld[0] * (1 - (m_old[0] * densityWaterOld[0]) / (m[0] * densityWater[0]));
                            // double g2 = dt / (RR[0] * m[0] * densityWater[0]) * (_permeability / _viscosityWater) * PermWater[1] * (_pressureWaterNew[1] - _pressureWaterNew[0]) / du * (densityWater[1] + densityWater[0]);
                            double g2 = dt / (RR[0] * m[0]) * (_permeability / _viscosityWater) * PermWater[1] * (_pressureWaterNew[1] - _pressureWaterNew[0]) / du * (1 + densityWater[1] / densityWater[0]);


                            _saturationWaterNew[0] = _saturationWaterOld[0] - g1 + g2;
                            if (_saturationWaterNew[0] > (1 - _saturationOilOst))
                            {
                                _saturationWaterNew[0] = 1 - _saturationOilOst;
                            }
                        }
                        if (_saturationWaterOld[0] >= (1 - _saturationOilOst) && Qsum != 0)
                        {
                            _saturationWaterNew[0] = 1 - _saturationOilOst;// _saturationWaterOld[0];
                        }
                        Double[] gg1 = new Double[_n + 1];
                        Double[] gg2 = new Double[_n + 1];
                        Double[] gg3 = new Double[_n + 1];
                        Double[] gg4 = new Double[_n + 1];

                        for (Int32 i = 1; i < _n; i++)
                        {
                            // gg1[i] = _saturationWaterOld[i] * (-1 * (m_old[i] * densityWaterOld[i]) / (m[i] * densityWater[i]) + 1);
                            gg1[i] = _saturationWaterOld[i] * (1 - (m_old[i] * densityWaterOld[i]) / (m[i] * densityWater[i]));
                            gg2[i] = dt / (RR[i] * m[i] * densityWater[i]);
                            gg3[i] = (_permeability / _viscosityWater) * PermWater[i + 1] * (_pressureWaterNew[i + 1] - _pressureWaterNew[i]) / du * (densityWater[i + 1] + densityWater[i]);
                            gg4[i] = (_permeability / _viscosityWater) * PermWater[i] * (_pressureWaterNew[i] - _pressureWaterNew[i - 1]) / du * (densityWater[i - 1] + densityWater[i]);
                            _saturationWaterNew[i] = _saturationWaterOld[i] - gg1[i] + gg2[i] * (gg3[i] - gg4[i]); //_saturationWaterOld[i] * (1 - (m_old[i] * densityWaterOld[i]) / (m[i] * densityWater[i])) + dt / (RR[i] * m[i]) * ((_permeability / _viscosityOil) * PermWater[i + 1] * (_pressureWaterNew[i + 1] - _pressureWaterNew[i]) / du * (densityWater[i + 1] / densityWater[i] + 1) - (_permeability / _viscosityWater) * PermWater[i] * (_pressureWaterNew[i] - _pressureWaterNew[i - 1]) / du * (1 + densityWater[i - 1] / densityWater[i]));

                        }
                        _saturationWaterNew[_n] = _saturationWaterOld[_n] - _saturationWaterOld[_n] * (-1 * (m_old[_n] * densityWaterOld[_n]) / (m[_n] * densityWater[_n]) + 1) + dt / (RR[_n] * m_old[_n] * densityWaterOld[_n]) * (-1 * (_permeability / _viscosityWater) * PermWater[_n] * (_pressureWaterNew[_n] - _pressureWaterNew[_n - 1]) / du * (densityWater[_n - 1] + densityWater[_n]));


                        #endregion
                        //

                        for (Int32 l = 0; l <= _n; l++)
                        {
                            PermWater[l] = 0.05 * Math.Pow((_saturationWaterNew[l] - _saturationWaterOst), 1.2) / Math.Pow((1 - _saturationOilOst - _saturationWaterOst), 1.2);
                            PermOil[l] = Math.Pow((1 - _saturationWaterNew[l] - _saturationOilOst), 2.2) / Math.Pow((1 - _saturationWaterOst - _saturationOilOst), 2.2);

                        }




                        for (Int32 i = 0; i <= _n; i++)
                        {
                            MO = MO + (densityOil[i] * m[i] * Math.PI * h * Rc * Rc * (Math.Exp(2 * (u[i] + du / 2)) - Math.Exp(2 * (u[i] - du / 2))));
                            M_oldO = M_oldO + (densityOilOld[i] * m_old[i] * Math.PI * h * Rc * Rc * (Math.Exp(2 * (u[i] + du / 2)) - Math.Exp(2 * (u[i] - du / 2))));
                            MW = MW + (densityWater[i] * m[i] * Math.PI * h * Rc * Rc * (Math.Exp(2 * (u[i] + du / 2)) - Math.Exp(2 * (u[i] - du / 2))));
                            M_oldW = M_oldW + (densityWaterOld[i] * m_old[i] * Math.PI * h * Rc * Rc * (Math.Exp(2 * (u[i] + du / 2)) - Math.Exp(2 * (u[i] - du / 2))));

                        }
                        mmO = (MO - M_oldO - Q * dt) / M_oldO;
                        mmW = (MW - M_oldW - (2 * Math.PI * h * (_permeability * PermWater[1] / _viscosityWater) * (_pressureOilNew[1] - PressCap[1] - _pressureOilNew[0] + PressCap[0]) / du) * dt) / M_oldW;

                        Q_w = 2 * Math.PI * h * (_permeability * PermWater[1] / _viscosityWater) * (_pressureOilNew[1] - PressCap[1] - _pressureOilNew[0] + PressCap[0]) / du;
                        Q_o = 2 * Math.PI * h * (_permeability * PermOil[1] / _viscosityOil) * (_pressureOilNew[1] - _pressureOilNew[0]) / du;


                        Q_w_d = 2 * Math.PI * h * (_permeability * PermWater[_n] / _viscosityWater) * (_pressureOilNew[_n] - PressCap[_n] - _pressureOilNew[0] + PressCap[0]) / (Math.Log(Rc / rs));
                        Q_o_d = 2 * Math.PI * h * (_permeability * PermOil[_n] / _viscosityOil) * (_pressureOilNew[_n] - _pressureOilNew[0]) / (Math.Log(Rc / rs));

                    } // конец else
                } // конец t > 0
                dataExcel.Add("dt_sat", new double[] { dt_s });
                dataExcel.Add("dt_pres", new double[] { dt_pr });

                for (Int32 g = 0; g <= _n; g++)
                {
                    _saturationWaterOld[g] = _saturationWaterNew[g];
                }
                //save_toFile("j = " + jj.ToString(), _saturationWaterNew_filePath);
                //save_toFile("dt = " + dt.ToString(), _saturationWaterNew_filePath);
                //save_toFile("T = " + t.ToString(), _saturationWaterNew_filePath);
                // dataExcel.Add("iter2", new double[] { iter });

                dataExcel.Add("j", new double[] { jj });
                dataExcel.Add("dt", new double[] { dt });
                dataExcel.Add("T", new double[] { t });


                //string text = "";
                //for (Int32 p = 0; p < _pressureOilNew.GetLength(0); p++)
                //{
                //    text = text + _pressureOilNew[p].ToString() + " " + " \r";
                //}
                //textBox1.Text = text;
                //save_toFile("T = " + t.ToString(), _saturationWaterNew_filePath);
                //save_toFile("Поле давлений нефти = ", _saturationWaterNew_filePath);
                //save_toFile(text, _saturationWaterNew_filePath);
                dataExcel.Add("Поле давлений нефти", (double[])_pressureOilNew.Clone());

                //string textt = "";
                //for (Int32 p = 0; p < _pressureWaterNew.GetLength(0); p++)
                //{
                //    textt = textt + _pressureWaterNew[p].ToString() + " " + " \r";
                //}
                ////textBox1.Text = text;
                //save_toFile("Поле давлений воды = ", _saturationWaterNew_filePath);
                //save_toFile(textt, _saturationWaterNew_filePath);
                //save_toFile("\n ", _saturationWaterNew_filePath);

                dataExcel.Add("Поле давлений воды", (double[])_pressureWaterNew.Clone());

                //save_toFile("Баланс нефти = " + mmO.ToString(), _saturationWaterNew_filePath);
                //save_toFile("Баланс воды = " + mmW.ToString(), _saturationWaterNew_filePath);
                //save_toFile("Дебит воды = " + Q_w.ToString(), _saturationWaterNew_filePath);
                //save_toFile("Дебит нефти = " + Q_o.ToString(), _saturationWaterNew_filePath);

                //save_toFile("Дебит воды (Дюпюи) = " + Q_w_d.ToString(), _saturationWaterNew_filePath);

                //save_toFile("Дебит нефти (Дюпюи) = " + Q_o_d.ToString(), _saturationWaterNew_filePath);



                //save_toFile("\n ", _saturationWaterNew_filePath);
                dataExcel.Add("Баланс нефти", new double[] { mmO });
                dataExcel.Add("Баланс воды", new double[] { mmW });
                dataExcel.Add("Дебит воды", new double[] { Q_w });
                dataExcel.Add("Дебит нефти", new double[] { Q_o });
                dataExcel.Add("Дебит воды (Дюпюи)", new double[] { Q_w_d });
                dataExcel.Add("Дебит нефти (Дюпюи)", new double[] { Q_o_d });


                //string text1 = "";
                //for (Int32 q = 0; q < _saturationWaterNew.GetLength(0); q++)
                //{
                //    text1 = text1 + _saturationWaterNew[q].ToString() + " " + " \r";
                //}
                //save_toFile("Насыщенность = ", _saturationWaterNew_filePath);
                //save_toFile(text1, _saturationWaterNew_filePath);

                dataExcel.Add("Насыщенность", (double[])_saturationWaterNew.Clone());


                //string text0n = "";
                //for (Int32 q = 0; q < ObvodnNew.GetLength(0); q++)
                //{
                //    text0n = text0n + ObvodnNew[q].ToString() + " " + " \r";
                //}
                //save_toFile("Обводненность = ", _saturationWaterNew_filePath);
                //save_toFile(text0n, _saturationWaterNew_filePath);
                //save_toFile("\n ", _saturationWaterNew_filePath);
                dataExcel.Add("Обводненность", (double[])ObvodnNew.Clone());
                dataExcel.Add("Проницаемость нефти", (double[])PermOil.Clone());

                //save_toFile("\n ", _saturationWaterNew_filePath);
                jj = jj + 1;

                t = t + dt;
                timeIndex++;
                if (!timesExist) {
					dt = kt * dt;
                }
                else if (timesExist && timeIndex < timeArray.Count) {
                    dt = timeArray[timeIndex];
                }
                //
                //save_toFile("\n ", _saturationWaterNew_filePath);
                datas.Add(dataExcel);
                if (timeIndex == timeArray.Count)
                    break;

            }//end while

            try
            {
                Console.WriteLine("Создаем TXT файл");
                save_toFile(dataToExcel, datas, _saturationWaterNew_filePath);
            }
            catch (Exception)
            {
                Console.WriteLine("Готово. Экспорт в TXT не произошел");
            }
            //try
            //{
            Console.WriteLine("Создаем Excel файл");
            saveToExcel(dataToExcel, datas, "Results-" + now);
            Console.WriteLine("Готово");
            //}
            //catch (Exception)
            //{
            //    status.Content += "Готово. Экспорт в Excel не произошел";
            //}


        }

        private string get_filePath()
        {
            return "results/result-" + now + ".txt";
            //SaveFileDialog saveFileDialog = new SaveFileDialog();
            //saveFileDialog.Filter = "txt files (*.txt)|*.txt";
            //if (saveFileDialog.ShowDialog().Value && saveFileDialog.FileName.Length > 0)
            //    return saveFileDialog.FileName;
            //else
                //return "";
        }

        private void save_toFile(Dictionary<String, double[]> initialData, List<Dictionary<String, double[]>> data, string filePath)
        {
            StringBuilder builder = new StringBuilder();
            foreach (var item in initialData.Keys)
            {
                builder.AppendLine("");
                builder.AppendLine(item + "=");
                foreach (var value in initialData[item])
                {
                    builder.AppendLine(value.ToString());
                }
            }
            save_toFile(builder, filePath);
            builder = new StringBuilder();
            foreach (var listItem in data)
            {
                foreach (var item in listItem.Keys)
                {
                    builder.AppendLine("");
                    builder.AppendLine(item + "=");
                    foreach (var value in listItem[item])
                    {
                        builder.AppendLine(value.ToString());
                    }
                }
            }
            save_toFile(builder, filePath);
        }

        private void saveToExcel(Dictionary<String, double[]> initialData, List<Dictionary<String, double[]>> data, String name)
        {
            try
            {
                File.Delete(Directory.GetCurrentDirectory() + "\\" + name + ".xlsx");
            }
            catch (Exception)
            {

                Console.WriteLine("Готово, файл не удалился");
            }
            //var fileDownloadName = "report" + DateTime.Now.ToString() + ".xlsx";
            //var reportsFolder = "reports";

            //using (var package = createExcelPackage())
            //{
            //    package.SaveAs(new FileInfo(Path.Combine(reportsFolder, fileDownloadName)));
            //}
            

            var book = new ExcelPackage();
            var worksheet = book.Workbook.Worksheets.Add("Result");
            int row = 1;
            worksheet.Cells[row, 1].Value = "Зона проникновения";
            worksheet.Cells[1, 1, 1, 2].Merge = true;
            row++;

            #region Насыщенность
            worksheet.Cells[row, 1].Value = "Насыщенность";
            worksheet.Cells[2, 1, 2, 2].Merge = true;
            row++;
            worksheet.Cells[row, 1].Value = "u";
            for (int i = 0; i < initialData["u"].Length; i++)
            {
                worksheet.Cells[row + i + 1, 1].Value = initialData["u"][i];
            }
            // T
            int column = 1;
            List<string> usedKeys = new List<string>();
            for (int i = 0; i < data.Count; i++)
            {
                int index = i;
                int hourT = Convert.ToInt32(Math.Floor(data[i]["T"][0] / 3600));
                if (hoursTime.ContainsKey(hourT) && usedKeys.IndexOf(hoursTime[hourT]) == -1)
                {
                    if (i != 0 && hourT - data[i - 1]["T"][0] / 3600 < data[i]["T"][0] / 3600 - hourT)
                    {
                        index = i - 1;
                    }
                    usedKeys.Add(hoursTime[hourT]);
                    column++;
                    worksheet.Cells[row, column].Value = hoursTime[hourT];
                    for (int j = 0; j < data[index]["Насыщенность"].Length; j++)
                    {
                        if (double.IsNaN(data[index]["Насыщенность"][j]))
                            worksheet.Cells[row + j + 1, column].Value = "NAN";
                        else worksheet.Cells[row + j + 1, column].Value = data[index]["Насыщенность"][j];
                    }
                }
            }
            #endregion

            column = 5;
            row += initialData["u"].Length + 1;

            #region Три столбца
            worksheet.Cells[row, column + 1].Value = "Шаг на входе";//Обводненность
            worksheet.Cells[row, column + 1, row, column + 2].Merge = true;
            worksheet.Cells[row, column + 5].Value = "Шаг давление";//Дебит нефти
            worksheet.Cells[row, column + 5, row, column + 6].Merge = true;
            worksheet.Cells[row, column + 9].Value = "Шаг насыщенность";//Дебит воды
            worksheet.Cells[row, column + 9, row, column + 10].Merge = true;
            //worksheet.Cell(row, column + 13).Value = "Дебит воды";
            //worksheet.Range(worksheet.Cell(row, column + 13).Address, worksheet.Cell(row, column + 14).Address).Row(1).Merge();
            row++;
            worksheet.Cells[row, column].Value = "Сутки";
            worksheet.Cells[row, column + 1].Value = "Часы";

            worksheet.Cells[row, column + 4].Value = "Сутки";
            worksheet.Cells[row, column + 5].Value = "Часы";

            worksheet.Cells[row, column + 8].Value = "Сутки";
            worksheet.Cells[row, column + 9].Value = "Часы";

            //worksheet.Cell(row, column + 13).Value = "ОФП";

            usedKeys = new List<string>();
            //for (int i = 0; i < data.Count; i++)
            //{
            //    int index = i;
            //    int hourT = Convert.ToInt32(Math.Floor(data[i]["T"][0] / 3600));
            //    if (hoursTime.ContainsKey(hourT) && usedKeys.IndexOf(hoursTime[hourT]) == -1)
            //    {
            //        if (i != 0 && hourT - data[i - 1]["T"][0] / 3600 < data[i]["T"][0] / 3600 - hourT)
            //        {
            //            index = i - 1;
            //        }
            //        usedKeys.Add(hoursTime[hourT]);
            //        row++;

            //        worksheet.Cell(row, column).Value = hoursTime[hourT];
            //        worksheet.Cell(row, column + 1).Value = hourT;
            //        if (double.IsNaN(data[index]["Обводненность"][1]))
            //            worksheet.Cell(row, column + 2).Value = "NAN";
            //        else worksheet.Cell(row, column + 2).Value = data[index]["Обводненность"][1];
            //        worksheet.Cell(row, column + 4).Value = hoursTime[hourT];
            //        worksheet.Cell(row, column + 5).Value = hourT;
            //        if (double.IsNaN(data[index]["Дебит нефти"][0]))
            //            worksheet.Cell(row, column + 6).Value = "NAN";
            //        else worksheet.Cell(row, column + 6).Value = data[index]["Дебит нефти"][0];
            //        worksheet.Cell(row, column + 8).Value = hoursTime[hourT];
            //        worksheet.Cell(row, column + 9).Value = hourT;
            //        if (double.IsNaN(data[index]["Дебит воды"][0]))
            //            worksheet.Cell(row, column + 10).Value = "NAN";
            //        else worksheet.Cell(row, column + 10).Value = data[index]["Дебит воды"][0];
            //    }
            //}
            var book1 = new ExcelPackage();
            var worksheet1 = book1.Workbook.Worksheets.Add("Result");
            worksheet1.Cells[1, 1].Value = "ОФП";
            //////////
            for (int i = 0; i < data.Count; i++)
            {
                int index = i;
                var hourT = data[i]["T"];//
                                         //int hourT = Convert.ToInt32(Math.Floor(data[i]["T"][0] / 3600));
                                         //if (hoursTime.ContainsKey(hourT) && usedKeys.IndexOf(hoursTime[hourT]) == -1)
                                         //{
                                         //    if (i != 0 && hourT - data[i - 1]["T"][0] / 3600 < data[i]["T"][0] / 3600 - hourT)
                                         //    {
                                         //        index = i - 1;
                                         //    }
                                         //    usedKeys.Add(hoursTime[hourT]);
                row++;

                worksheet.Cells[row, column].Value = hourT;
                worksheet.Cells[row, column + 1].Value = hourT;
                if (double.IsNaN(data[index]["dt"][0]))// Насыщенность [1]!!!!!!!!!!!!!!!!Поле давлений нефти Насыщенность Обводненность
                    worksheet.Cells[row, column + 2].Value = "NAN";
                else worksheet.Cells[row, column + 2].Value = data[index]["dt"][0]; //Насыщенность [1]!!!!!!!!!!!!!Поле давлений нефти Насыщенность
                worksheet.Cells[row, column + 4].Value = hourT;
                worksheet.Cells[row, column + 5].Value = hourT;
                if (double.IsNaN(data[index]["Поле давлений нефти"][0]))//dt_pres Дебит нефти
                    worksheet.Cells[row, column + 6].Value = "NAN";
                else worksheet.Cells[row, column + 6].Value = data[index]["Поле давлений нефти"][0];//dt_pres  Дебит нефти
                worksheet.Cells[row, column + 8].Value = hourT;
                worksheet.Cells[row, column + 9].Value = hourT;
                if (double.IsNaN(data[index]["dt_sat"][0]))//Дебит воды
                    worksheet.Cells[row, column + 10].Value = "NAN";
                else worksheet.Cells[row, column + 10].Value = data[index]["dt_sat"][0];//Дебит воды


                worksheet1.Cells[i + 2, 1].Value = hourT;
                worksheet1.Cells[i + 2, 2].Value = hourT;

                for (int j = 0; j < data[index]["Проницаемость нефти"].Length; j++)//Поле давлений воды Проницаемость нефти
                {
                    if (double.IsNaN(data[index]["Проницаемость нефти"][j]))
                        worksheet1.Cells[i + 2, j + 3].Value = "NAN";
                    else worksheet1.Cells[i + 2, j + 3].Value = data[index]["Проницаемость нефти"][j];
                }
                // }
            }
            #endregion

            #region Поле давлений нефти
            row++;
            worksheet.Cells[row, 1].Value = "давление нефти";
            worksheet.Cells[row, 1, row + 1, 2].Merge = true;
            row++;

            worksheet.Cells[row, 1].Value = "u";
            for (int i = 0; i < initialData["u"].Length; i++)
            {
                worksheet.Cells[row + i + 1, 1].Value = initialData["u"][i];
            }
            // T
            column = 1;
            usedKeys = new List<string>();
            for (int i = 0; i < data.Count; i++)
            {
                int index = i;
                int hourT = Convert.ToInt32(Math.Floor(data[i]["T"][0] / 3600));
                if (hoursTime.ContainsKey(hourT) && usedKeys.IndexOf(hoursTime[hourT]) == -1)
                {
                    if (i != 0 && hourT - data[i - 1]["T"][0] / 3600 < data[i]["T"][0] / 3600 - hourT)
                    {
                        index = i - 1;
                    }
                    usedKeys.Add(hoursTime[hourT]);
                    column++;
                    worksheet.Cells[row, column].Value = hoursTime[hourT];
                    for (int j = 0; j < data[index]["Поле давлений нефти"].Length; j++)
                    {
                        if (double.IsNaN(data[index]["Поле давлений нефти"][j]))
                            worksheet1.Cells[row + j + 1, column].Value = "NAN";
                        else worksheet.Cells[row + j + 1, column].Value = data[index]["Поле давлений нефти"][j];
                    }
                }
            }
            #endregion

            column = 5;
            row += initialData["u"].Length + 1;

            #region Поле давления воды

            row++;
            worksheet.Cells[row, 1].Value = "Поле давлений воды";//давление воды Проницаемость нефти
            worksheet.Cells[row, 1, row + 1, 2].Merge = true;
            row++;

            worksheet.Cells[row, 1].Value = "u";
            for (int i = 0; i < initialData["u"].Length; i++)
            {
                worksheet.Cells[row + i + 1, 1].Value = initialData["u"][i];
            }
            // T
            column = 1;
            usedKeys = new List<string>();
            for (int i = 0; i < data.Count; i++)
            {
                int index = i;
                int hourT = Convert.ToInt32(Math.Floor(data[i]["T"][0] / 3600));
                if (hoursTime.ContainsKey(hourT) && usedKeys.IndexOf(hoursTime[hourT]) == -1)
                {
                    if (i != 0 && hourT - data[i - 1]["T"][0] / 3600 < data[i]["T"][0] / 3600 - hourT)
                    {
                        index = i - 1;
                    }
                    usedKeys.Add(hoursTime[hourT]);
                    column++;
                    worksheet.Cells[row, column].Value = hoursTime[hourT];
                    for (int j = 0; j < data[index]["Поле давлений воды"].Length; j++)//Поле давлений воды Проницаемость нефти
                    {
                        if (double.IsNaN(data[index]["Поле давлений воды"][j]))
                            worksheet1.Cells[row + j + 1, column].Value = "NAN";
                        else worksheet.Cells[row + j + 1, column].Value = data[index]["Поле давлений воды"][j];//Поле давлений воды
                    }
                }
            }

           

            #endregion


            book.SaveAs(new FileInfo("reports/" + name + ".xlsx"));
            book1.SaveAs(new FileInfo("reports/" + name + "-насыщенность.xlsx"));
        }


        private void save_toFile(StringBuilder str, string filePath)
        {
            if (filePath != "" && filePath.Length > 0)
            {
                StreamWriter streamWriter = new StreamWriter(filePath, append: true);
                streamWriter.Write(str);
                streamWriter.Close();
            }
        }

        private void save_toFile(string str, string filePath)
        {
            if (filePath != "" && filePath.Length > 0)
            {
                StreamWriter streamWriter = new StreamWriter(filePath, append: true);
                streamWriter.WriteLine(str);
                streamWriter.Close();
            }
        }

        private void save_toFile(double[] array, string filePath)
        {
           
            var _array_string = string_fromArray_double(array);
            if (filePath != "" && filePath.Length > 0)
            {
                StreamWriter streamWriter = new StreamWriter(filePath, append: true);
                for (int l = 0; l < _array_string.Count - 1; l++)
                    streamWriter.Write(_array_string[l] + "; ");
                streamWriter.WriteLine(_array_string[_array_string.Count - 1]);

                streamWriter.Close();
            }
        }

        private void save_toFile(List<double> array, string filePath = "")
        {
            save_toFile(array.ToArray(), filePath);
        }

        #region CreateNewFile Methods

        private static void CreateNewFile(out String filePath)
        {
            
            filePath = "input.txt";
        }

        #endregion // CreateNewFile Methods
    }
}


