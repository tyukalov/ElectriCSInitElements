//
//  Author:
//    Igor Tyukalov tyukalov@bk.ru
//
//  Copyright (c) 2016, sp_r00t
//
//  All rights reserved.
//
//  Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
//
//     * Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
//     * Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in
//       the documentation and/or other materials provided with the distribution.
//     * Neither the name of the [ORGANIZATION] nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.
//
//  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
//  "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
//  LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
//  A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR
//  CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
//  EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
//  PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
//  PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
//  LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
//  NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
//  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
//
using System;
using System.Collections.Generic;
using System.Numerics;
using System.Data;
using System.Data.SQLite;

namespace ElectriCSInitElements
{
	public class Init
	{
		/// <summary>
		/// The airway x0 x1.
		/// Словарь коэффициентов соотношений между сопротивлениями прямой и
		/// нулевой последовательности для различных типов линий. Значения:
		/// OCWR			- одноцепная линия без заземлённых тросов;
		/// OCSR			- то же, со стальными заземлёнными тросами;
		/// OCR			- то же, с хорошо проводящими заземлёнными тросами:
		/// Те же параметры с буквой Т вместо О ознаяают двухцепную линию
		/// </summary>
		static Dictionary<string,double> AirwayX0X1 	= new Dictionary<string, double>(){{"OCWR",3.5},{"OCSR",3},{"OCR",2},{"TCWR",5.5},{"TCSR",4.7},{"TCR",3}};
		static Dictionary<string,double> Ro					= new Dictionary<string, double>(){{"al",0.02994}, {"cu",0.0178}};
		public SQLiteConnection Connect { get; set; }

		public Init ()
		{
		}

        private void ValidInit(string name, out double result)
        {
            try
            {
                double.TryParse(name, out result);
            }
            catch
            {
                throw new InvalidInitArgument();
            }
        }

		private void Connected ()
		{
			Connect = new SQLiteConnection ("Data Source=reactance.dat; Version=3;");
			try {
				Connect.Open ();
			} catch (SQLiteException ex){
				throw new InvalidConnect();
			}
		}

		private void Initialize (dynamic obj, double resistance, double reactance, double zero_resistance, double zero_reactance)
		{
			obj.Resistance			= resistance;
			obj.Reactance			= reactance;
			obj.ZeroReactance	= zero_reactance;
			obj.ZeroResistance	= zero_resistance;
			obj.Impedance			= new Complex(resistance, reactance);
			obj.ZeroImpedance	= new Complex(zero_resistance, zero_reactance);
			obj.Abs						= obj.Impedance.Magnitude;
			obj.ZeroAbs				= obj.ZeroImpedance.Magnitude;
		}

		/// <summary>
		/// Inits the cable.
		/// </summary>
		/// <param name='obj'>
		/// Объект инициализируемого класса. Передаётся в конструкторе как this;
		/// </param>
		/// <param name='args'>
		/// Словарь аргументов:
		/// material					- материал жилы. Должен принимать значения "al" или "cu". Обязательный.
		/// cross_section		- сечение в мм2 из стандартного ряда от 1.5 (для меди) до 240. Обязательный.
		/// lenght						- длина в метрах;
		/// </param>
		public void InitCable(dynamic obj, Dictionary<string,string> args)
		{
			Connected();
			if (Connect.State == ConnectionState.Open){
				string cabtype;
				switch (args["material"]) {
				case "al": 
				{
					cabtype = "plumbum_shell";
					break;
				}
				case "cu":
				{
					cabtype = "steel_shell";
					break;
				}
				default:
					throw new InvalidInitArgument();
					//break;
				}
				SQLiteCommand cmd = new SQLiteCommand(Connect);
				cmd.CommandText = "SELECT resistance, reactance, zero_resistance, zero_reactance FROM cable WHERE material='" + args["material"] + "' AND type='" +cabtype + "' AND cross_section='" + args["cross_section"] + "' AND cores='4'";
				SQLiteDataReader reader = cmd.ExecuteReader();
//				while(reader.Read()){
//					Console.WriteLine ("resist -> " + reader["resistance"].ToString() + "react -> " + reader["reactance"].ToString());
//				}
                double lenght;
                //try
                //{
                //    double.TryParse(args["lenght"], out lenght);
                //}
                //catch
                //{
                //    throw new InvalidInitArgument();
                //}
                ValidInit(args["lenght"], out lenght);
				if (reader.Read())
				{
//					Dictionary<string,double> result = new Dictionary<string, double>(){{"resistance", lenght * (double)reader["resistance"]}, {"reactance",  lenght *(double) reader["reactance"]}, {"zero_resistance", lenght * (double) reader["zero_resistance"]}, {"zero_reactance",  lenght *(double) reader["zero_reactance"]}}; 
					Initialize(obj, lenght * (double)reader["resistance"], lenght *(double) reader["reactance"],  lenght * (double) reader["zero_resistance"], lenght *(double) reader["zero_reactance"]); 
					Connect.Dispose();
//					return result;
				}
				else
				{
					Connect.Dispose();
					throw new InvalidInit();
//					return null;
				}
			}
			Connect.Dispose();
//			throw new InvalidInit();
//			return null;
		}

		/// <summary>
		/// Inits the bus.
		/// </summary>
		/// <param name='obj'>
		/// Объект инициализируемого класса. Передаётся в конструкторе как this;
		/// </param>
		/// <param name='args'>
		/// Словарь параметров:
		/// amperage				- расчётный ток из ряда 250б 400, 630, 1250, 1600, 2500, 3200, 4000;
		/// lenght						- длина в метрах;
		/// </param>
		// TODO Доработать для более широково спектра шинопроводов!
		public void InitBus (dynamic obj, Dictionary<string,string> args)
		{
			double lenght, cs;
            //try {
            //    double.TryParse (args ["lenght"], out lenght);
            //    double.TryParse (args ["amperage"], out cs);
            //} catch {
            //    throw new InvalidInitArgument ();
            //}	
            ValidInit(args["lenght"], out lenght);
            ValidInit(args["amperage"], out cs);
			if (!(cs == 250 || cs == 400 || cs == 630 || cs == 1250 || cs == 1600 || cs == 2500 || cs == 3200 || cs == 4000)) {
				throw new InvalidInitArgument ();
			}
			Connected();
			if (Connect.State == ConnectionState.Open) {
				SQLiteCommand cmd = new SQLiteCommand(Connect);
				cmd.CommandText = "SELECT resistance, reactance, zero_resistance, zero_reactance FROM bus WHERE amperage='" + args["amperage"] + "'";
				SQLiteDataReader reader = cmd.ExecuteReader();
				if (reader.Read())
				{
//					Dictionary<string,double> result = new Dictionary<string, double>(){{"resistance", lenght * (double)reader["resistance"]}, {"reactance",  lenght *(double) reader["reactance"]}, {"zero_resistance", lenght * (double) reader["zero_resistance"]}, {"zero_reactance",  lenght *(double) reader["zero_reactance"]}}; 
					Initialize(obj, lenght * (double)reader["resistance"], lenght *(double) reader["reactance"], lenght * (double) reader["zero_resistance"], lenght *(double) reader["zero_reactance"]); 
					Connect.Dispose();
//					return result;
				}
				else
				{
					Connect.Dispose();
					throw new InvalidInit();
//					return null;
				}
			}
			Connect.Dispose();
//			return null;
		}

		/// <summary>
		/// Inits the airway.
		/// </summary>
		/// <param name='obj'>
		/// Объект инициализируемого класса. Передаётся в конструкторе как this;
		/// </param>
		/// <param name='args'>
		/// Словарь параметров:
		/// cross_section		- сечение в мм2;
		/// material					- материал жилы. Должен принимать значения "al" или "cu". Обязательный.
		/// lenght						- длина в метрах;
		/// a								- расстояние между проводниками, м;
		/// Опционально можно передавать справочные значение именованных параметров
		/// (resistance, reactance, zero_resistance, zero_reactance)
		/// </param>
		public void InitAirway (dynamic obj, Dictionary<string,string> args)
		{
//			Dictionary<string,double> result = new Dictionary<string, double> ();
			double resistance, reactance, zero_resistance, zero_reactance;
			double lenght, cross_section, a;
            //try {
            //    double.TryParse (args ["lenght"], out lenght);
            //    double.TryParse (args ["cross_section"], out cross_section);
            //} catch {
            //    throw new InvalidInitArgument ();
            //}
            ValidInit(args["lenght"], out lenght);
            ValidInit("cross_section", out cross_section);
			if (args.ContainsKey ("resistance")) {
                //double.TryParse (args ["resistance"], out resistance);
                ValidInit(args["resistance"], out resistance);
			} else {
				resistance = Ro[args["material"]] * lenght / cross_section;
			}
//			result.Add("resistance", resistance * lenght);
			if (args.ContainsKey ("reactance")) {
                //double.TryParse (args ["reactance"], out reactance);
                ValidInit(args["reactance"], out reactance);
			} else {
                //try
                //{
                //    double.TryParse(args["a"], out a);
                //}catch{
                //    throw new InvalidInitArgument();
                //}
                ValidInit(args["a"], out a);
				reactance = 0.000145 * Math.Log10(1000 * a / (Math.Sqrt(cross_section / Math.PI)));
			}
//			result.Add("reactance", reactance * lenght);
			if (args.ContainsKey ("zero_resistance")) {
                //double.TryParse (args ["zero_resistance"], out zero_resistance);
                ValidInit(args["zero_resistance"], out zero_resistance);
			} else {
				zero_resistance = 0.00015 * resistance;
			}
//			result.Add("zero_resistance", zero_resistance * lenght);
			if (args.ContainsKey ("zero_reactance")) {
				double.TryParse (args ["zero_reactance"], out zero_reactance);
			} else {
				zero_reactance = AirwayX0X1[args["type"]] * reactance;
			}
//			result.Add("zero_reactance", zero_reactance * lenght);
//			return result;
			Initialize(obj, resistance * lenght, reactance * lenght, zero_resistance * lenght, zero_reactance * lenght);
		}

		/// <summary>
		/// Inits the transformer.
		/// </summary>
		/// <param name='obj'>
		/// Объект инициализируемого класса. Передаётся в конструкторе как this;
		/// </param>
		/// <param name='args'>
		/// Словарь параметров:
		/// voltage					- нижнее напряжение, В;
		/// power						- номинальная мощность, ВА;
		/// Pk							- потери КЗ, Вт:
		/// uk							- напряжение КЗ, %;
		/// </param>
		public void InitTransformer (dynamic obj, Dictionary<string,string> args)
		{
			double power, voltage, Pk, uk;
			double resistance, reactance, zero_resistance, zero_reactance;
            //try {
            //    double.TryParse (args ["voltage"], out voltage);
            //    double.TryParse (args ["power"], out power);
            //    double.TryParse (args ["Pk"], out Pk);
            //    double.TryParse (args ["uk"], out uk);
            //} catch {
            //    throw new InvalidInitArgument ();
            //}
            ValidInit(args["voltage"], out voltage);
            ValidInit(args["power"], out power);
            ValidInit(args["Pk"], out Pk);
            ValidInit(args["uk"], out uk);
			resistance = Pk * Math.Pow ((voltage / power), 2);
			reactance = (Math.Pow (voltage, 2) / (100 * power)) * Math.Sqrt (Math.Pow (uk, 2) - Math.Pow ((100 * Pk / power), 2));
			if (args.ContainsKey ("zero_resistance")) {
                //try {
                //    double.TryParse (args ["zero_resistance"], out zero_resistance);
                //} catch {
                //    throw new InvalidInitArgument ();
                //}
                ValidInit(args["zero_resistance"], out zero_resistance);
			} else {
				if (args.ContainsKey("scheme") && args["scheme"]=="DD"){
					zero_resistance = 3 * resistance;
				}else{
					zero_resistance = resistance;
				}
			}
			if (args.ContainsKey ("zero_reactance")) {
                //try {
                //    double.TryParse (args ["zero_reactance"], out zero_reactance);
                //} catch {
                //    throw new InvalidInitArgument ();
                //}
                ValidInit(args["zero_reactance"], out zero_reactance);
			} else {
				if (args.ContainsKey("scheme") && args["scheme"]=="DD"){
					zero_reactance = 3 * reactance;
				}else{
					zero_reactance = reactance;
				}
			}
			//return new Dictionary<string, double>(){{"resistance", resistance}, {"reactance", reactance}, {"zero_resistance", zero_resistance}, {"zero_reactance", zero_reactance}};
			Initialize(obj, resistance, reactance, zero_resistance, zero_reactance);
		}

		/// <summary>
		/// Inits the system.
		/// </summary>
		/// <param name='obj'>
		/// Объект инициализируемого класса. Передаётся в конструкторе как this;
		/// </param>
		/// <param name='args'>
		/// Словарь параметров:
		/// power					- мощность КЗ, ВА;
		/// amperage			- ток КЗ, А;
		/// highvoltage, lowvoltage - высшее и низжее напряжение, В;
		/// reactance			- именованный параметр, Ом;
		/// </param>
		public void InitSystem (dynamic obj, Dictionary<string,string> args)
		{
			//Dictionary<string,double> result = new Dictionary<string, double> (){{"resistance",0}, {"zero_resistance", 0}, {"zero_reactance", 0}};
			double reactance = 0;
			if (args.ContainsKey ("reactance")) {
                //try {
                //    double.TryParse (args ["reactance"], out reactance);
                //} catch {
                //    throw new InvalidInitArgument ();
                //}
                ValidInit(args["reactance"], out reactance);
			}
			if (args.ContainsKey ("power") && args.ContainsKey ("lowvoltage")) {
				double power, lowvoltage;
                //try {
                //    double.TryParse (args ["power"], out power);
                //    double.TryParse (args ["lowvoltage"], out lowvoltage);
                //} catch {
                //    throw new InvalidInitArgument ();
                //}
                ValidInit(args["power"], out power);
                ValidInit(args["lowvoltage"], out lowvoltage);
				reactance = Math.Pow (lowvoltage, 2) / power;
			}
			if (args.ContainsKey ("amperage") && args.ContainsKey ("highvoltage") && args.ContainsKey ("lowvoltage")) {
				double amperage, highvoltage, lowvoltage;
                //try{
                //    double.TryParse(args["amperage"], out amperage);
                //    double.TryParse(args["highvoltage"], out highvoltage);
                //    double.TryParse(args["lowvoltage"], out lowvoltage);
                //}catch{
                //    throw new InvalidInitArgument();
                //}
                ValidInit(args["amperage"], out amperage);
                ValidInit(args["highvoltage"], out highvoltage);
                ValidInit(args["lowvoltage"], out lowvoltage);
				reactance = Math.Pow(lowvoltage, 2) / (Math.Sqrt(3) * amperage * highvoltage);
			}
			Initialize(obj, 0, reactance, 0, 0);
			//return result;
		}

		/// <summary>
		/// Inits the reactor.
		/// </summary>
		/// <param name='obj'>
		/// Объект инициализируемого класса. Передаётся в конструкторе как this;
		/// </param>
		/// <param name='args'>
		/// Словарь параметров:
		/// amperage					- номинальный ток, А;
		/// dP								- потери в фазе, Вт;
		/// L, M								- индуктивность и взаимная индуктивность, Гн;
		/// reactance					- опциональный именованный параметр, Ом;
		/// </param>
		public void InitReactor (dynamic obj, Dictionary<string,string> args)
		{
			double resistance, reactance, zero_resistance, zero_reactance;
			double deltaP, amperage;
			if (args.ContainsKey ("dP") && args.ContainsKey ("amperage")) {
				ValidInit (args ["dP"], out deltaP);
				ValidInit (args ["amperage"], out amperage);
				resistance = deltaP / Math.Pow (amperage, 2);
			} else {
				throw new InvalidInitArgument ();
			}
			if (args.ContainsKey ("reactance")) {
				ValidInit (args ["reactance"], out reactance);
			} else {
				if (args.ContainsKey("L") && args.ContainsKey("M")){
					double L, M;
					ValidInit(args["L"], out L);
					ValidInit(args["M"], out M);
					reactance	= 100 * Math.PI * (L - M);
				}else{
					throw new InvalidInitArgument();
				}
			}
			zero_reactance = reactance;
			zero_resistance = resistance;
			Initialize(obj, resistance, reactance, zero_resistance, zero_reactance);
		}
	}
}

