using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace OurDatabase
{
    public partial class Form5 : Form
    {
        public static string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb;"; //строка подключения к базе данных
        private OleDbConnection myConnection; //подключение
        private OleDbDataAdapter dataadapter; //адаптер для таблицы
        List<int> codes_of_country = new List<int>(); //коды стран, городов, континентов и певцов. По ним идет обращение к базе
        List<int> codes_of_city = new List<int>();
        List<int> codes_of_singer = new List<int>();
        List<int> codes_of_continents = new List<int>();

        int nContinent = 0; // номер выбранного города, страны, континента, певца. Формирует запрос where
        int nCountry = 0;
        int nCity = 0;
        int nGroup = 0;
        string answer;
        bool from_city; // флаги, показывающие, на какой фильтр было нажатие
        bool from_country;
        bool from_continent;
        bool from_singer;

        bool from_continent_go; //если ставим фильтры вперед
        bool from_country_go;
        bool from_city_go;
        int difference = 0;
        int difference1 = 0;
        int difference2 = 0;

       
       

        public Form5()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();

        }

        private string Function(int nGroup, int nCity, int nCountry, int nContinent)
        {
            //функция, формирующая запрос WHERE для установки фильтра
            string sw = ""; //нач строка
            if (nGroup == 0)
            {
                sw = " WHERE [SingerGroup].[Code] >0 "; //для правильного создания запроса, если певец не выбран
            }
            else
            {
                sw = " WHERE [SingerGroup].[Code] = " + nGroup.ToString(); // вывод тех полей, где номер певца = номеру выбранного 
            }

            if (nCity > 0) //если выбран еще город
            {
                sw += " AND [City].[Code] = " + nCity.ToString();  // те поля, где город=выбранный город
            }
            if (nCountry > 0)
            {
                sw += " AND [Country].[Code] = " + nCountry.ToString();
            }
            if (nContinent > 0)
            {
                sw += " AND [Continent].[Code] = " + nContinent.ToString();
            }
            return sw;
        }

        private void Count_title(List<int> codes_of_singer, int count)
        {
            int c = 0;
            int cc = 0;
            for (int i = 0; i < count; i++)
            {
                string count_songs = "SELECT COUNT([Title]) FROM [Single] WHERE [Single].[NumSingerGroup] = " + codes_of_singer[i].ToString();
                OleDbCommand songs = new OleDbCommand(count_songs, myConnection);
                OleDbDataReader read_songs = songs.ExecuteReader();
                while (read_songs.Read())
                {
                    c = Convert.ToInt32(read_songs[0]);

                }
                read_songs.Close();
                cc += c;
            }

            label5.Text = cc.ToString();
        }


        private void button4_Click(object sender, EventArgs e)
        {
            if (from_singer == true) //если выбран певец(неважно,в какой последовательности-по два, три, один или все), то считаем лишь песни
            {
                Count_title(codes_of_singer, MySinger.Items.Count);
                label5.Text = "4";
                label6.Text = "1";
                label7.Text = "1";
                label8.Text = "1";
                label9.Text = "1";

            }
            else //иначе
            {
                if (from_city == true) //если выбран город,то страна и континент - 1. Считаем остальное по чилу оставшегося в фильтрах
                {
                    label7.Text = "1";
                    label8.Text = "1";
                    label9.Text = "1";
                    label6.Text = MySinger.Items.Count.ToString();
                    Count_title(codes_of_singer, MySinger.Items.Count);
                }
                else if (from_country == true) // то же со страной
                {
                    label8.Text = "1";
                    label9.Text = "1";
                    label7.Text = MyCity.Items.Count.ToString();
                    label6.Text = MySinger.Items.Count.ToString();
                    Count_title(codes_of_singer, MySinger.Items.Count);
                }
                else if (from_continent == true) // то же с континентом
                {
                    label8.Text = comboBox3.Items.Count.ToString();
                    label9.Text = "1";
                    label7.Text = MyCity.Items.Count.ToString();
                    label6.Text = MySinger.Items.Count.ToString();
                    Count_title(codes_of_singer, MySinger.Items.Count);
                }
            }

            string connectData = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb;"; //подключение
            string sql = "SELECT [Title] as Название,[ArtistName] as Группа,[City] as Город, " +
                "[CountryName] as Страна,[Continent] as Континент"; //часть главного запроса для вывода полей

            string tables = " FROM [Single],[SingerGroup],[City],[Country],[Continent] "; //с каких таблиц берем данные

            string sw = Function(nGroup, nCity, nCountry, nContinent); //получаем условие WHERE из фильтров
            string origin = " AND [SingerGroup].[Code] = [Single].[NumSingerGroup] AND [City].[Code] = [SingerGroup].[NumCity] AND [Country].[Code] = [City].[NumCountry]" +
                " AND [Country].[NumContinent] = [Continent].[Code]"; //соединение таблиц для лииквидации дублирования
            sw += origin;
            sw += " ORDER BY [Title]"; //сортировка

            sql += tables + sw;
            myConnection = new OleDbConnection(connectData);

            dataadapter = new OleDbDataAdapter(sql, myConnection);

            myConnection.Open(); //создание таблицы 
            DataSet ds = new DataSet();
            dataadapter.Fill(ds);
            table.DataSource = ds.Tables[0];
            table.ReadOnly = true;

        }

        private void Form5_Load(object sender, EventArgs e)
        {
            //запрос соединения таблиц 
            string connectData = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb;";
            string sql = "SELECT [Title] as Название,[ArtistName] as Группа,[City] as Город,[CountryName] as Страна,[Continent] as Континент" +
               " FROM [Single],[SingerGroup],[City],[Country],[Continent]" +
               " WHERE [SingerGroup].[Code] = [Single].[NumSingerGroup] AND [City].[Code] = [SingerGroup].[NumCity] AND [Country].[Code] = [City].[NumCountry]" +
               " AND [Country].[NumContinent] = [Continent].[Code]";

            myConnection = new OleDbConnection(connectData);
            dataadapter = new OleDbDataAdapter(sql, myConnection);
            myConnection.Open();
            DataSet ds = new DataSet();
            dataadapter.Fill(ds);
            table.DataSource = ds.Tables[0];
            table.ReadOnly = true;

            myConnection = new OleDbConnection(connectString);
            myConnection.Open();
            string ls1 = "SELECT [Continent],[Code] FROM [Continent] ORDER BY [Code]";
            OleDbCommand continent = new OleDbCommand(ls1, myConnection);
            OleDbDataReader que_reader_c = continent.ExecuteReader();
            while (que_reader_c.Read())
            {
                MyContinent.Items.Add(que_reader_c[0].ToString());
                codes_of_continents.Add(Convert.ToInt32(que_reader_c[1]));
            }
            que_reader_c.Close();

            string ls2 = "SELECT [ArtistName],[Code] FROM [SingerGroup] ORDER BY [Code]";
            OleDbCommand all = new OleDbCommand(ls2, myConnection);
            OleDbDataReader allr = all.ExecuteReader();
            while (allr.Read())
            {
                MySinger.Items.Add(allr[0].ToString());
                codes_of_singer.Add(Convert.ToInt32(allr[1]));
            }
            allr.Close();

            string city_list = "SELECT [City],[Code] FROM [City]  ORDER BY [Code]";
            OleDbCommand city = new OleDbCommand(city_list, myConnection);
            OleDbDataReader read_city = city.ExecuteReader();
            while (read_city.Read())
            {
                MyCity.Items.Add(read_city[0].ToString());
                codes_of_city.Add(Convert.ToInt32(read_city[1]));

            }
            read_city.Close();


            string list_country = "SELECT [CountryName],[Code] FROM [Country]  ORDER BY [Code] ";
            OleDbCommand country = new OleDbCommand(list_country, myConnection);

            OleDbDataReader read_country = country.ExecuteReader();
            while (read_country.Read())
            {
                comboBox3.Items.Add(read_country[0].ToString());
                codes_of_country.Add(Convert.ToInt32(read_country[1]));

            }
            read_country.Close();

            string count_songs = "SELECT COUNT([Title]) FROM [Single] ";
            OleDbCommand songs = new OleDbCommand(count_songs, myConnection);
            OleDbDataReader read_songs = songs.ExecuteReader();
            while (read_songs.Read())
            {
                label5.Text = read_songs[0].ToString();

            }
            read_songs.Close();

            label6.Text = MySinger.Items.Count.ToString();
            label7.Text = MyCity.Items.Count.ToString();
            label8.Text = comboBox3.Items.Count.ToString();
            label9.Text = MyContinent.Items.Count.ToString();
            button4.Enabled = true;
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            from_country = true;
            answer = comboBox3.Text;
            nCountry = codes_of_country[comboBox3.SelectedIndex];
            difference++;

            if (from_city == true && from_singer == true) // если берем певца, затем город
            {

                MyContinent.Items.Clear();
                MyContinent.Text = "";
                codes_of_continents.Clear();
            }
            else if (from_continent_go == true)
            {
                from_country_go = true;
            }
            else if (from_continent == false && from_city == false && from_singer == false)
            {
                MyContinent.Items.Clear();
                MyContinent.Text = "";
                codes_of_continents.Clear();

                MyCity.Items.Clear();
                MyCity.Text = "";
                codes_of_city.Clear();

                string ls3 = "SELECT [City],[Code] FROM [City] WHERE [City].[NumCountry] = " + nCountry.ToString();
                OleDbCommand orig1 = new OleDbCommand(ls3, myConnection);
                OleDbDataReader read_or = orig1.ExecuteReader();
                while (read_or.Read())
                {
                    MyCity.Items.Add(read_or[0].ToString());
                    codes_of_city.Add(Convert.ToInt32(read_or[1]));

                }
                read_or.Close();

                MySinger.Items.Clear();
                codes_of_singer.Clear();
                MySinger.Text = "";

                for (int i = 0; i < codes_of_city.Count; i++)
                {
                    string ls34 = "SELECT [ArtistName],[Code] FROM [SingerGroup] WHERE [SingerGroup].[NumCity]=" + codes_of_city[i].ToString();
                    OleDbCommand orig14 = new OleDbCommand(ls34, myConnection);
                    OleDbDataReader read_or4 = orig14.ExecuteReader();
                    while (read_or4.Read())
                    {
                        MySinger.Items.Add(read_or4[0].ToString());
                        codes_of_singer.Add(Convert.ToInt32(read_or4[1]));

                    }
                    read_or4.Close();
                }


            }
            else if (from_continent == false && from_singer == false && from_city == true)
            {
                MyContinent.Items.Clear();
                MyContinent.Text = "";
                codes_of_continents.Clear();

            }

            if (from_continent_go == true && from_city == false && from_singer == false)
            {
                MyCity.Items.Clear();
                MyCity.Text = "";
                codes_of_city.Clear();

                string ls1 = "SELECT [City],[Code] FROM [City] WHERE [NumCountry]=@p";
                OleDbCommand continent = new OleDbCommand(ls1, myConnection);
                continent.Parameters.AddWithValue("@p", nCountry);
                OleDbDataReader que_reader_c = continent.ExecuteReader();
                while (que_reader_c.Read())
                {
                    MyCity.Items.Add(que_reader_c[0].ToString());
                    codes_of_city.Add(Convert.ToInt32(que_reader_c[1]));

                }
                que_reader_c.Close();
                from_country_go = true;
            }

            if (difference1 > 0)
            {
                int code_continent = 0;
                string ls1 = "SELECT [NumContinent] FROM [Country] WHERE [Country].[Code] = " + nCountry.ToString();
                OleDbCommand country = new OleDbCommand(ls1, myConnection);
                OleDbDataReader read_country = country.ExecuteReader();
                while (read_country.Read())
                {

                    code_continent = Convert.ToInt32(read_country[0]);

                }
                read_country.Close();

                string ls2 = "SELECT [Continent],[Code] FROM [Continent] WHERE [Continent].[Code] = " + code_continent.ToString();
                OleDbCommand continent = new OleDbCommand(ls2, myConnection);
                OleDbDataReader read_continent = continent.ExecuteReader();
                while (read_continent.Read())
                {
                    MyContinent.Items.Add(read_continent[0].ToString());
                    codes_of_continents.Add(Convert.ToInt32(read_continent[1]));

                }
                read_continent.Close();
            }
            else
            {
                int code_continent = 0;
                string ls1 = "SELECT [NumContinent] FROM [Country] WHERE [Country].[Code] = " + nCountry.ToString();
                OleDbCommand country = new OleDbCommand(ls1, myConnection);
                OleDbDataReader read_country = country.ExecuteReader();
                while (read_country.Read())
                {

                    code_continent = Convert.ToInt32(read_country[0]);

                }
                read_country.Close();

                string ls2 = "SELECT [Continent],[Code] FROM [Continent] WHERE [Continent].[Code] = " + code_continent.ToString();
                OleDbCommand continent = new OleDbCommand(ls2, myConnection);
                OleDbDataReader read_continent = continent.ExecuteReader();
                while (read_continent.Read())
                {
                    MyContinent.Items.Add(read_continent[0].ToString());
                    codes_of_continents.Add(Convert.ToInt32(read_continent[1]));

                }
                read_continent.Close();
            }
        }

        private void MyContinent_SelectedIndexChanged(object sender, EventArgs e)
        {
            nContinent = codes_of_continents[MyContinent.SelectedIndex];
            from_continent = true;
            from_continent_go = true;
            difference++;

            if (from_country == false && from_city == false && from_singer == false)
            {

                comboBox3.Items.Clear();
                comboBox3.Text = "";
                codes_of_country.Clear();

                MySinger.Items.Clear();
                codes_of_singer.Clear();
                MySinger.Text = "";

                codes_of_city.Clear();
                MyCity.Items.Clear();
                MyCity.Text = "";


                string ls1 = "SELECT [CountryName],[Code] FROM [Country] WHERE [NumContinent]=@p";
                OleDbCommand continent = new OleDbCommand(ls1, myConnection);
                continent.Parameters.AddWithValue("@p", nContinent);
                OleDbDataReader que_reader_c = continent.ExecuteReader();
                while (que_reader_c.Read())
                {
                    comboBox3.Items.Add(que_reader_c[0].ToString());
                    codes_of_country.Add(Convert.ToInt32(que_reader_c[1]));

                }
                que_reader_c.Close();


                for (int i = 0; i < codes_of_country.Count; i++)
                {
                    string ls12 = "SELECT [City],[Code] FROM [City] WHERE [NumCountry]=@p";
                    OleDbCommand continent2 = new OleDbCommand(ls12, myConnection);
                    continent2.Parameters.AddWithValue("@p", codes_of_country[i]);
                    OleDbDataReader que_reader_c2 = continent2.ExecuteReader();
                    while (que_reader_c2.Read())
                    {
                        MyCity.Items.Add(que_reader_c2[0].ToString());
                        codes_of_city.Add(Convert.ToInt32(que_reader_c2[1]));

                    }
                    que_reader_c2.Close();
                }

                for (int j = 0; j < codes_of_city.Count; j++)
                {
                    string ls123 = "SELECT [ArtistName],[Code] FROM [SingerGroup] WHERE [NumCity]=@p";
                    OleDbCommand continent23 = new OleDbCommand(ls123, myConnection);
                    continent23.Parameters.AddWithValue("@p", codes_of_city[j]);
                    OleDbDataReader que_reader_c23 = continent23.ExecuteReader();
                    while (que_reader_c23.Read())
                    {
                        MySinger.Items.Add(que_reader_c23[0].ToString());
                        codes_of_singer.Add(Convert.ToInt32(que_reader_c23[1]));

                    }
                    que_reader_c23.Close();
                }
            }
            else if (from_country == true && from_city == false && from_singer == false)
            {

            }


        }

        private void MyCity_SelectedIndexChanged(object sender, EventArgs e)
        {
            nCity = codes_of_city[MyCity.SelectedIndex];
            from_city = true;
            difference1++;

            int code_coutry = 0;
            int code_continent = 0;
            difference2++;

            if (from_singer == true && from_city == false && from_country == false && from_continent == false) // нажимаем вначале певцов
            {
                comboBox3.Items.Clear();
                comboBox3.Text = "";
                MyContinent.Items.Clear();
                MyContinent.Text = "";
                codes_of_country.Clear();
                codes_of_continents.Clear();


            }
            else if (difference2 > 0 && from_singer == false) // вначале нажимаем страны
            {
                MySinger.Items.Clear();
                codes_of_singer.Clear();
                MySinger.Text = "";
                from_city = true;
                string ls01 = "SELECT [ArtistName],[Code] FROM [SingerGroup] WHERE [SingerGroup].[NumCity] = " + nCity.ToString();
                OleDbCommand aa1 = new OleDbCommand(ls01, myConnection);
                OleDbDataReader read_aa1 = aa1.ExecuteReader();
                while (read_aa1.Read())
                {
                    MySinger.Items.Add(read_aa1[0].ToString());
                    codes_of_singer.Add(Convert.ToInt32(read_aa1[1]));
                }
                read_aa1.Close();
            }
            else if (from_country == true && from_singer == false && from_city == true && from_continent == true && difference > 0) // вначале нажимаем страны, затем континент
            {
                comboBox3.Items.Clear();
                comboBox3.Text = "";
                MyContinent.Items.Clear();
                MyContinent.Text = "";
                codes_of_country.Clear();
                codes_of_continents.Clear();
            }

            else if (from_singer == false && from_country == false && from_city == true && from_continent == false && difference == 0) // нажимаем вначале города
            {
                comboBox3.Items.Clear();
                comboBox3.Text = "";
                MyContinent.Items.Clear();
                MyContinent.Text = "";
                codes_of_country.Clear();
                codes_of_continents.Clear();

                MySinger.Items.Clear();
                codes_of_singer.Clear();
                MySinger.Text = "";

                string ls01 = "SELECT [ArtistName],[Code] FROM [SingerGroup] WHERE [SingerGroup].[NumCity] = " + nCity.ToString();
                OleDbCommand aa1 = new OleDbCommand(ls01, myConnection);
                OleDbDataReader read_aa1 = aa1.ExecuteReader();
                while (read_aa1.Read())
                {
                    MySinger.Items.Add(read_aa1[0].ToString());
                    codes_of_singer.Add(Convert.ToInt32(read_aa1[1]));
                }
                read_aa1.Close();
            }
            if (from_continent_go == false && from_country_go == false)
            {
                comboBox3.Items.Clear();
                comboBox3.Text = "";
                MyContinent.Items.Clear();
                MyContinent.Text = "";
                codes_of_country.Clear();
                codes_of_continents.Clear();

            }
            else
            {

            }

            string ls0 = "SELECT [NumCountry] FROM [City] WHERE [City].[Code] = " + nCity.ToString();
            OleDbCommand aa = new OleDbCommand(ls0, myConnection);
            OleDbDataReader read_aa = aa.ExecuteReader();
            while (read_aa.Read())
            {
                code_coutry = Convert.ToInt32(read_aa[0]);
            }
            read_aa.Close();

            if (from_continent != true && from_city == true && from_continent_go != true)
            {
                string ls1 = "SELECT [CountryName],[NumContinent],[Code] FROM [Country] WHERE [Country].[Code] = " + code_coutry.ToString();
                OleDbCommand country = new OleDbCommand(ls1, myConnection);
                OleDbDataReader read_country = country.ExecuteReader();
                while (read_country.Read())
                {
                    comboBox3.Items.Add(read_country[0].ToString());
                    codes_of_country.Add(Convert.ToInt32(read_country[2]));
                    code_continent = Convert.ToInt32(read_country[1]);
                }
                read_country.Close();
            }
            else
            {

            }

            string ls2 = "SELECT [Continent],[Code] FROM [Continent] WHERE [Continent].[Code] = " + code_continent.ToString();
            OleDbCommand continent = new OleDbCommand(ls2, myConnection);
            OleDbDataReader read_continent = continent.ExecuteReader();
            while (read_continent.Read())
            {
                MyContinent.Items.Add(read_continent[0].ToString());
                codes_of_continents.Add(Convert.ToInt32(read_continent[1]));

            }
            read_continent.Close();
            comboBox3.Text = answer;
        }

        private void table_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }


        private void MySinger_Click(object sender, EventArgs e)
        {

        }

        private void MySinger_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            int code_coutry = 0;
            int code_continent = 0;
            int code_city = 0;
            from_singer = true;
            nGroup = codes_of_singer[MySinger.SelectedIndex];

            if (from_city == false && from_continent == false && from_country == false) //еали вначале нажимаем певцов
            {
                codes_of_country.Clear();
                comboBox3.Items.Clear();
                comboBox3.Text = "";

                codes_of_continents.Clear();
                MyContinent.Items.Clear();
                MyContinent.Text = "";

                codes_of_city.Clear();
                MyCity.Items.Clear();
                MyCity.Text = "";
                string ls = "SELECT [NumCity] FROM [SingerGroup] WHERE [SingerGroup].[Code] = " + nGroup.ToString();
                OleDbCommand aaa = new OleDbCommand(ls, myConnection);
                OleDbDataReader read_aaa = aaa.ExecuteReader();
                while (read_aaa.Read())
                {

                    code_city = Convert.ToInt32(read_aaa[0]);
                }
                read_aaa.Close();

                string ff = "SELECT [City],[Code] FROM [City] WHERE [City].[Code] = " + code_city.ToString();
                OleDbCommand fa = new OleDbCommand(ff, myConnection);
                OleDbDataReader read_ff = fa.ExecuteReader();
                while (read_ff.Read())
                {
                    MyCity.Items.Add(read_ff[0].ToString());
                    codes_of_city.Add(Convert.ToInt32(read_ff[1]));

                }
                read_ff.Close();

                string ls0 = "SELECT [NumCountry] FROM [City] WHERE [City].[Code] = " + code_city.ToString();
                OleDbCommand aa = new OleDbCommand(ls0, myConnection);
                OleDbDataReader read_aa = aa.ExecuteReader();
                while (read_aa.Read())
                {
                    code_coutry = Convert.ToInt32(read_aa[0]);
                }
                read_aa.Close();

                string ls1 = "SELECT [CountryName],[NumContinent],[Code] FROM [Country] WHERE [Country].[Code] = " + code_coutry.ToString();
                OleDbCommand country = new OleDbCommand(ls1, myConnection);
                OleDbDataReader read_country = country.ExecuteReader();
                while (read_country.Read())
                {
                    comboBox3.Items.Add(read_country[0].ToString());
                    codes_of_country.Add(Convert.ToInt32(read_country[2]));
                    code_continent = Convert.ToInt32(read_country[1]);
                }
                read_country.Close();

                string ls2 = "SELECT [Continent],[Code] FROM [Continent] WHERE [Continent].[Code] = " + code_continent.ToString();
                OleDbCommand continent = new OleDbCommand(ls2, myConnection);
                OleDbDataReader read_continent = continent.ExecuteReader();
                while (read_continent.Read())
                {
                    MyContinent.Items.Add(read_continent[0].ToString());
                    codes_of_continents.Add(Convert.ToInt32(read_continent[1]));

                }
                read_continent.Close();
            }
            else if (from_city == false && from_country == false && from_continent == false)
            {

            }
            button4.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form5_Load(this, null);
            codes_of_continents.Clear();
            codes_of_city.Clear();
            codes_of_singer.Clear();
            codes_of_country.Clear();
            MySinger.Items.Clear();
            MyContinent.Items.Clear();
            comboBox3.Items.Clear();
            MyCity.Items.Clear();

            MySinger.Text = "";
            MyContinent.Text = "";
            comboBox3.Text = "";
            MyCity.Text = "";

            difference = 0;
            difference1 = 0;
            difference2 = 0;

            from_city = false;
            from_country = false;
            from_continent = false;
            from_singer = false;

            from_continent_go = false;
            from_country_go = false;
            from_city_go = false;

            nContinent = 0;
            nCountry = 0;
            nCity = 0;
            nGroup = 0;
            answer = "";


            myConnection = new OleDbConnection(connectString);
            myConnection.Open();

            string ls1 = "SELECT [Continent],[Code] FROM [Continent] ORDER BY [Code]";
            OleDbCommand continent = new OleDbCommand(ls1, myConnection);
            OleDbDataReader que_reader_c = continent.ExecuteReader();
            while (que_reader_c.Read())
            {
                MyContinent.Items.Add(que_reader_c[0].ToString());
                codes_of_continents.Add(Convert.ToInt32(que_reader_c[1]));
            }
            que_reader_c.Close();

            string ls2 = "SELECT [ArtistName],[Code] FROM [SingerGroup] ORDER BY [Code]";
            OleDbCommand all = new OleDbCommand(ls2, myConnection);
            OleDbDataReader allr = all.ExecuteReader();
            while (allr.Read())
            {
                MySinger.Items.Add(allr[0].ToString());
                codes_of_singer.Add(Convert.ToInt32(allr[1]));
            }
            allr.Close();

            string city_list = "SELECT [City],[Code] FROM [City]  ORDER BY [Code]";
            OleDbCommand city = new OleDbCommand(city_list, myConnection);
            OleDbDataReader read_city = city.ExecuteReader();
            while (read_city.Read())
            {
                MyCity.Items.Add(read_city[0].ToString());
                codes_of_city.Add(Convert.ToInt32(read_city[1]));

            }
            read_city.Close();


            string list_country = "SELECT [CountryName],[Code] FROM [Country]  ORDER BY [Code] ";
            OleDbCommand country = new OleDbCommand(list_country, myConnection);
            OleDbDataReader read_country = country.ExecuteReader();
            while (read_country.Read())
            {
                comboBox3.Items.Add(read_country[0].ToString());
                codes_of_country.Add(Convert.ToInt32(read_country[1]));

            }
            read_country.Close();

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }

        private void PopulateRows()
        {
            for (int i = 1; i <= 10; i++)
            {
                DataGridViewRow row =
                    (DataGridViewRow)table.RowTemplate.Clone();

                row.CreateCells(table, string.Format("Song{0}", i),
                    string.Format("Singer{0}", i), string.Format("City{0}", i), string.Format("Country{0}", i), string.Format("Continent{0}", i));

                table.Rows.Add(row);

            }
        }

        private void ExportToExcel()
        {

            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            try
            {
                worksheet = workbook.ActiveSheet;
                worksheet.Name = "Songs";

                int cellRowIndex = 1;
                int cellColumnIndex = 1;


                for (int i = 0; i < table.Rows.Count - 1; i++)
                {
                    for (int j = 0; j <table.Columns.Count; j++)
                    {
                        // Excel index starts from 1,1. As first Row would have the Column headers, adding a condition check.
                        if (cellRowIndex == 1)
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = table.Columns[j].HeaderText;
                        }
                        else
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = table.Rows[i].Cells[j].Value.ToString();
                        }
                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 2;

                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    workbook.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Export Successful");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }
        }
    }
}
