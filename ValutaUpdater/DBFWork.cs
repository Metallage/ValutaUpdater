using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Odbc;
using System.Data;
using System.Data.OleDb;
using System.Globalization;

namespace ValutaUpdater
{
    class DBFWork
    {
        private string dbfFilePath;
        private OdbcConnection conDBF = null;
        //private string[] drivers = new string[2];  


        public DBFWork(string dbfFilePath)
        {

            this.conDBF = new OdbcConnection();
            conDBF.ConnectionString = @"Driver={Microsoft Access dBase Driver (*.dbf, *.ndx, *.mdx)}; datasource=dBase Files;";
            this.dbfFilePath = dbfFilePath;
        }

        public DBFWork(string dbfFilePath, int year)
        {
            this.conDBF = new OdbcConnection();
            conDBF.ConnectionString = @"Driver={Microsoft Access dBase Driver (*.dbf, *.ndx, *.mdx)}; datasource=dBase Files;";
            this.dbfFilePath = dbfFilePath;
        }


        public void InsertNew(long kol, string buk, double okurs, string kod, DateTime data)
        {

            conDBF.Open();

            OdbcCommand dbfCommand = conDBF.CreateCommand();
            dbfCommand.CommandText = $"INSERT INTO {dbfFilePath} VALUES ('{kod}', '{buk.ToUpper()}', {kol}, {okurs.ToString("F4", CultureInfo.GetCultureInfo("en-US"))}, 0, {data.Date.ToOADate()});";
            dbfCommand.ExecuteNonQuery();
            conDBF.Close();

        }


        /// <summary>
        /// Выбирает данные за определённый период
        /// </summary>
        /// <param name="dateFrom">Начало периода</param>
        /// <param name="dateTo">Конец периода</param>
        /// <returns>Курсы валют</returns>
        public DataTable ReadbyDate(DateTime dateFrom, DateTime dateTo)
        {
            DataTable resultTable = new DataTable();
            long minDate = (long)dateFrom.Date.ToOADate();
            long maxDate = (long)dateTo.Date.ToOADate();

            conDBF.Open();

            OdbcCommand dbfCommand = conDBF.CreateCommand();
            dbfCommand.CommandText = $"SELECT * FROM {dbfFilePath} as V2 WHERE V2.DATA >= {minDate} AND V2.DATA <= {maxDate}; ";
            resultTable.Load(dbfCommand.ExecuteReader());
            conDBF.Close();

            return resultTable;
        }

        /// <summary>
        /// Проверяет имеет ли смысл что либо искать в таблице в интересующий нас период
        /// </summary>
        /// <param name="periodFrom">Дата с которой начинается интересующий нас период</param>
        /// <param name="periodTo">Дата которой оканчивается интересующий нас период</param>
        /// <returns>Есть ли какие-либо записи</returns>
        public bool CheckCount(DateTime periodFrom, DateTime periodTo)
        {
            bool hasRows = false;

            //Переводим даты начала и конца в понятный для DBF формат
            long minDate = (long)periodFrom.Date.ToOADate();
            long maxDate = (long)periodTo.Date.ToOADate();

            //Таблица хранящая число записей
            DataTable rowsCount = new DataTable();

            //Выполняем SQL запрос на подсчёт строк в итересующий нас период
            conDBF.Open();
            OdbcCommand dbfCommand = conDBF.CreateCommand();
            dbfCommand.CommandText = $"SELECT COUNT(*) FROM {dbfFilePath} as V2 WHERE V2.DATA >= {minDate} AND V2.DATA <= {maxDate}; ";
            rowsCount.Load(dbfCommand.ExecuteReader());
            conDBF.Close();

            //Если значение больше 0, значит в этот период что-то есть
            if(rowsCount.Rows[0].Field<int>(0) > 0)
            {
                hasRows = true;
            }

            return hasRows;
        }

        /// <summary>
        /// Выбираем все уникальные даты курсов валют за месяц
        /// </summary>
        /// <param name="year">Год</param>
        /// <param name="month">Месяц</param>
        /// <returns>Таблица дат</returns>
        public DataTable SelectDatesFromMonth(int year, int month)
        {
            DataTable dates = new DataTable();

            //Определяем границы дат
            DateTime firstDay = new DateTime(year, month, 1);
            DateTime lastDay = new DateTime(year, month, DateTime.DaysInMonth(year,month));

            //Переводим границы дат в понятный для DBF вид
            long firstDayOA = (long)firstDay.ToOADate();
            long lastDayOA = (long)lastDay.ToOADate();

            //Выполняем SQL запрос
            conDBF.Open();
            OdbcCommand dbfCommand = conDBF.CreateCommand();
            dbfCommand.CommandText = $"SELECT DISTINCT V2.DATA FROM {dbfFilePath} as V2 WHERE V2.DATA >= {firstDayOA} AND V2.DATA <= {lastDayOA}; ";
            dates.Load(dbfCommand.ExecuteReader());
            conDBF.Close();

            return dates;
        }

        /// <summary>
        /// Выбирает данные по курсам валют на определённую дату
        /// </summary>
        /// <param name="date">Дата</param>
        /// <returns>Таблица валют</returns>
        public DataTable SelectByDate(DateTime date)
        {
            DataTable valuta = new DataTable();

            //Преобразуем дату в понятный для DBF таблицы формат 
            long dateOA = (long)date.ToOADate();

            //Выполняем SQL запрос
            conDBF.Open();
            OdbcCommand dbfCommand = conDBF.CreateCommand();
            dbfCommand.CommandText = $"SELECT KOL, BUK, OKURS, KOD FROM {dbfFilePath} as V2 WHERE V2.DATA = {dateOA}; ";
            valuta.Load(dbfCommand.ExecuteReader());
            conDBF.Close();

            return valuta;
        }


    }
}
