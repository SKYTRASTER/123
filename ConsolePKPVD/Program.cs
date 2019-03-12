using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.ManagedDataAccess.Client;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;

namespace ConsolePKPVD
{
    class Program
    {
        static string connStringOracle;
        static OracleConnection connor = new OracleConnection();
        static string[,] list;
        static string[,] temper;

        static string datest= DateTime.Today.ToShortDateString();
        static string datefi= DateTime.Today.ToShortDateString();
        //static string datest = "11.03.2019";
        //static string datefi = "11.03.2019";
        static string[] ColumN;
        static object misValue = System.Reflection.Missing.Value;
        static string temp = "";
        static void Main(string[] args)
        {
            connector();
            Request();
            connor.Close();
            //exportexcel();
            UpdateListSMS();
            //Console.ReadKey();

        }

       
        //запрос с sql oracle
        static void Request()
        {
            try
            {
                OracleCommand orclCommand = new OracleCommand(@"SELECT distinct rb.regnum AS ""Номер заявления"", 
                CASE sp.id_operation WHEN 'IssuedDocumentPrinting' THEN 'печать документов'   
                                             WHEN 'MDDataWaiting' THEN ' ожидание данных из ГКН'
                                             WHEN 'PaymentWaiting' THEN ' ожидание платежа' 
                                             WHEN 'DocumentIssue' THEN ' выдача документов' 
                ELSE ' Иное' END AS ""Технологическая операция"",
                CASE cs.state WHEN 0 THEN 'в работе' 
                                WHEN 2 THEN 'приостановлено'
                                WHEN 1 THEN 'исполнено' 
                ELSE ' иное' END AS ""Состояние дела"", 
                CASE cs.statusmd WHEN '001' THEN 'Сформировано' 
                                     WHEN '002' THEN 'Загружено' 
                                     WHEN '003' THEN 'Отказано в загрузке' 
                                     WHEN '004' THEN 'Принято в работу' 
                                     WHEN '005' THEN 'Ошибка_маршрутизации' 
                                     WHEN '006' THEN 'Приостановлено' 
                                     WHEN '007' THEN 'Завершено' 
                                     WHEN '008' THEN 'Завершено отказом' 
                                     WHEN '009' THEN 'Отказано в обработке' 
                                     WHEN '010' THEN 'Приостановление снято' 
                                     WHEN '011' THEN 'Ожидание ГКУ' 
                                     WHEN '012' THEN 'Сведения отсутствуют' 
                                     WHEN '014' THEN 'Возврат без рассмотрения' 
                 ELSE ' иное' END AS ""Статус МД дела"", 
                 ' ' || dd.name AS ""Исходящий документ"", 
                 fd.num AS "" Номер"",
                 ' ' || TO_CHAR (rb.regdate, 'DD.MM.YYYY HH24:MI:ss') AS ""Дата регистрации"", 
                 ' ' || TO_CHAR (fd.createdate, 'DD.MM.YYYY HH24:MI:ss') AS ""Получен ПВД"", 
                 (select to_char(max(s.dateend), 'dd.mm.yyyy hh24:mi:ss')  from pvd.dps$step s
                 where  rb.id_cause = s.id_cause and s.id_operation ='IssuedDocumentPrinting')  ""Дата получения исх. док"", 
                 CASE WHEN sb.placedescript IS NULL
                 THEN  sb.postindex || ' ' || sb.region || ' ' || sb.regiontype || ', ' || sb.district || ' ' || sb.districttype || ' ' || sb.citytype || ' ' || sb.city || 
                           ' ' || sb.urbandistrict || ' ' || sb.urbandistricttype || ' ' || sb.townhall || ' ' || sb.townhalltype || ' ' || sb.localitytype || ' ' || sb.locality ||
                            ' ' || sb.streettype || ' ' || sb.street || ' ' || sb.hometype || ' ' || sb.home || ' ' || sb.buildingtype || ' ' || sb.building || ' ' || sb.constructiontype ||
                             ' ' || sb.construction || ' ' || sb.apartamenttype || ' ' || sb.apartament || ' ' || sb.other
                 ELSE sb.placedescript END AS ""Адрес заявителя"", 
                 CASE WHEN sb.phone IS NULL THEN null ELSE sb.phone END AS "" № телефона"", 
                 CASE WHEN sb.shortname IS NULL THEN(sb.surname || ' ' || sb.firstname || ' ' || sb.patronymic) 
                 ELSE sb.shortname END AS ""Заявитель"", 
                 CASE WHEN sb.email IS NULL THEN null ELSE sb.email END AS "" email""
                    FROM pvd.dps$recbookitem rb,
                       pvd.dps$curstep cr,
                       pvd.dps$step sp,
                       pvd.dps$extract xt,
                       pvd.dps$applicant pp,
                       pvd.dps$subject sb,
                       pvd.dps$formeddoc fd,
                       pvd.dps$doc dd,
                       pvd.dps$d_cause cs
                    WHERE rb.id_cause = fd.id_cause
                       AND trunc(fd.createdate) BETWEEN to_date('" + datest + @"', 'dd.mm.yyyy') AND to_date ('" + datefi + @"', 'dd.mm.yyyy')
--                     AND rb.regnum = '' 
                       AND cr.id_cause = rb.id_cause
                       AND sp.id = cr.id_step
                       AND xt.id_cause = rb.id_cause
                       AND xt.receiptplace = 5 
                       AND xt.id_applicant = pp.id
                       AND sb.id = pp.id_subject
                       AND fd.id_cause = rb.id_cause
                       AND dd.id_cause = rb.id_cause
                       AND dd.causenum = rb.regnum
                       AND fd.num = dd.num
                       AND cs.id = rb.id_cause
                       AND cs.statusmd = 007
                       AND cs.state = '0'
--                       AND sp.id_operation = 'IssuedDocumentPrinting'
               ORDER BY rb.regnum, fd.num ");
                Console.WriteLine("Запрос к SQL-базе выполнен");
                orclCommand.Connection = connor;
                OracleDataReader npgSqlDataReader = orclCommand.ExecuteReader();
                List<string> listtemp = new List<string>();
                if (npgSqlDataReader.HasRows)
                {
                    Console.WriteLine("Таблица не пустая");
                    //формируем список данных из таблицы SQL                   
                    int rows = 0;
                    while (npgSqlDataReader.Read())
                    {
                        for (int i = 0; i < npgSqlDataReader.FieldCount; i++)
                        {
                            listtemp.AddRange(new string[] { npgSqlDataReader[i].ToString() });
                        }
                        rows++;
                    }
                    Console.WriteLine("Список из из SQL сформирован");

                    //формируем двумерный массив данных из listtemp    

                    list = new string[rows, npgSqlDataReader.FieldCount];
                    int row = 0;
                    int column = 0;
                    foreach (string listik in listtemp)
                    {
                        if (column < npgSqlDataReader.FieldCount)
                        {
                            list[row, column] = listik;
                            column++;
                        }
                        else
                        {   row++;
                            column = 0;
                            list[row, column] = listik;
                            column++;
                        }
                    }
                    Console.WriteLine("Двумерный массив сформирован");
                    temper = new string[list.GetLength(0), 2];
                    for (int i = 0; i < list.GetLength(0); i++)
                    {
                        temper[i, 0] = list[i, 10];
                    }




                    ColumN = new string[list.GetLength(1)];
                    for (int i = 0; i < list.GetLength(1); i++)
                    {
                        ColumN[i] = npgSqlDataReader.GetName(i);
                    }
                }
             else { 
                    Console.WriteLine("Не найдено ни одной записи");
                    ToLogFileError(string.Concat(new object[] { DateTime.Now, String.Format(" Не найдено ни одной записи  "), Environment.NewLine }));
                    Environment.Exit(0);
                }   
            }
            catch (Exception ex)
            {
                ToLogFileError(string.Concat(new object[] { DateTime.Now, String.Format(" Ошибка в Request  " + ex.Message), Environment.NewLine }));
                Console.WriteLine("Ошибка в Request " + ex.Message);

            }

        }
        
        //разбор мобильного номера к виду
        private unsafe static string DelNoDigits(string s)
        {
            string Result = new string('\0', s.Length);
            fixed (char* _PString = s)
            {
                fixed (char* _PResult = Result)
                {
                    char* PString, PResult; char c; int Len = 0;
                    for (PString = _PString, PResult = _PResult; (c = *PString) != 0; PString++)
                    {
                        if ((c >= '0') && (c <= '9'))
                        {
                            *PResult++ = c;
                            Len++;
                            if(Len==2&&c=='8')
                            {
                              return  Result = "Городской номер";
                               // break;
                            }
                        }
                    }
                    return Result.Substring(0, Len);
                }
            }
        }

        //коннектор к sql
        static void connector()
        {
            // Connection String для прямого подключения к Oracle.
            connStringOracle = "Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = 10.0.112.27)(PORT = 1521))(CONNECT_DATA = (SERVER = DEDICATED)(SERVICE_NAME =ORCL))); Password=Qwerty123;User ID=progs_user";
            //
            connor.ConnectionString = connStringOracle;
            try
            {
                connor.Open();
                Console.WriteLine("Соединено успешно");
            }
            catch (Exception ex)
            {
                ToLogFileError(string.Concat(new object[] { DateTime.Now, String.Format(" Ошибка connector " +ex.Message), Environment.NewLine }));
            }
            //finally
            //{
            //    conn.Dispose();
            //}
        }
        
        //лог ошибок catch
        private static void ToLogFileError(string text)
            {
               try
               {
                string pathDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\logs";
                DirectoryInfo dirInfo = new DirectoryInfo(pathDir);
                if (!dirInfo.Exists)
                    dirInfo.Create();
               
                string pathFile = pathDir + "\\" + "error_" + DateTime.Today.ToString("yyyy-MM-dd") + ".log";

                File.AppendAllText(pathFile, text);
                }
                catch(Exception ex)
               {
                ToLogFileError(string.Concat(new object[] { DateTime.Now, String.Format("Ошибка 218 строка " + ex.Message), Environment.NewLine }));
            }
        }

        //лог отправленных 
        private static void ToLogFileSending(string text)
        {
            try
            {
                string pathDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\logssending";
                DirectoryInfo dirInfo = new DirectoryInfo(pathDir);
                if (!dirInfo.Exists)

                    dirInfo.Create();

                string pathFile = pathDir + "\\" + "send_to_" + DateTime.Today.ToString("yyyy-MM-dd") + ".log";

                File.AppendAllText(pathFile, text);
            }
            catch (Exception ex)
            {
                ToLogFileError(string.Concat(new object[] { DateTime.Now, String.Format("Ошибка 218 строка " + ex.Message), Environment.NewLine }));
            }
        }




        //выгрузка excel
        static void exportexcel()
        {
            try
            {
                Excel.Application excelapp = new Excel.Application();
                Excel.Workbook workbook = excelapp.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.ActiveSheet;
                for (int j = 0; j < list.GetLength(1); j++)
                {
                    worksheet.Cells[1, j + 1].Columns.EntireColumn.AutoFit();//автовыравнивание яечйки
                    worksheet.Cells[1, j + 1].Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;//границы ячейки 
                    worksheet.Cells[1, j + 1].Font.Color = Excel.XlRgbColor.rgbGreen;//цвет шрифта
                    worksheet.Rows[1].Columns[j + 1].Style.Font.Size = 12;//размер шрифта
                    worksheet.Rows[1].Columns[j + 1] = ColumN[j];
                }
                Excel.Range r = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[list.GetLength(0) + 1, list.GetLength(1)]];
                r.Value = list;
                r.Columns.EntireColumn.AutoFit();
                r.Style.Font.Size = 12;//размер шрифта
                string pathDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\upl";
                DirectoryInfo dirInfo = new DirectoryInfo(pathDir);
                if (!dirInfo.Exists)
                    dirInfo.Create();
                string pathFile = pathDir + "\\" + "выгрузка_за_период_c__" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".xlsx";
                workbook.SaveAs(pathFile);
                Console.WriteLine("Excel-документ успешно сформирован");
                workbook.Close(true, misValue, misValue);
                excelapp.Quit();
            }
            catch (Exception ex)
            {
                ToLogFileError(string.Concat(new object[] { DateTime.Now, String.Format(" Ошибка 254 строка  " + ex.Message), Environment.NewLine }));
                Console.WriteLine(" Ошибка exportexcel " + ex.Message);

            }
        }


        //отправка СМС 
        static void UpdateListSMS()
        {
            try
            {
                int flag = 0;
                string Phone;
                for (int i = 0; i < list.GetLength(0); i++)
                {
                    string number = list[i, 10];
                    Phone = DelNoDigits(number);
                    if (list[i, 10] == "")
                    {
                        ToLogFileError(string.Concat(new object[] { DateTime.Now, String.Format(" Не указан номер телефона  " + list[i, 11]), Environment.NewLine }));
                        Console.WriteLine("Не указан номер " + list[i, 11] + "\n");
                    }
                    else if (Phone=="Городской номер")
                    {
                        ToLogFileError(string.Concat(new object[] { DateTime.Now, String.Format(" Указан городской номер  " + list[i, 11]), Environment.NewLine }));
                        Console.WriteLine("Указан городской номер " + list[i, 11] + "\n");
                    }
                    else
                    {
                        Phone = "+7" + Phone.Substring(Phone.Length - 10, 10);
                        string smsmessage = null;
                        flag = FindArray(Phone, flag);
                        if (flag == 1)
                        {
                            smsmessage = "Мои документы 43.Ваша заявка №" + list[i, 0] + " готова к выдаче.";
                            sendmessage(Phone, smsmessage);
                            flag = 0;
                            temper[i, 1] = "1";
                        }
                        if (flag == 2)
                        {
                            smsmessage = "Мои документы 43.Ваши документы готовы к выдаче.";
                            sendmessage(Phone, smsmessage);
                            flag = 0;
                        }

                    }
                }           
            }
            catch (NullReferenceException ex)
            {
                Console.WriteLine("Ошибка отправки UpdateListSMS \n");
                ToLogFileError(string.Concat(new object[] { DateTime.Now, String.Format("Ошибка отправки UpdateListSMS" + ex.Message), Environment.NewLine }));

            }
        }

        //поиск по массиву list, создание флага 
        static int FindArray(string Phone,int flag)
        {
            int count=0;            
            for (int i=0; i< list.GetLength(0); i++) //прогоняем массив с БД
            {
                if (Phone == list[i,10])// ищем номер в массиве
                {
                    count++;  //количество номеров в массиве                                                                   
                }
            }

            if (count == 1) { flag = 1; }
            if (count>1)//если больше одного номера в массиве
            {
                for (int j = 0; j < list.GetLength(0); j++)//прогоняем временный массив
                {
                    if (temper[j, 0] == Phone)//ищем номер во временном массиве
                    {
                        if (temper[j,1]!="1")//проверяем статус номера, если не отправляли то отправим
                        {
                            temper[j, 1] = "1";//добавляем всем номерам статус 1(отправлено)
                            flag = 2;//ставим такой флаг, отправляем один раз                            
                        }
                    }
                }
            }         


             
            return flag;

        }

        static void sendmessage (string Phone,string smsmessage)
        {
            //раскоментировать отправку
            // string url = String.Format("http://192.168.112.174:13025/cgi-bin/sendsms?username=user1&password=pass&charset=UTF-8&coding=2&to={0}&text={1}", Phone, smsmessage);
            string url = String.Format("http://192.168.0.1:13025");
            //Подключение по http и отправка сообщения

            HttpWebResponse resp = null;

            try
            {
                WebRequest webr = WebRequest.Create(url);
                resp = (HttpWebResponse)webr.GetResponse();
                Console.WriteLine("На номер  " + Phone + " отправлено успешно сообщение  " + smsmessage + "\n");
                ToLogFileSending(string.Concat(new object[] { DateTime.Now, String.Format(" На номер  " + Phone + " отправлено успешно сообщение  " + smsmessage), Environment.NewLine }));
                resp.Close();//обязательное закрытие соединения                       

            }
            catch (WebException ex)
            {
                resp = (HttpWebResponse)ex.Response;
                Console.WriteLine("Ошибка отправки смс на номер " + Phone + " \n");
                ToLogFileSendingError(string.Concat(new object[] { DateTime.Now, String.Format(" Ошибка отправки смс " + Phone+ " " + ex.Message), Environment.NewLine }));
            }

        }

        private static void ToLogFileSendingError(string text)
        {
            try
            {
                string pathDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\logssendingerror";
                DirectoryInfo dirInfo = new DirectoryInfo(pathDir);
                if (!dirInfo.Exists)

                    dirInfo.Create();

                string pathFile = pathDir + "\\" + "send_to_" + DateTime.Today.ToString("yyyy-MM-dd") + ".log";

                File.AppendAllText(pathFile, text);
            }
            catch (Exception ex)
            {
                ToLogFileError(string.Concat(new object[] { DateTime.Now, String.Format("Ошибка 218 строка " + ex.Message), Environment.NewLine }));
            }
        }






    }
}
