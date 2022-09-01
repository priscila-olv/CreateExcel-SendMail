using System;
using System.IO;
using System.Net;
using System.Reflection;


namespace CreateExcel_SendMail
{
    class Program
    {
        static void Main(string[] args)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            string executionMessage = "";
            try
            {
                var report = new Report();
                report.Generate();

                executionMessage = "Executado com sucesso. \n" + report.Log;
            }
            catch (Exception ex)
            {
                executionMessage = ex.Message + " > " + ex.StackTrace + ex.ToString();
            }

            StreamWriter w = new StreamWriter(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "execucao.log"), false);
            w.WriteLine(executionMessage);
            w.Close();
        }
    }
}
