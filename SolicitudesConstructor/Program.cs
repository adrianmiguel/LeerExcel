using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using SolicitudesConstructor.Logica;

namespace SolicitudesConstructor
{
    class Program
    {
        static void Main(string[] args)
        {
            DataSet ds = new DataSet();
            XmlDocument xml = new XmlDocument();

            LeerExcelForma1 l1 = new LeerExcelForma1();
            ds = l1.LeerExcel();

            xml.LoadXml(ds.GetXml());

            XmlNodeList infoSolicitudes;

            infoSolicitudes = xml.SelectNodes("Solicitudes/Table");

        }
    }
}
