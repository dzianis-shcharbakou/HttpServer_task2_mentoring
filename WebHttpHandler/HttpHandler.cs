using ClosedXML.Excel;
using NorthwindDB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Web;
using System.Web.Hosting;
using System.Web.UI.WebControls;
using System.Xml.Linq;

namespace WebHttpHandler
{
    public class HttpHandler : IHttpHandler
    {
        public bool IsReusable => true;

        public void ProcessRequest(HttpContext context)
        {
            var request = context.Request;
            var response = context.Response;

            ReportQuery reportQuery;
            if (IsParamInQuery(request.Url))
            {
                reportQuery = MapQuery(request);
            }
            else
            {
                reportQuery = MapBody(request);
            }

            List<Order> resultOrders = reportQuery.Result();

            XDocument xmlReport;
            XLWorkbook excelReport;
            bool isXml = IsXml(request.AcceptTypes);

            if (isXml)
            {
                xmlReport = CreateXml(resultOrders);
                response.ContentType = "text/xml";
                response.AddHeader("content-disposition", "attachment;filename=\"HelloWorld.xml\"");
                response.Write(xmlReport);
            }
            else
            {
                excelReport = CreateExcel(resultOrders);
                response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                response.AddHeader("content-disposition", "attachment;filename=\"HelloWorld.xlsx\"");

                using (MemoryStream stream = new MemoryStream())
                {
                    excelReport.SaveAs(stream);
                    stream.WriteTo(response.OutputStream);
                    stream.Close();
                }
            }
        }

        private bool IsXml(string[] acceptTypes)
        {
            if (acceptTypes.Contains("text/xml") || acceptTypes.Contains("application/xml"))
            {
                return true;
            }

            return false;
        }
        private bool IsParamInQuery(Uri uri)
        {
            return uri.PathAndQuery.Contains("?");
        }

        #region Reports
        private XLWorkbook CreateExcel(List<Order> result)
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Orders");
            worksheet.Cell("A" + 1).Style.Fill.BackgroundColor = XLColor.Red;
            worksheet.Cell("B" + 1).Style.Fill.BackgroundColor = XLColor.Red;
            worksheet.Cell("C" + 1).Style.Fill.BackgroundColor = XLColor.Red;

            worksheet.Cell("A" + 1).Value = "OrderID";
            worksheet.Cell("B" + 1).Value = "CustomerID";
            worksheet.Cell("C" + 1).Value = "OrderDate";

            int row = 2;
            foreach (var item in result)
            {
                worksheet.Cell("A" + row).Value = item.OrderID;
                worksheet.Cell("B" + row).Value = item.CustomerID;
                worksheet.Cell("C" + row).Value = item.OrderDate.Value;
                row++;
            }

            return workbook;
        }

        private XDocument CreateXml(List<Order> result)
        {
            XDocument xdoc = new XDocument();
            XElement orders = new XElement("orders");

            foreach (var item in result)
            {
                XElement order = new XElement("order");
                XAttribute orderId = new XAttribute("orderId", item.OrderID);
                XElement customerId = new XElement("customerId", item.CustomerID);
                XElement orderDate = new XElement("orderDate", item.OrderDate.Value);

                order.Add(orderId);
                order.Add(customerId);
                order.Add(orderDate);

                orders.Add(order);
            }

            xdoc.Add(orders);

            return xdoc;
        }
        #endregion

        #region MapRequestParameters
        private ReportQuery MapBody(HttpRequest request)
        {
            var customerId = Convert.ToString(request["customerId"]);
            var dateFrom = Convert.ToDateTime(request["dateFrom"]);
            var dateTo = Convert.ToDateTime(request["dateTo"]);
            var take = Convert.ToInt32(request["take"]);
            var skip = Convert.ToInt32(request["skip"]);

            return new ReportQuery(customerId, dateFrom, dateTo, take, skip);
        }

        private ReportQuery MapQuery(HttpRequest request)
        {
            var customerId = Convert.ToString(request.Url.ParseQueryString()["customerId"]);
            var dateFrom = Convert.ToDateTime(request.Url.ParseQueryString()["dateFrom"]);
            var dateTo = Convert.ToDateTime(request.Url.ParseQueryString()["dateTo"]);
            var take = Convert.ToInt32(request.Url.ParseQueryString()["take"]);
            var skip = Convert.ToInt32(request.Url.ParseQueryString()["skip"]);

            return new ReportQuery(customerId, dateFrom, dateTo, take, skip);
        }
        #endregion

    }
}