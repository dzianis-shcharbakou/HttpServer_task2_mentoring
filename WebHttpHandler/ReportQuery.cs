using NorthwindDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebHttpHandler
{
    public class ReportQuery
    {
        private string CustomerId { get; set; }
        private DateTime? DateFrom { get; set; }
        private DateTime? DateTo { get; set; }
        private int? Take { get; set; }
        private int? Skip { get; set; }

        public ReportQuery(string customerId, DateTime? dateFrom, DateTime? dateTo, int? take, int? skip)
        {
            this.CustomerId = customerId;
            this.DateFrom = dateFrom;
            this.DateTo = dateTo;
            this.Take = take;
            this.Skip = skip;
        }

        public List<Order> Result()
        {
            bool isCustomerIdQueryParam = CustomerId != default(string) ? true : false;
            bool isDateFromQueryParam = DateFrom != default(DateTime) ? true : false;
            bool isDateToQueryParam = DateTo != default(DateTime) ? true : false;
            bool isTakeQueryParam = Take != default(int) ? true : false;
            bool isSkipQueryParam = Skip != default(int) ? true : false;

            using (var contextDb = new NorthwindDB.NorthwindDB())
            {
                IQueryable<Order> result = contextDb.Orders;

                result = isCustomerIdQueryParam == true ? result.Where(x => x.CustomerID == CustomerId) : result;
                result = isDateFromQueryParam == true ? result.Where(x => x.OrderDate.Value >= DateFrom.Value) : result;
                result = isDateToQueryParam == true ? result.Where(x => x.OrderDate.Value <= DateTo.Value) : result;
                result = isTakeQueryParam == true ? result.Take(Take.Value) : result;
                result = result.OrderBy(x => x.OrderID);

                result = isSkipQueryParam == true ? result.Skip(Skip.Value) : result;

                return result.ToList();
            }
        }
    }
}