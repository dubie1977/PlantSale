using System;
using System.Text;

namespace Plant_Sale_Tool
{
    public class HTMLGenerator
    {
        public HTMLGenerator()
        {
        }

        public static string GetHTMLString(Plant plant)
        {
            //var employees = DataStorage.GetAllEmployess();
            string title = string.Format("{0}", plant.Name);

            var sb = new StringBuilder();
            sb.AppendFormat(@"
                        <html>
                            <head>
                                <link rel='stylesheet' href='ReportStyles.css'>
                            </head>
                            <body>
                                <div class='header'><h1>{0}</h1></div>
                                <table align='center'>
                                    <tr>
                                        <th>Seller</th>
                                        <th>Customer</th>
                                        <th>Amount Sold</th>
                                    </tr>", title);

            foreach (var order in plant.Orders)
            {
                sb.AppendFormat(@"<tr>
                                    <td>{0}</td>
                                    <td>{1}</td>
                                    <td>{2}</td>
                                  </tr>", order.seller, order.customer, order.orderAmount);
            }

            sb.Append(@"
                                </table>
                            </body>
                        </html>");

            return sb.ToString();
        }
    }
}
