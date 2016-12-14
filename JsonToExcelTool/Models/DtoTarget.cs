using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonToExcelTool.Models
{
   public class DtoTarget
    {
       public string offerID { get; set; }
       public string approvalStatus { get; set; }

       public string offerStatus { get; set; }

       public string trackingLink { get; set; }
       public string[] countries { get; set; }
       public string[] platforms { get; set; }
       public DtoPayOut payout { get; set; }
       public string endDate { get; set; }
       public string dailyConversionCap { get; set; }
       public DtoRestrictions restrictions { get; set; }
       
    }
}
