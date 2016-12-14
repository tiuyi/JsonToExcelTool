using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonToExcelTool.Models
{
    public class DtoOffer
    {
        public string offerType { get; set; }
        public string linkDest { get; set; }

        public string convertsOn { get; set; }

        public DtoAppInfo appInfo { get; set; }

        public List<DtoTarget> targets { get; set; }

  
    }
}
