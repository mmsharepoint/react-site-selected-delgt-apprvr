using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ApplyPermissions.Model
{
  internal class Request
  {
    public string URL { get; set; }
    public string Permission { get; set; }
    public string AppID { get; set; }
  }
}
