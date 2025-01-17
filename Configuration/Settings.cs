using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Documents.Providers.FileNetCEWS.Configuration
{

  public sealed class Settings
  {

    #region Class Properties

    public List<string> AppContentExportPropertyExclusions { get; set; }

    public List<string> UserContentExportPropertyExclusions { get; set; }

    #endregion

  }
}
