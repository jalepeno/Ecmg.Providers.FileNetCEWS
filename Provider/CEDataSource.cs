using Documents.Data;
using Documents.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Documents.Providers.FileNetCEWS.Provider
{
  public class CEDataSource : DataSource
  {

    #region Constructors

    public CEDataSource() : base() { }

    public CEDataSource(string connectionString, string queryTarget, string sourceColumn, Criteria criteria) : base(connectionString, queryTarget, sourceColumn, criteria) { }

    public CEDataSource(string xmlFilePath) : base(xmlFilePath) { }

    #endregion

    #region Public Methods

    #endregion

    #region Private Methods

    //private string BuildSQLStringContent()
    //{
    //  try
    //  {
    //    StringBuilder sqlBuilder = new StringBuilder();

    //    sqlBuilder.Append("SELECT ");

    //    if (LimitResults > 0) { sqlBuilder.Append($" TOP {LimitResults} "); }
    //    sqlBuilder.Append($"{ResultColumnsString} FROM [{QueryTarget}]");

    //    return sqlBuilder.ToString();

    //  }
    //  catch (Exception ex)
    //  {
    //    ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
    //    //  Re - throw the exception to the caller
    //    throw;
    //  }
    //}

    #endregion

  }
}
