using Documents.Arguments;
using Documents.Core;
using Documents.Data;
using Documents.Utilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Documents.Providers.FileNetCEWS.Provider
{
  public class CEWSSearch : CSearch
  {

    #region Class Variables

    CEDataSource _CEDataSource;
    string _url;
    string _userName;
    string _password;
    string _objectStoreName;

    #endregion

    #region Constructors

    public CEWSSearch() { }

    public CEWSSearch(string connectionString) : this(new CEWSProvider(connectionString), new Criteria()) { }

    public CEWSSearch(ContentSource contentSource) : this(contentSource, new Criteria()) { }

    public CEWSSearch(ContentSource contentSource, Criteria criteria) : this((CEWSProvider)contentSource.Provider, criteria) { }

    public CEWSSearch(CEWSProvider provider) : this(provider, new Criteria()) { }

    public CEWSSearch(CProvider provider) : this((CEWSProvider)provider, new Criteria()) { }

    public CEWSSearch(IProvider provider, Criteria criteria) : base(ref provider, criteria, ID_COLUMN, DOCUMENT_QUERY_TARGET, new CEDataSource())
    {
      try
      {
        UserName = (string)provider.ProviderProperties["UserName"].PropertyValue;
        Password = (string)provider.ProviderProperties["Password"].PropertyValue;
        ObjectStore = (string)provider.ProviderProperties["ObjectStore"].PropertyValue;
        URL = ((CEWSProvider)provider).URL;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    #endregion

    #region Public Properties

    public string URL { get { return _url; } set { _url = value; } }

    public string UserName { get { return _userName; } set { _userName = value; } }

    public string Password { get { return _password; } set { _password = value; } }
    
    public string ObjectStore { get { return _objectStoreName; } set { _objectStoreName = value; } }

    protected override string DefaultDelimitedResultColumns { get { return "Id,DocumentTitle"; } }

    public override string DefaultQueryTarget { get { return DOCUMENT_QUERY_TARGET; } }

    #endregion

    #region Public Methods

    #region Override Methods

    public override SearchResultSet Execute(SearchArgs args)
    {
      try
      {
        return ExecuteSearch(args);
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public override SearchResultSet Execute()
    {
      try
      {
        return ExecuteSearch();
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public override DataTable SimpleSearch(SimpleSearchArgs Args)
    {
      try
      {
        throw new NotImplementedException();
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    #endregion

    protected SearchResultSet ExecuteSearch(SearchArgs args)
    {
      string sql = string.Empty;
      try
      {
        string errorMessage = string.Empty;
        SearchResultSet resultSet;
      try
      {
          InitializeSearch(args);
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
          return new SearchResultSet(ex);
      }

        //  Initialize the DataSource with the correct QueryTarget and IDColumn
        if (args.Document != null)
        {
          InitializeDataSource(ID_COLUMN, args.Document.DocumentClass);
        }
        else
        {
          InitializeDataSource(ID_COLUMN, DOCUMENT_QUERY_TARGET);
        }

        //  If the Criteria includes the Document Class we need to remove it.
        //  It will be set as the query target in this case.
        foreach (Criterion criterion in Criteria)
        {
          if (criterion.PropertyName == "Document Class")
          {
            Criteria.Remove(criterion);
            break;
          }
        }

        //  Copy the document object
        Document document = args.Document;
        //  Clear the document class so that it will not be parsed as a where clause.
        //  By passing it to the InitializeDataSource method above we effectively
        //  set the query target to the document class.
        if (document != null) { document.DocumentClass = string.Empty; }
        
        if (args.UseDocumentValuesInCriteriaValues)
        {
          sql = DataSource.BuildSQLString(document, args.VersionIndex, ref errorMessage);
        }
        else
        {
          sql = DataSource.BuildSQLString(ref errorMessage);
        }

        //  Write the SQL statement to the log for debugging
        ApplicationLogging.WriteLogEntry($"CEWSSearch::ExecuteSearch SQL Initialized as '{sql}'", System.Diagnostics.TraceEventType.Verbose, 9411);

        if (!string.IsNullOrEmpty(errorMessage))
        {
          resultSet = new SearchResultSet(new ApplicationException($"Error Creating SQL Statement: {errorMessage}"));
          return resultSet;
        }

        CEWSServices cewsServices = new CEWSServices((CEWSProvider)Provider);

        resultSet = cewsServices.GetDocumentIDSet(this);

        if (resultSet.Exception != null)
        {
          resultSet = new SearchResultSet(new Exception($"Unable to Execute Search: '{sql}'", resultSet.Exception));
          return resultSet;
        }

        return resultSet;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        args.ErrorMessage += $"Exception: '{ex.Message}': SQL Statement: '{sql}'";
        throw new ApplicationException($"Unable to Execute Search: {sql}, ErrorMessage: {args.ErrorMessage}");
      }
    }

    protected SearchResultSet ExecuteSearch()
    {
      try
      {

        try
        {
          InitializeSearch();
        }
        catch (Exception ex)
        {
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
          return new SearchResultSet(ex);
        }

        //  If the Criteria includes the Document Class we need to remove it.
        //  It will be set as the query target in this case.
        foreach (Criterion criterion in Criteria)
        {
          if (criterion.PropertyName == "Document Class")
          {
            Criteria.Remove(criterion);
            break;
          }
        }

        CEWSServices cewsServices = new CEWSServices((CEWSProvider)Provider);

        return cewsServices.ExecuteSearch(this);

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        return new SearchResultSet(new Exception($"Unable to Execute Search: {ex.Message}", ex));
      }
    }

    #endregion

  }
}
