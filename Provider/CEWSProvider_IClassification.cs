using Documents.Core;
using Documents.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Documents.Providers.FileNetCEWS
{
  public partial class CEWSProvider : IClassification
  {

    #region Class Variables

    //  For IClassification
    DocumentClasses _documentClasses;
    ClassificationProperties _properties;
    DocumentClasses _requestedDocumentClasses;

    #endregion

    #region IClassification Implementation

    public ClassificationProperties ContentProperties
    {
      get
      {
        string errorMessage = string.Empty;
        try
        {
          if (_properties == null) { _properties = _cewsServices.GetAllPropertyTemplates(ref errorMessage); }
          return _properties;
        }
        catch (Exception ex)
        {
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
          if (string.IsNullOrEmpty(errorMessage)) { ApplicationLogging.WriteLogEntry(errorMessage, MethodBase.GetCurrentMethod(), TraceEventType.Error, 202); }
          //  Re - throw the exception to the caller
          throw;
        }
      }
    }

    public DocumentClasses DocumentClasses
    {
      get
      {
        try
        {
          if (_documentClasses == null) { _documentClasses = _cewsServices.GetAllDocumentClassDefinitions(ContentExportPropertyExclusions); }
          return _documentClasses;
        }
        catch (Exception ex)
        {
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
          //  Re - throw the exception to the caller
          throw;
        }
      }
    }

    public DocumentClass get_DocumentClass(string documentClassName)
    {
      try
      {
        if ((_requestedDocumentClasses == null) || (_requestedDocumentClasses[documentClassName] == null))
        {
          DocumentClass documentClass = _cewsServices.GetDocumentClassDefinition(documentClassName);
          if (_requestedDocumentClasses == null) { _requestedDocumentClasses = new DocumentClasses(); }
          if (documentClass != null) { _requestedDocumentClasses.Add(documentClass); }
          return documentClass;
        }
        else { return _requestedDocumentClasses[documentClassName]; }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    #endregion

  }
}
