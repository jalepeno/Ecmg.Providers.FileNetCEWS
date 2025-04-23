using CPEServiceReference;
using Documents.Annotations;
using Documents.Arguments;
using Documents.Core;
using Documents.Data;
using Documents.Exceptions;
using Documents.Providers.FileNetCEWS.Annotations;
using Documents.Providers.FileNetCEWS.Provider;
using Documents.Transformations;
using Documents.Utilities;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static Documents.Core.Content;
using static Documents.Providers.FileNetCEWS.CEWSServices;
using Document = Documents.Core.Document;

namespace Documents.Providers.FileNetCEWS
{
  public partial class CEWSProvider : IDocumentExporter
  {

    #region IDocumentExporter Implementation

    public override string ExportPath
    { 
      get { return base.ExportPath; }
      set { base.ExportPath = value; }
    }

    public event DocumentExportEventHandler DocumentExported;
    public event FolderDocumentExportEventHandler FolderDocumentExported;
    public event FolderExportedEventHandler FolderExported;
    public event DocumentExportErrorEventHandler DocumentExportError;
    public event DocumentExportMessageEventHandler DocumentExportMessage;

    public long DocumentCount(string lpFolderPath, RecursionLevel lpRecursionLevel = RecursionLevel.ecmThisLevelOnly)
    {
      throw new NotImplementedException();
    }

    public bool ExportDocument(string id)
    {
      try
      {
        return ExportDocument(new ExportDocumentEventArgs(id));
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public bool ExportDocument(ExportDocumentEventArgs args)  //  TODO: Test and Add Annotation export
    {
      Core.Document document = null;
      //string documentPath;
      //string versionPath;
      //string[] folders;
      string id = string.Empty;
      try
      {
        if (args == null) { throw new ArgumentNullException(nameof(args)); }
        id = args.Id;
        
        //  If we are not going to get the content, we will not get the annotations either, regardless of whether or notthey were requested.
        if (!args.GetContent) { args.GetAnnotations = false; }

        document = GetDocumentById(id, args.GetContent, args.GetAnnotations, TargetVersion.AllVersions, "", args.StorageType, args.Transformation);

        args.Document = document;

        return ExportDocumentComplete(this, args);

      }
      catch (ContentTooLargeException largeContentEx)
      {
        ApplicationLogging.LogException(largeContentEx, MethodBase.GetCurrentMethod());
        throw;
      }
      catch (ZeroLengthContentException zeroContentEx)
      {
        ApplicationLogging.LogException(zeroContentEx, MethodBase.GetCurrentMethod());
        throw;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        throw new DocumentException(id, $"GetDocumentById failed: '{ex.Message}'", ex);
      }
    }

    public void OnDocumentExported(ref DocumentExportedEventArgs e)
    {
      DocumentExported?.Invoke(this, ref e);
    }

    public void OnDocumentExportError(ref DocumentExportErrorEventArgs e)
    {
      DocumentExportError?.Invoke(this, e);
    }

    public void OnDocumentExportMessage(ref WriteMessageArgs e)
    {
      DocumentExportMessage?.Invoke(this, e);
    }

    public void OnFolderDocumentExported(ref FolderDocumentExportedEventArgs e)
    {
      FolderDocumentExported?.Invoke(this, ref e);
    }

    public void OnFolderExported(ref FolderExportedEventArgs e)
    {
      FolderExported?.Invoke(this, ref e);
    }

    #endregion

    #region Public Enums

    public enum TargetVersion
    {
      AllVersions = -100,
      FirstVersion = -101,
      LatestVersion = -102,
      Release = 103
    }

    #endregion

    #region Private Methods

    /// <summary>
    /// Gets the document using the unique document id
    /// By calling GetDocumentWithContent with no destination folder no cdf file will be written out
    /// </summary>
    /// <param name="id">The unique identifier of the document/version.</param>
    /// <param name="exportContent">Determines whether or not we will get the document content as well or only the document metadata.  
    /// This parameter is used with a value of false for the implementation of the GetDocument method. 
    /// It is used with a value of true with the ExportDocument methods.
    /// </param>
    /// <param name="targetVersion"></param>
    /// <param name="destinationFolder"></param>
    /// <param name="storageType"></param>
    /// <param name="transformation"></param>
    /// <returns>A Document object</returns>
    private Document GetDocumentById(string id, bool exportContent, bool exportAnnotations, TargetVersion targetVersion = TargetVersion.AllVersions, string destinationFolder = "", StorageTypeEnum storageType = StorageTypeEnum.Reference, Transformation transformation = null)
    {
      try
      {
        Document document = null;
        Core.Version version;
        //ECMProperty property;
        //Values values;
        //string documentPath;
        //string versionPath;
        //string copyPath;
        //string[] folders;
        int versionIndex = 0;
        string versionSeriesId;
        SearchResultSet versionSeriesInfo;
        Document versionDocument;

        //  Check to make sure we got an id
        if (string.IsNullOrEmpty(id)) { throw new ArgumentNullException(nameof(id), "The specified ID is empty.  Please provide a valid document identifier."); }

        //  Get the VersionSeriesId
        versionSeriesId = _cewsServices.GetVersionSeriesID(id);

        //  Get the list of versions
        versionSeriesInfo = GetVersionSeriesList(id);

        //  Make sure we actually got some results
        if (versionSeriesInfo.Count == 0) { throw new DocumentException(id, "The version series list could not be found"); }

        switch (targetVersion)
        {
          case TargetVersion.AllVersions:
            {
              foreach (DataRow dataRow in versionSeriesInfo.ToDataTable().Rows)
              {
                versionDocument = GetDocumentVersionById(id, dataRow.ItemArray[0].ToString(), exportContent, exportAnnotations, versionIndex, destinationFolder, storageType);
                version = versionDocument.GetLatestVersion();

                //  Set the VersionId
                version.SetPropertyValue("VersionId", dataRow.ItemArray[0].ToString(), true, Core.PropertyType.ecmString);

                //  Set the MajorVersionNumber
                version.SetPropertyValue("MajorVersionNumber", dataRow.ItemArray[1].ToString(), true, Core.PropertyType.ecmLong);

                //  Set the MinorVersionNumber
                version.SetPropertyValue("MinorVersionNumber", dataRow.ItemArray[2].ToString(), true, Core.PropertyType.ecmLong);

                //  Set the VersionStatus
                version.SetPropertyValue("VersionStatus", dataRow.ItemArray[3].ToString(), true, Core.PropertyType.ecmLong);

                //  Set the VersionSeriesId
                version.SetPropertyValue("VersionSeriesId", versionSeriesId, true, Core.PropertyType.ecmString);

                if (document == null)
                {
                  document = versionDocument;
                  document.ContentSource = ContentSource;
                }
                else
                {

                  //  Check for folders filed in
                  if (versionDocument.PropertyExists("Folders Filed In"))
                  {
                    MultiValueStringProperty foldersProperty = (MultiValueStringProperty)versionDocument.Properties["Folders Filed In"].Clone();
                    document.Properties.Add(foldersProperty);
                  }

                  //  Do not add this version to the export if it is in a Reservation state
                  VersionStatus versionstatus = (VersionStatus)version.Properties["VersionStatus"].Value;
                  if (versionstatus != VersionStatus.Reservation)
                  {
                    document.Versions.Add(version);
                  }
                  else
                  {
                    ApplicationLogging.WriteLogEntry($"Skipped export of version '{dataRow.ItemArray[0]}' of document '{versionSeriesId}' as it is in a reservation state.", MethodBase.GetCurrentMethod(), TraceEventType.Warning, 8736);
                  }

                }
                versionIndex++;
              }
              break;
            }

          default:
            {
              throw new NotImplementedException();
            }
        }

        document.ID = id;

        //  Return the document
        return document;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private  Document GetAnnotationDocument(string id)
    {
      try
      {
        ObjectValue annotationProperties;
        Document annotationDocument = new Document();
        Core.Version version;
        ECMProperty property;

        //  Make sure the annotation exists
        if (!DocumentExists(id, "Annotation")) { throw new DocumentException(id, $"Annotation '{id}' does not exist."); }

        //  Get the annotation information
        annotationProperties = _cewsServices.GetDocumentProperties(id, RequestType.ExtendedProperties, null, true, true, "Annotation");

        annotationDocument.ObjectID = id;
        annotationDocument.ID = id;

        //  Get the Document Class
        annotationDocument.DocumentClass = annotationProperties.classId;

        version = annotationDocument.CreateVersion();
        version.ID = 0;

        foreach (CPEServiceReference.PropertyType propertyType in annotationProperties.Property)
        {
          property = CreateECMProperty(propertyType);
          if (property != null) { version.Properties.Add(property); }
        }

        //  Get the content elements
        ListOfObject contentElements = (ListOfObject)CEWSServices.GetPropertyByID(annotationProperties, "ContentElements");
        //Contents contents = _cewsServices.GetContents(id, contentElements);
        Contents contents = ReadContents(contentElements);
        version.Contents = contents;

        //  Sort the version properties
        version.Properties.Sort();

        //  Add the version to the document
        annotationDocument.Versions.Add(version);

        annotationDocument.StorageType = StorageTypeEnum.Reference;

        //  Return the document
        return annotationDocument;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private Contents ReadContents(ListOfObject contentElements)
    {
      Contents contents = new Contents();
      try
      {
        NamedStream contentStream;
        CPEServiceReference.PropertyType contentProperty;
        SingletonString retrievalNameProperty;
        foreach (DependentObjectType contentElement in contentElements.Value)
        {
          retrievalNameProperty = (SingletonString)GetPropertyByName("RetrievalName", contentElement);
          contentProperty = GetPropertyByName("Content", contentElement);
          InlineContent inlineContent = (InlineContent)((CPEServiceReference.ContentData)contentProperty).Value;

          contentStream = new NamedStream(Helper.CopyByteArrayToStream(inlineContent.Binary), retrievalNameProperty.Value);
          contents.Add(contentStream);
        }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
      return contents;
    }

    private Document GetDocumentVersionById(string versionSeriesId, string id, bool exportContent, bool exportAnnotations, int versionIndex, string destinationFolder = "", StorageTypeEnum storageType = StorageTypeEnum.Reference)
    {
      try
      {
        ObjectValue basicDocumentProperties;
        ObjectValue extendedDocumentProperties;

        Document document = new Document();
        Core.Version version;
        //Content content;
        ECMProperty property;
        Values values = new Values();
        //string documentPath;
        //string versionPath;
        //string copyPath;
        string[] folders;
        //SearchResultSet versionSeriesInfo;
        //string contentType;

        //  Check to make sure we got an id
        if (id == null) { throw new ArgumentNullException(nameof(id)); }
        if (id.Length == 0) { throw new ArgumentException("The specified ID is a zero length string.  Please provide a valid document identifier.", nameof(id)); }

        //  Make sure the document exists
        if (!DocumentExists(id)) { throw new DocumentException(id, $"Document '{id}' does not exist."); }

        //  Get the basic document information
        basicDocumentProperties = _cewsServices.GetDocumentProperties(id, RequestType.BasicProperties, ContentExportPropertyExclusions.ToArray(), exportContent, exportAnnotations);

        //  Get the extended document information
        extendedDocumentProperties = _cewsServices.GetDocumentProperties(id, RequestType.ExtendedProperties, null, exportContent, exportAnnotations);

        //  Create the Document object
        document.ID = versionSeriesId;
        document.ObjectID = id;

        //  Get the Document Class
        document.DocumentClass = basicDocumentProperties.classId;

        //  Get the folders filed in
        values.Clear();
        folders = CEWSServices.GetFoldersFiledIn((EnumOfObject)CEWSServices.GetPropertyByID(extendedDocumentProperties, "FoldersFiledIn"));
        if (folders != null) 
        { 
          for (int i = 0; i < folders.Length; i++) { values.Add(folders[i]); } 
          document.FolderPathsProperty.Values = values;
        }

        version = document.CreateVersion();
        version.ID = versionIndex;

        foreach (CPEServiceReference.PropertyType propertyType in basicDocumentProperties.Property)
        {
          property = CreateECMProperty(propertyType);
          if (property != null) { version.Properties.Add(property); }
        }

        if (exportContent)
        {
          ////  Get the Content Retrieval Names
          ListOfObject contentElements = (ListOfObject)CEWSServices.GetPropertyByID(extendedDocumentProperties, "ContentElements");
          //string[] retrievalNames = new string[0];
          //int contentElementCounter = 1;
          //if (contentElements.Value != null)
          //{
          //  Array.Resize(ref retrievalNames, contentElements.Value.Length);   
          //  foreach (DependentObjectType contentElement in contentElements.Value)
          //  {
          //    if (string.Compare(contentElement.classId, "ContentReference", true) != 0)
          //    {
          //      //  It is not a content reference, get the retrieval name
          //      retrievalNames[contentElementCounter] = CEWSServices.GetRetrievalName(contentElement);
          //    }



          //    contentElementCounter++;
          //  }

          //  //  Get the Content
          //  //  For each attachment that we got back, write the data to a file
          //  ContentRequest
          //  GetContentRequest contentRequest = new GetContentRequest() { ;
          //  if ()
          //}
          Contents contents = _cewsServices.GetContents(id, contentElements);
          version.Contents = contents;
        }

        if (exportAnnotations)
        {
          EnumOfObject annotations = (EnumOfObject)CEWSServices.GetPropertyByID(extendedDocumentProperties, "Annotations");
          Console.WriteLine($"{annotations.Value.Length} annotations found.");
          AnnotationExporter annotationExporter = new AnnotationExporter();
          Document annotationDocument;
          Stream annotationStream;
          string annotationMimeType;

          foreach (var nativeAnnotation in annotations.Value)
          {
            annotationDocument = GetAnnotationDocument(nativeAnnotation.objectId);
            //annotationStream = annotationDocument.LatestVersion.PrimaryContent.ToMemoryStream();
            //annotationMimeType = annotationDocument.LatestVersion.PrimaryContent.MIMEType;
            //  Finish propery implementing this later, for now, assume the annotation content is valid
            //Annotation ctsAnnotation = annotationExporter.ExportAnnotationObject(annotationStream, annotationMimeType);

            Annotation ctsAnnotation = new Annotation(annotationDocument.LatestVersion.PrimaryContent) { ID = annotationDocument.ID };
            version.Contents.First().Annotations.Add(ctsAnnotation);
          }

        }

        //  Sort the version properties
        version.Properties.Sort();

        //  Add the version to the document
        document.Versions.Add(version);

        document.StorageType = storageType;

        //  Return the document
        return document;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private ECMProperty CreateECMProperty(CPEServiceReference.PropertyType propertyType)
    {
      ECMProperty property = null;
      try
      {
        if (propertyType.GetType().Name.StartsWith("ListOf"))
        {
          switch (propertyType.GetType().Name)
          {
            case "ListOfString":
              {
                ListOfString ceProperty = (ListOfString)propertyType;
                property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmString, ceProperty.propertyId, Cardinality.ecmMultiValued);
                if (ceProperty.Value != null)
                {
                  Values values = new Values();
                  string[] ceValues = ceProperty.Value;
                  for (int i = 0; i < ceValues.Length; i++) { values.Add(ceValues[i]); }
                  property.Values = values;
                }
                break;
              }

            default:
              {
                //  TODO: Implement Multi-valued retreival for other than strings
                break;
              }
          }
        }
        else
        {
          //  This is a single-valued property
          switch (propertyType.GetType().Name)
          {
            case "SingletonString":
              {
                SingletonString ceProperty = (SingletonString)propertyType;
                property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmString, ceProperty.propertyId, ceProperty.Value);
                break;
              }

            case "SingletonDateTime":
              {
                SingletonDateTime ceProperty = (SingletonDateTime)propertyType;
                property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmDate, ceProperty.propertyId, Cardinality.ecmSingleValued);
                if (ceProperty.ValueSpecified) { property.Value = ceProperty.Value; }
                break;
              }

            case "SingletonId":
              {
                SingletonId ceProperty = (SingletonId)propertyType;
                property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmGuid, ceProperty.propertyId, ceProperty.Value);
                break;
              }

            case "SingletonInteger32":
              {
                SingletonInteger32 ceProperty = (SingletonInteger32)propertyType;
                property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmLong, ceProperty.propertyId, Cardinality.ecmSingleValued);
                if (ceProperty.ValueSpecified) { property.Value = ceProperty.Value; }
                break;
              }

            case "SingletonBoolean":
              {
                SingletonBoolean ceProperty = (SingletonBoolean)propertyType;
                property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmBoolean, ceProperty.propertyId, Cardinality.ecmSingleValued);
                if (ceProperty.ValueSpecified) { property.Value = ceProperty.Value; }
                break;
              }

            case "SingletonFloat64":
              {
                SingletonFloat64 ceProperty = (SingletonFloat64)propertyType;
                property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmDouble, ceProperty.propertyId, Cardinality.ecmSingleValued);
                if (ceProperty.ValueSpecified) { property.Value = ceProperty.Value; }
                break;
              }

            case "SingletonObject":
              {
                SingletonObject ceProperty = (SingletonObject)propertyType;
                property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmObject, ceProperty.propertyId, Cardinality.ecmSingleValued);
                if (ceProperty.Value != null) { property.Value = ((ObjectReference)ceProperty.Value).objectId; }
                break;
              }

            case "SingletonBinary":
              {
                SingletonBinary ceProperty = (SingletonBinary)propertyType;
                property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmBinary, ceProperty.propertyId, Cardinality.ecmSingleValued);
                property.Value = ceProperty.Value;
                break;
              }

            default:
              {
                break;
              }
          }
        }

        if (property != null)
        {
          return property;
        }
        else
        {
          return null;
        }

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    protected virtual SearchResultSet GetVersionSeriesList(string versionId)
    {
      try
      {
        CEWSSearch search = (CEWSSearch)Search;
        string versionSeriesId = _cewsServices.GetVersionSeriesID(versionId);

        search.Reset();
        search.DataSource.QueryTarget = "Document";
        search.DataSource.ResultColumns.Add("MajorVersionNumber");
        search.DataSource.ResultColumns.Add("MinorVersionNumber");
        search.DataSource.ResultColumns.Add("VersionStatus");
        Criterion criterion = new Criterion("VersionSeries") { Value = versionSeriesId, DataType = Criterion.pmoDataType.ecmObject };
        search.Criteria.Add(criterion);
        search.DataSource.OrderBy.Add(new OrderItem("MajorVersionNumber", SortDirection.Asc));
        search.DataSource.OrderBy.Add(new OrderItem("MinorVersionNumber", SortDirection.Asc));

        return search.Execute();

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }

      #endregion
    }
  }
}
