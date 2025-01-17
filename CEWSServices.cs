
#region Using

using CPEServiceReference;
using Documents.Core;
using Documents.Core.ChoiceLists;
using Documents.Data;
using Documents.Exceptions;
using Documents.Migrations;
using Documents.Utilities;
using Microsoft.IdentityModel.Protocols;
using Microsoft.IdentityModel.Tokens;
using Newtonsoft.Json.Linq;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;

//using System.IdentityModel.Protocols.WSTrust;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Remoting;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Xml.Linq;
using static OfficeOpenXml.ExcelErrorValue;
using static System.Net.WebRequestMethods;
using Values = Documents.Core.Values;

#endregion

namespace Documents.Providers.FileNetCEWS
{
  public class CEWSServices
  {

    #region Class Constants

    private const string MAJOR_VERSION = "MajorVersion";
    private const string NO_TITLE_SPECIFIED = "No Title Specified";

    const string PROP_CAN_DECLARE = "CanDeclare";
    const string PROP_CREATOR = "Creator";
    const string PROP_CREATE_DATE = "DateCreated";
    const string PROP_CHECK_IN_DATE = "DateCheckedIn";
    const string PROP_MODIFY_USER = "LastModifier";
    const string PROP_MODIFY_DATE = "DateLastModified";
    const string PROP_IS_CURRENT_VERSION = "IsCurrentVersion";
    const string PROP_MINOR_VERSION_NUMBER = "MinorVersionNumber";
    const string PROP_MAJOR_VERSION_NUMBER = "MajorVersionNumber";
    const string PROP_MIME_TYPE = "MimeType";
    const string PROP_OBJECT_NAME = "Name";
    const string PROP_OWNER = "Owner";
    const string PROP_RECORD_INFORMATION = "RecordInformation";
    const string PROP_VERSION_STATUS = "VersionStatus";
    const string PROP_VERSION_ID = "VersionId";
    const string PROP_VERSION_SERIES_ID = "VersionSeriesId";

    const string CE_ERR_MSG_OBJECT_MODIFIED = "The object has been modified since it was retrieved.";
    const string CE_ERR_MSG_VERSION_RESERVED = "The version series is already holding a reservation.";
    const string CE_ERR_MSG_UNABLE_TO_CONNECT = "The underlying connection was closed: Unable to connect to the remote server.";
    const string CE_ERR_MSG_REQUESTED_ITEM_NOT_FOUND = "Requested item not found.";
    const string CE_ERR_MSG_INVALID_PROPERTY_IDENTIFIER = "A property identifier is not valid.";
    const string CE_ERR_MSG_VALUE_EXCEEDS_MAXIMUM_PERMITTED_LENGTH = "exceeds the maximum permitted length.";
    const string CE_ERR_MSG_UNDERLYING_CONNECTION_WAS_CLOSED = "The underlying connection was closed";
    const string CE_ERR_MSG_IDENTIFIER_DOES_NOT_REFERENCE_AVAILABLE_CLASS_OF_OBJECT = "The supplied identifier does not reference an available class of object.";

    const int PRIVILEGED_WRITE_AS_INT = 268435456;

    #endregion

    #region Class Variables

    private readonly ObjectStoreScope _objectStoreScope;

    private readonly string _url;
    private readonly string _userName;
    private readonly string _password;
    private readonly string _objectStoreName;
    private readonly FNCEWS40PortTypeClient _client;
    private readonly Localization _localization;
    private readonly List<string> _contentExportPropertyExclusions;
    private ObjectReference _objectStoreReference = null;
    private CEWSProvider _provider;
    private bool? _hasPriviledgedAccess;

    #endregion

    #region Constructors

    public CEWSServices(CEWSProvider provider)
    {
      try
      {
        _url = provider.URL;
        _userName = provider.UserName;
        _password = provider.Password;
        _objectStoreName = provider.ObjectStoreName;
        _client = WSIUtil.ConfigureBinding(_userName, _password, _url);
        _localization = WSIUtil.GetLocalization();
        _contentExportPropertyExclusions = provider.ContentExportPropertyExclusions;
        _provider = provider;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    #endregion

    #region Enumerations

    public enum RequestType
    {
      BasicProperties = 0,
      ExtendedProperties = 1,
      FoldersFiledIn = 2,
      Content = 3,
      Versions = 4
    }

    #endregion

    #region Public Properties

    public bool HasPrivelegedAccess
    {
      get
      {
        if (!_hasPriviledgedAccess.HasValue) { _hasPriviledgedAccess = GetHasPriviledgedAccess(); }
        return _hasPriviledgedAccess.Value;
      }
    }

    internal ObjectReference ObjectStoreReference
    {
      get
      {
        if (_objectStoreReference == null) { _objectStoreReference = CreateObjectStoreReference(); }
        return _objectStoreReference;
      }
    }

    #endregion

    #region Public Methods

    public bool TestConnection()
    {
      try
      {
        string sql = "SELECT [Id] FROM [DocumentClassDefinition]";
        RepositorySearch repositorySearch = Factory.RepositorySearch(_objectStoreName, sql, RepositorySearchModeType.Rows);

        Task<ExecuteSearchResponse> results = _client.ExecuteSearchAsync(_localization, repositorySearch);

        results.Wait();

        if (results.Result != null)
        { return true; }
        else
        { return false; }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }

    }

    public bool AddDocument(ref Document document, ref object newCEObject, ref string errorMessage)
    {
      try
      {
        return AddDocument(ref document, null, true, VersionTypeEnum.Unspecified, ref newCEObject, ref errorMessage);
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    //public bool AddDocument(Document document, string[] filePath = null, bool checkin = true, bool isRecord = false, VersionTypeEnum lpVersionType = VersionTypeEnum.Unspecified, [Optional][DefaultParameterValue("")] ref string errorMessage,  [Optional][DefaultParameterValue(null)] ref object newCEObject)
    public bool AddDocument(ref Document document, string[] filePath, bool checkin, VersionTypeEnum versionType, ref object newCEObject, ref string errorMessage)
    {
      try
      {

        string documentName = string.Empty;
        DateTime dateCreated = new DateTime();
        bool addAsMajor = true;

        //  Build the Create action for a document
        CreateAction createVerb = Factory.CreateAction(document.DocumentClass);

        //  Assign the actions to the ChangeRequestType element
        ChangeRequestType changeRequest = new ChangeRequestType();

        if (checkin)
        {
          switch (versionType)
          {
            case VersionTypeEnum.Unspecified:
              {
                addAsMajor = IsMajorVersion(document.FirstVersion, true);
                break;
              }

            case VersionTypeEnum.Major:
              {
                addAsMajor = true;
                break;
              }

            case VersionTypeEnum.Minor:
              {
                addAsMajor = false;
                break;
              }
          }

          changeRequest.Action = new ActionType[2];

          //  Build the Checkin Action
          CheckinAction checkinVerb;
          
          //  Check to see if the version should be checked in as a Major or Minor version
          if (addAsMajor)
          {
            checkinVerb = Factory.CheckinAction(false, false);
          }
          else
          {
            checkinVerb = Factory.CheckinAction(true, true);
          }

          //  Assign Create action
          changeRequest.Action[0] = createVerb;
          //  Assign Checkin action
          changeRequest.Action[1] = checkinVerb;
        }
        else
        {
          changeRequest.Action = new ActionType[1];
          //  Assign Create action
          changeRequest.Action[0] = createVerb;
        }

        //  Specify the target object(an object store) for the actions
        //changeRequest.TargetSpecification = _objectStoreReference; ObjectStoreReference
        changeRequest.TargetSpecification = ObjectStoreReference; 
        changeRequest.id = "1";

        //  Build a list of properties to set in the new doc
        ModifiablePropertyType[] inputProperties = BuildPropertyList(document.FirstVersion);

        //  Is there any content?
        if (document.FirstVersion.HasContent())
        {
          //  Build the ContentElement list
          //BuildContentElementList(document.FirstVersion, ref inputProperties[document.FirstVersion.Properties.Count]); 
          BuildContentElementList(document.FirstVersion, ref inputProperties);
        }

        //  Assign list of document properties to set in ChangeRequestType element
        changeRequest.ActionProperties = inputProperties;

        //  Build a list of properties to exclude on the new doc object that will be returned
        string[] propertyExclusions = new string[2] { "Owner", "DateLastModified" };

        //  Assign the list of excluded properties to the ChangeRequestType element
        changeRequest.RefreshFilter = Factory.PropertyFilterType(1, true, null, propertyExclusions);

        //  Create array of ChangeRequestType elements and assign ChangeRequestType element to it 
        ChangeRequestType[] changeRequestArray = new ChangeRequestType[1] { changeRequest };

        //  Create ChangeResponseType element array
        ChangeResponseType[] changeResponseArray = null;

        ////  Build ExecuteChangesRequest element and assign ChangeRequestType element array to it
        //ExecuteChangesRequest executeChangesRequest = new ExecuteChangesRequest() { ChangeRequest = changeRequestArray };

        ////  Return a refreshed object
        //executeChangesRequest.refresh = true;
        //executeChangesRequest.refreshSpecified = true;

        try
        {
          //Task<ExecuteChangesResponse> response = _client.ExecuteChangesAsync(_localization, executeChangesRequest);
          //Task.WaitAll(response);

          //changeResponseArray = response.Result.ExecuteChangesResponse1;

          changeResponseArray = ExecuteChanges(changeRequest, true);

          //  The new document object should be returned, unless there is an error
          if (changeResponseArray == null)
          {
            errorMessage = errorMessage + "A valid object was not returned from the ExecuteChanges operation";
            return false;
          }

          if (changeResponseArray.Length < 1)
          {
            errorMessage = errorMessage + "A valid object was not returned from the ExecuteChanges operation";
            return false;
          }

          newCEObject = changeResponseArray[0];


          //  Capture value of the lstrDocumentTitle property in the returned doc object
          SingletonString nameProperty = (SingletonString)GetPropertyValue("Name", changeResponseArray[0]);
          documentName = nameProperty.Value;

          if (documentName == null)
          {
            errorMessage = "'Name' is not set on newly created document object.";
            ApplicationLogging.WriteLogEntry(errorMessage, MethodBase.GetCurrentMethod(), TraceEventType.Error, 404);
          }

          string classId = changeResponseArray[0].classId;

          //  Capture value of the ldatDateCreated property in the returned doc object
          SingletonDateTime dateCreatedProperty = (SingletonDateTime)GetPropertyValue("DateCreated", changeResponseArray[0]);
          dateCreated = dateCreatedProperty.Value;

          //  Capture value of the 'Id' property in the returned doc object
          SingletonId idProperty = (SingletonId)GetPropertyValue("Id", changeResponseArray[0]);
          document.ObjectID = idProperty.Value;

          //  If there are more versions, build the rest of the stack
          if (document.Versions.Count > 1)
          {
            UpdateDocument(document, 1, true);
          }

          //  If requested, file the document
          if (filePath != null)
          {
            int folderCount = filePath.Length;
            for (int folderCounter = 0; folderCounter < folderCount; folderCounter++)
            {
              ObjectValue value = GetFolder(filePath[folderCounter], 0, ref errorMessage);
              if (value == null)
              {
                errorMessage += $"Unable to get folder '{filePath[folderCounter]}'";
                ApplicationLogging.WriteLogEntry($"An error occurred in CEWebServices::AddDocument: {errorMessage}", MethodBase.GetCurrentMethod(), TraceEventType.Error, 401);
              }
              FileDocument(document.ObjectID, document.DocumentClass, documentName, filePath[folderCounter], ref errorMessage);
            }
          }

          string currentVersionId = string.Empty;

          //  Get the version series ID and return it in the document
          SingletonObject currentVersionProperty = (SingletonObject)GetPropertyValue("CurrentVersion", changeResponseArray[0]);
          currentVersionId = ((ObjectReference)currentVersionProperty.Value).objectId;

          //foreach (CPEServiceReference.PropertyType property in changeResponseArray[0].Property)
          //{
          //  if (property.propertyId == "VersionSeries")
          //  {
          //    SingletonObject versionSeriesProperty = (SingletonObject)property;
          //    ObjectValue versionSeriesObject = (ObjectValue)versionSeriesProperty.Value;
          //    versionSeriesId = versionSeriesObject.objectId;
          //    break;
          //  }
          //}

          if (!string.IsNullOrEmpty(currentVersionId)) 
          {
            document.LatestVersion.SetPropertyValue("CurrentVersionId", currentVersionId, true, Core.PropertyType.ecmString);
          }

          return true;

        }
        catch (Exception ex)
        {
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
          return false;
        }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public bool AddAnnotation(string annotationId, string parentDocumentId, Stream annotationStream, ref string errorMessage)
    {
      try
      {


        //  Build the Create action for thw annotation
        CreateAction createVerb = Factory.CreateAction("Annotation");

        //  Assign the actions to the ChangeRequestType element
        ChangeRequestType changeRequest = new ChangeRequestType();
        changeRequest.Action = new ActionType[1];

        //  Assign Create action
        changeRequest.Action[0] = createVerb;

        //  Specify the target object(an object store) for the actions
        //changeRequest.TargetSpecification = _objectStoreReference; ObjectStoreReference
        changeRequest.TargetSpecification = ObjectStoreReference;
        changeRequest.id = "1";

        //  Build a list of properties to set in the new annotation
        ModifiablePropertyType[] inputProperties = BuildAnnotationPropertyList(annotationId, parentDocumentId);

        //  Add the content
        //  Create an object reference to dependently persistable ContentTransfer object
        DependentObjectType contentTransfer;

        //  Create reference to the object set of ContentTransfer objects returned by the Document.ContentElements property
        ListOfObject contentElements = Factory.ListOfObject("ContentElements", true, true);

        //  Create the array to hold the content objects
        contentElements.Value = new DependentObjectType[1];

        //  Create an object reference to dependently persistable ContentTransfer object
        contentTransfer = Factory.DependentObjectType("ContentTransfer", DependentObjectTypeDependentAction.Insert, true);
        contentTransfer.Property = new CPEServiceReference.PropertyType[3];
        contentElements.Value[0] = contentTransfer;

        //Stream attachmentStream = annotationStream;
        InlineContent inlineContent = new InlineContent();

        inlineContent.Binary = new byte[annotationStream.Length];
        annotationStream.Read(inlineContent.Binary, 0, (int)annotationStream.Length);
        annotationStream.Close();

        // Create reference to Content pseudo-property
        CPEServiceReference.ContentData contentData = new CPEServiceReference.ContentData();
        contentData.Value = (ContentType)inlineContent;
        contentData.propertyId = "Content";

        // Assign Content property to ContentTransfer object
        contentTransfer.Property[0] = contentData;

        //  Create and assign ContentType string-valued property to ContentTransfer object
        SingletonString contentTypeProperty = Factory.SingletonString("ContentType", "application/octet-stream", true);
        contentTransfer.Property[1] = contentTypeProperty;

        //  Create and assign RetrievalName string-valued property to ContentTransfer object
        SingletonString retrievalNameProperty = Factory.SingletonString("RetrievalName", "file0", true);
        contentTransfer.Property[2] = retrievalNameProperty;

        Array.Resize(ref inputProperties, inputProperties.Length + 1);

        inputProperties[inputProperties.Length - 1] = contentElements;

        //  Assign list of properties to set in ChangeRequestType element
        changeRequest.ActionProperties = inputProperties;

        //  Build a list of properties to exclude on the new doc object that will be returned
        string[] propertyExclusions = new string[2] { "Owner", "DateLastModified" };

        //  Assign the list of excluded properties to the ChangeRequestType element
        changeRequest.RefreshFilter = Factory.PropertyFilterType(1, true, null, propertyExclusions);

        //  Create array of ChangeRequestType elements and assign ChangeRequestType element to it 
        ChangeRequestType[] changeRequestArray = new ChangeRequestType[1] { changeRequest };

        //  Create ChangeResponseType element array
        ChangeResponseType[] changeResponseArray = null;

        changeResponseArray = ExecuteChanges(changeRequest, true);

        //  The new document object should be returned, unless there is an error
        if (changeResponseArray == null)
        {
          errorMessage = errorMessage + "A valid object was not returned from the ExecuteChanges operation";
          return false;
        }

        if (changeResponseArray.Length < 1)
        {
          errorMessage = errorMessage + "A valid object was not returned from the ExecuteChanges operation";
          return false;
        }

        return true;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }


    public void FileDocument(string documentId, string documentClass, string documentName, string folderPath, [Optional][DefaultParameterValue("")] ref string errorMessage)
    {
      try
      {

        //  Create a ChangeRequest, populate it to create a new DRCR
        //  NOTE: This must be created as a DRCR, not a simple RCR, Otherwise the folder will always link to the exact version created here.
        CreateAction createAction = new CreateAction() { autoUniqueContainmentName = true, autoUniqueContainmentNameSpecified = true, classId = "DynamicReferentialContainmentRelationship" };
        ChangeRequestType changeRequest = new ChangeRequestType() { Action = new CreateAction[1] { createAction }, TargetSpecification = ObjectStoreReference };

        //  Create the properties of the new DRCR
        ObjectReference headReference = new ObjectReference() { classId = documentClass, objectId = documentId, objectStore = _objectStoreName };
        SingletonObject head = new SingletonObject() { propertyId = "Head", Value = headReference };

        ObjectReference tailReference = new ObjectReference() { classId = "Folder", objectId = folderPath, objectStore = _objectStoreName };
        SingletonObject tail = new SingletonObject() { propertyId = "Tail", Value = tailReference };

        //  Check to see if we happened to get a blank document name
        if (string.IsNullOrEmpty(documentName))
        {
          documentName = NO_TITLE_SPECIFIED;
          ApplicationLogging.LogWarning($"Set containment name for document '{documentId}' to '{NO_TITLE_SPECIFIED}; the document title was null or empty.", MethodBase.GetCurrentMethod());
        }

        SingletonString containmantName = new SingletonString() { propertyId = "ContainmentName", Value = CleanContainmentName(documentName) };
        
        ModifiablePropertyType[] objProps = new ModifiablePropertyType[3] {tail, head, containmantName };
        changeRequest.ActionProperties = objProps;

        try
        {
          //  Send off the request
          ChangeResponseType[] responseArrary = ExecuteChanges(changeRequest);
        }
          catch (Exception ex)
        {
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
          ////  Re-throw the exception to the caller
          //throw;
          if (ex.Message.Contains("access is denied"))
          {
            errorMessage = $"Access is denied for filing document '{documentId}' with Title '{documentName}' of class '{documentClass}' in folder '{folderPath}'";
          }
          else
          {
            errorMessage = $"An exception occurred while filing a document: [{ex.Message}]";
          }
            ApplicationLogging.WriteLogEntry(errorMessage, MethodBase.GetCurrentMethod(), TraceEventType.Error, 401);
            return;
        }

        errorMessage = "Successfully filed a document!";
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    public ECMProperty GetCtsProperty(CPEServiceReference.PropertyType ceProperty)
    {
      try
      {
        ECMProperty property = null;

        //  Determine the property type
        switch (ceProperty.GetType().Name)
        {
          case "SingletonString":
            {
              property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmString, ceProperty.propertyId, Cardinality.ecmSingleValued);
              property.Value = ((SingletonString)ceProperty).Value;
              break;
            }

          case "SingletonDateTime":
            {
              property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmDate, ceProperty.propertyId, Cardinality.ecmSingleValued);
              property.Value = ((SingletonDateTime)ceProperty).Value;
              break;
            }

          case "SingletonBoolean":
            {
              property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmBoolean, ceProperty.propertyId, Cardinality.ecmSingleValued);
              property.Value = ((SingletonBoolean)ceProperty).Value;
              break;
            }

          case "SingletonInteger32":
            {
              property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmLong, ceProperty.propertyId, Cardinality.ecmSingleValued);
              property.Value = ((SingletonInteger32)ceProperty).Value;
              break;
            }

          case "SingletonFloat64":
            {
              property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmDouble, ceProperty.propertyId, Cardinality.ecmSingleValued);
              property.Value = ((SingletonFloat64)ceProperty).Value;
              break;
            }

          case "SingletonObject":
            {
              property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmObject, ceProperty.propertyId, Cardinality.ecmSingleValued);
              property.Value = ((SingletonObject)ceProperty).Value;
              break;
            }

          case "SingletonId":
            {
              property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmGuid, ceProperty.propertyId, Cardinality.ecmSingleValued);
              property.Value = ((SingletonId)ceProperty).Value;
              break;
            }

          case "SingletonBinary":
            {
              property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmBinary, ceProperty.propertyId, Cardinality.ecmSingleValued);
              property.Value = ((SingletonBinary)ceProperty).Value;
              break;
            }

          case "ListOfString":
            {
              property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmString, ceProperty.propertyId, Cardinality.ecmMultiValued);
              foreach (var propertyValue in ((ListOfString)ceProperty).Value) { property.Values.Add(propertyValue); }
              break;
            }

          case "ListOfDateTime":
            {
              property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmDate, ceProperty.propertyId, Cardinality.ecmMultiValued);
              foreach (var propertyValue in ((ListOfDateTime)ceProperty).Value) { property.Values.Add(propertyValue); }
              break;
            }

          case "ListOfBoolean":
            {
              property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmBoolean, ceProperty.propertyId, Cardinality.ecmMultiValued);
              foreach (var propertyValue in ((ListOfBoolean)ceProperty).Value) { property.Values.Add(propertyValue); }
              break;
            }

          case "ListOfInteger32":
            {
              property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmLong, ceProperty.propertyId, Cardinality.ecmMultiValued);
              foreach (var propertyValue in ((ListOfInteger32)ceProperty).Value) { property.Values.Add(propertyValue); }
              break;
            }

          case "ListOfFloat64":
            {
              property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmDouble, ceProperty.propertyId, Cardinality.ecmMultiValued);
              foreach (var propertyValue in ((ListOfFloat64)ceProperty).Value) { property.Values.Add(propertyValue); }
              break;
            }

          case "ListOfObject":
            {
              property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmObject, ceProperty.propertyId, Cardinality.ecmMultiValued);
              foreach (var propertyValue in ((ListOfObject)ceProperty).Value) { property.Values.Add(propertyValue); }
              break;
            }

          case "ListOfId":
            {
              property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmGuid, ceProperty.propertyId, Cardinality.ecmMultiValued);
              foreach (var propertyValue in ((ListOfId)ceProperty).Value) { property.Values.Add(propertyValue); }
              break;
            }

          case "ListOfBinary":
            {
              property = (ECMProperty)PropertyFactory.Create(Core.PropertyType.ecmBinary, ceProperty.propertyId, Cardinality.ecmMultiValued);
              foreach (var propertyValue in ((ListOfBinary)ceProperty).Value) { property.Values.Add(propertyValue); }
              break;
            }

        }

        return property;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public string GetFolderID(string folderPath, ref string errorMessage)
    {
      try
      {
        ObjectValue folderValue = GetFolder(folderPath, 0, ref errorMessage);

        if (folderValue == null)
        {
          return string.Empty;
        }
        else
        {
          return folderValue.objectId;
        }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public ObjectValue GetFolder(string folderPath, int maxContentCount, [Optional][DefaultParameterValue("")] ref string errorMessage)
    {
      StringCollection folderPathCollection;
      string targetFolderPath = string.Empty;
      ObjectValue folderInfo = null;
      ObjectValue lastFolderInfo = null;
      string cleanFolderPath = string.Empty;

      try
      {


        //  Clean the folder path information to make sure that we get the valid folder names 
        //  in case there are invalid characters in the requested path

        folderPathCollection = PathFactory.CreateFolderPathCollection(folderPath);

        foreach (string folder in folderPathCollection)
        {
          targetFolderPath += $"/{folder}";
          folderInfo = GetFolderInfo(targetFolderPath, maxContentCount, ref errorMessage);
          if (folderInfo != null)
          {
            //  We found the folder
            lastFolderInfo = folderInfo;
          }
          else
          {
            //  We did not find the folder
            if (lastFolderInfo != null)
            {
              folderInfo = CreateFolder(folder, lastFolderInfo);
            }
            else
            {
              folderInfo = CreateFolder(folder);
            }
            if (folderInfo != null)
            {
              lastFolderInfo = folderInfo;
            }
            else
            {
              return null;
            }
          }
        }

        errorMessage = string.Empty;
        return folderInfo;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }



    public ObjectValue CreateFolder(string folderName, [Optional][DefaultParameterValue(null)] ObjectValue parentFolder)
    {
      try
      {
        //string newFolderName = string.Empty;
        //DateTime dateCreated = new DateTime();

        //  Build the Create Action for a folder object
        CreateAction verbCreate = new CreateAction();
        if (parentFolder != null)
        {
          verbCreate.classId = parentFolder.classId;
        }
        else
        {
          verbCreate.classId = "Folder";
        }

        //  Assign the action to the ChangeRequestType element and assign the Create action 
        ChangeRequestType elemChangeRequestType = new ChangeRequestType() { Action = new ActionType[1] { verbCreate }, TargetSpecification = ObjectStoreReference, id = "1" };

        //  Specify and set a string-valued property for the FolderName property
        SingletonString propContainmentName = new SingletonString() { propertyId = "FolderName", Value = folderName };

        ObjectReference objRootFolder;
        //  Do we have a value for 'parentFolder'?
        if (parentFolder == null)
        {
          //  No we don't
          //  Create an object reference to the root folder
          objRootFolder = new ObjectReference() { classId = "Folder", objectId = "{0F1E2D3C-4B5A-6978-8796-A5B4C3D2E1F0}", objectStore = _objectStoreName };
        }
        else
        {
          //  Yes we do
          //  Create an object reference to the parent folder
          objRootFolder = new ObjectReference() { classId = parentFolder.classId, objectId = parentFolder.objectId, objectStore = _objectStoreName };
        }

        //  Specify and set an object-valued property for the Parent property
        SingletonObject propParent = new SingletonObject() { propertyId = "Parent", Value = objRootFolder };


        //  Build a list of properties to set in the new doc and assign the contanment name and parent to them
        ModifiablePropertyType[] elemInputProps = new ModifiablePropertyType[2] { propContainmentName, propParent };

        //  Assign list of folder properties to set in ChangeRequestType element
        elemChangeRequestType.ActionProperties = elemInputProps;

        //  Build a list of properties to exclude on the new folder object that will be returned
        string[] excludeProps = new string[2] { "Owner", "DateLastModified" };

        //  Assign the list of excluded properties to the ChangeRequestType element
        elemChangeRequestType.RefreshFilter = new PropertyFilterType() { ExcludeProperties = excludeProps };

        ////  Create array of ChangeRequestType elements and assign ChangeRequestType element to it 
        //ChangeRequestType[] elemChangeRequestTypeArray = new ChangeRequestType[1] { elemChangeRequestType };

        //  Create ChangeResponseType element array
        ChangeResponseType[] elemChangeResponseTypeArray;

        try
        {
          //  Create ChangeResponseType element array
          elemChangeResponseTypeArray = ExecuteChanges(elemChangeRequestType);
        }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
          Debug.WriteLine($"An exception occurred while creating a folder: [{ex.Message}]");
          return null;
      }

        //  The new folder object should be returned, unless there is an error
        if ((elemChangeResponseTypeArray == null) || (elemChangeResponseTypeArray.Length < 1))
        {
          Debug.WriteLine("A valid object was not returned from the ExecuteChanges operation");
          return null;
        }

        ////  Capture value of the FolderName and DateCreated properties in the returned doc object
        //foreach (CPEServiceReference.PropertyType property in elemChangeResponseTypeArray[0].Property)
        //{
        //  //  If property found, store its value
        //  switch (property.propertyId)
        //  {
        //    case "FolderName":
        //      SingletonString folderNameProperty = (SingletonString)property;
        //      newFolderName = folderNameProperty.Value;
        //      break;

        //    case "DateCreated":
        //      SingletonDateTime dateCreatedProperty = (SingletonDateTime)property;
        //      dateCreated = dateCreatedProperty.Value;
        //      break;

        //    default:
        //      break;
        //  }          
        //}

        //Debug.WriteLine($"The folder {folderName} was successfully created {dateCreated}.");

        return elemChangeResponseTypeArray[0];

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    public ObjectValue GetFolderInfo([Optional][DefaultParameterValue("/")] string folderPath, [Optional][DefaultParameterValue(0)] int maxContentCount, [Optional][DefaultParameterValue("")] ref string errorMessage)
    {
      try
      {
        ObjectSpecification objSpec = new ObjectSpecification() { path = folderPath, classId = "Folder", objectStore = _objectStoreName };

        //  Retrieve the TopFolders
        ObjectRequestType objectRequest = new ObjectRequestType() { SourceSpecification = objSpec, id = "1", PropertyFilter = new PropertyFilterType() };

        //  If we're not retrieving content, then just get the doc title
        //Int32 filterElementCount = 9;
        Int32 filterElementCount = 16;
        FilterElementType[] incProps = new FilterElementType[filterElementCount];

        //  Id
        incProps[0] = new FilterElementType() { Value = "Name", maxRecursion = 1, maxRecursionSpecified = true };

        //  FolderName
        incProps[1] = new FilterElementType() { Value = "FolderName", maxRecursion = 1, maxRecursionSpecified = true };

        //  SubFolders
        incProps[2] = new FilterElementType() { Value = "SubFolders", maxRecursion = 1, maxRecursionSpecified = true };

        //  PathName
        incProps[3] = new FilterElementType() { Value = "PathName", maxRecursion = 1, maxRecursionSpecified = true };

        //  AllowedRMTypes
        incProps[4] = new FilterElementType() { Value = "AllowedRMTypes", maxRecursion = 1, maxRecursionSpecified = true };

        //  DateCreated
        incProps[5] = new FilterElementType() { Value = "DateCreated", maxRecursion = 1, maxRecursionSpecified = true };

        //  DateClosed
        incProps[6] = new FilterElementType() { Value = "DateClosed", maxRecursion = 1, maxRecursionSpecified = true };

        //  ReOpenedDate
        incProps[7] = new FilterElementType() { Value = "ReOpenedDate", maxRecursion = 1, maxRecursionSpecified = true };

        //  ContainedDocuments
        incProps[8] = new FilterElementType() { Value = "ContainedDocuments", maxRecursion = 1, maxRecursionSpecified = true };
        if (maxContentCount > -1)
        {
          incProps[8].maxElements = maxContentCount;
          incProps[8].maxElementsSpecified = true;
        }

        //  ContentElements
        incProps[9] = new FilterElementType() { Value = "ContentElements", maxRecursion = 2, maxRecursionSpecified = true };

        //  RetrievalName
        incProps[10] = new FilterElementType() { Value = "RetrievalName", maxRecursion = 3, maxRecursionSpecified = true };

        //  DateLastModified
        incProps[11] = new FilterElementType() { Value = "DateLastModified", maxRecursion = 2, maxRecursionSpecified = true };

        //  ContentSize
        incProps[12] = new FilterElementType() { Value = "ContentSize", maxRecursion = 2, maxRecursionSpecified = true };

        //  IsReserved
        incProps[13] = new FilterElementType() { Value = "IsReserved", maxRecursion = 2, maxRecursionSpecified = true };

        //  MajorVersionNumber
        incProps[14] = new FilterElementType() { Value = "MajorVersionNumber", maxRecursion = 2, maxRecursionSpecified = true };

        //  MinorVersionNumber
        incProps[15] = new FilterElementType() { Value = "MinorVersionNumber", maxRecursion = 2, maxRecursionSpecified = true };


        objectRequest.PropertyFilter.IncludeProperties = incProps;

        //  Create the request array
        ObjectRequestType[] objRequestArray = new ObjectRequestType[1] { objectRequest };

        //  Send off the request
        ObjectResponseType[] objResponseArray;

        try
        {
          objResponseArray = GetObjects(objRequestArray);
        }
        catch (Exception ex) 
        {
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
          if (ex.Message.StartsWith("Logon failure: unknown user name or bad password.")) { throw; }
          errorMessage = $"An exception occurred while querying for containees: [{ex.Message}]";
          return null;
        }

        if (objResponseArray[0].GetType().Name == "ErrorStackResponse")
        {
          ErrorStackResponse objErrResponse = (ErrorStackResponse)objResponseArray[0];
          ErrorStackType objStack = objErrResponse.ErrorStack;
          ErrorRecordType objErr = objStack.ErrorRecord[0];

          if (objErr.Description.Contains("requested item was not found"))
          {
            //  We did not find the folder
            errorMessage = $"No folder exists with the path '{folderPath}'";
          }
          else
          {
            errorMessage = $"Error [{objErr.Description}] occured.  Error source is [{objErr.Source}]";
            ApplicationLogging.WriteLogEntry($"Unable to get folder info in {errorMessage}: '{MethodBase.GetCurrentMethod().Name}'", TraceEventType.Error, 404);
          }
          return null;
        }

        //  Extract the folder object from the response
        ObjectValue folder;

        if (objResponseArray[0].GetType().Name == "SingleObjectResponse")
        {
          SingleObjectResponse objSingleObjectResponse = (SingleObjectResponse)objResponseArray[0];
          folder = objSingleObjectResponse.Object;
        }
        else
        {
          errorMessage = $"Unknown data type returned in ObjectReponse: [{objResponseArray[0].GetType()}] while querying for folder '{folderPath}'.";
          return null;
        }

        return folder;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }


    public void UpdateDocument(Document document, int versionIndex, bool updateAllChildren)
    {
      try
      {
        string ceVersionId = string.Empty;
        string blankValue = string.Empty;
        string errorMessage = string.Empty;

        if (updateAllChildren)
        {
          for (int versionCounter = versionIndex; versionIndex < document.Versions.Count; versionCounter++)
          {
            UpdateDocument(document, versionCounter, ref ceVersionId, ref errorMessage);
          }
        }
        else { UpdateDocument(document, versionIndex, ref blankValue, ref errorMessage);  }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public void UpdateDocument(Document document, int versionIndex, [Optional][DefaultParameterValue("")] ref string ceVersionId, [Optional][DefaultParameterValue("")] ref string errorMessage)
    {
      try
      {
        string logMessage = string.Empty;

        //  Build the Checkout Action
        CheckoutAction verbCheckout = Factory.CheckoutAction();

        //  Assign the action to the ChangeRequestType element
        ChangeRequestType elemChangeRequestType = new ChangeRequestType() { Action = new ActionType[1] { verbCheckout } };

        //  Set a reference to the document to check out
        ObjectSpecification objDocument = new ObjectSpecification() { classId = document.DocumentClass };

        if (ceVersionId.Length == 0)
        {
          objDocument.objectId = document.ObjectID;
        }
        else
        {
          objDocument.objectId = ceVersionId;
        }

        objDocument.objectStore = _objectStoreName;

        //  Specify the target object (a document) for the actions
        elemChangeRequestType.TargetSpecification = objDocument;
        elemChangeRequestType.id = "1";

        //  Create a Property Filter to get Reservation property
        PropertyFilterType elemPropFilter = Factory.PropertyFilterType(1, true);
        elemPropFilter.IncludeProperties = new FilterElementType[1] { new FilterElementType() { Value = "Reservation" } };

        //  Assign the list of included properties to the ChangeRequestType element
        elemChangeRequestType.RefreshFilter = elemPropFilter;

        // Create array of ChangeRequestType elements and assign ChangeRequestType element to it
        ChangeRequestType[] elemChangeRequestTypeArray = new ChangeRequestType[1] { elemChangeRequestType };

        //  Create ChangeResponseType element array 
        ChangeResponseType[] elemChangeResponseTypeArray = null;

        //  Build ExecuteChangesRequest element and assign ChangeRequestType element array to it
        ExecuteChangesRequest elemExecuteChangesRequest = new ExecuteChangesRequest() { ChangeRequest = elemChangeRequestTypeArray, refresh = true, refreshSpecified = true };

        try
        {
          //  
          Task<ExecuteChangesResponse> response = _client.ExecuteChangesAsync(_localization, elemExecuteChangesRequest);
          response.Wait();

          elemChangeResponseTypeArray = response.Result.ExecuteChangesResponse1;

          if (elemChangeResponseTypeArray == null)
          {
            errorMessage = "ExecuteChanges returned nothing...";
            return;
          }
        }
        catch (Exception ex)
        {
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());

          if (ex.Message.StartsWith(CE_ERR_MSG_OBJECT_MODIFIED))
          {
            //  Let's try it again...
            UpdateDocument(document, versionIndex, ref ceVersionId, ref errorMessage);
          }
          if (ex.Message.StartsWith(CE_ERR_MSG_VERSION_RESERVED)) { return; }
          errorMessage = $"An exception occurred while checking out a document: [{ex.Message}]: Document ID:{document.ID}, Version ID:{ceVersionId}, Version Index:{versionIndex}";
          ApplicationLogging.WriteLogEntry(errorMessage, MethodBase.GetCurrentMethod(), TraceEventType.Error, 201);
          return;
        }

        //  For some reason we occasionally get a null reference at this point.
        //  Even though we check for this above.
        if (elemChangeResponseTypeArray == null)
        {
          errorMessage = "ExecuteChanges returned nothing...";
            return;
        }

        //  Get Reservation object from Reservation property
        SingletonObject propReservation = (SingletonObject)GetPropertyByName("Reservation", elemChangeResponseTypeArray[0]);
        ObjectValue objReservation = (ObjectValue)propReservation.Value;

        //  Build the Checkin action
        CheckinAction verbCheckin = new CheckinAction();

        //  See if this should be a major version
        if (IsMajorVersion(document.Versions[versionIndex], true))
        {
          ApplicationLogging.WriteLogEntry($"Document '{document.ID}' will be checked in as a major version", MethodBase.GetCurrentMethod(), TraceEventType.Verbose, 301);
          verbCheckin = Factory.CheckinAction(false, false);
        }
        else
        {
          ApplicationLogging.WriteLogEntry($"Document '{document.ID}' will be checked in as a minor version", MethodBase.GetCurrentMethod(), TraceEventType.Verbose, 301);
          verbCheckin = Factory.CheckinAction(true, true);
        }

        //  Assign the action to the ChangeRequestType element
        elemChangeRequestType.Action = new ActionType[1] { verbCheckin };

        //  Build a list of properties to set in the new doc
        ModifiablePropertyType[] elemInputProps;
        elemInputProps = BuildPropertyList(document.Versions[versionIndex]);

        //  Build the ContentElement List
        //BuildContentElementList(document.Versions[versionIndex], ref elemInputProps[document.Versions[versionIndex].Properties.Count]);
        BuildContentElementList(document.Versions[versionIndex], ref elemInputProps);

        //  Assign list of document properties to set in ChangeRequestType element
        elemChangeRequestType.ActionProperties = elemInputProps;

        //  Specify the target object (Reservation object) for the actions
        elemChangeRequestType.TargetSpecification = new ObjectReference() { classId = document.DocumentClass, objectId = objReservation.objectId, objectStore = _objectStoreName };
        elemChangeRequestType.id = "1";

        //  Assign ChangeRequestType element
        elemChangeRequestTypeArray[0] = elemChangeRequestType;

        // Build ExecuteChangesRequest element and assign ChangeRequestType element array to it
        elemExecuteChangesRequest.ChangeRequest = elemChangeRequestTypeArray;
        elemExecuteChangesRequest.refresh = true;
        elemExecuteChangesRequest.refreshSpecified = true;

      try
      {
          //  Call ExecuteChanges operation to implement the doc checkout
          Task<ExecuteChangesResponse> response = _client.ExecuteChangesAsync(_localization, elemExecuteChangesRequest);
          response.Wait();

          elemChangeResponseTypeArray = response.Result.ExecuteChangesResponse1;

          if (elemChangeResponseTypeArray == null)
          {
            errorMessage = "ExecuteChanges returned nothing...";
            return;
          }
        }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());

          if (ex.Message.StartsWith(CE_ERR_MSG_OBJECT_MODIFIED))
          {
            //  Let's try again...
            UpdateDocument(document, versionIndex, ref ceVersionId, ref errorMessage);           
          }

          if (ex.Message.StartsWith(CE_ERR_MSG_UNDERLYING_CONNECTION_WAS_CLOSED))
          {
            errorMessage = $"{CE_ERR_MSG_UNDERLYING_CONNECTION_WAS_CLOSED}, check the file size: [{ex.Message}]";
            return;
          }
          else
          {
            errorMessage = $"An exception occurred while checking out a document: [{ex.Message}]";
            return;
          }

        }

        //  The new document object should be returned, unless there is an error
        if ((elemChangeResponseTypeArray == null) || (elemChangeResponseTypeArray.Length < 1))
        {
          errorMessage = "A valid object was not returned from the ExecuteChanges operation";
          return;
        }

        //  Capture value of the 'Id' property in the returned doc object
        ceVersionId = elemChangeResponseTypeArray[0].objectId;

        if (document.Versions[versionIndex].Contents.Count > 0)
        {
          logMessage = $"Checked-in Document '{document.Versions[versionIndex].Contents[0].ContentPath}'.";
        }
        else
        {
          logMessage = $"Checked-in Document '{document.ID}' Version '{versionIndex}' without content.";
        }
        ApplicationLogging.LogInformation(logMessage, MethodBase.GetCurrentMethod());
        Debug.WriteLine(logMessage);

      }
      catch (Exception ex)
        {
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
          //  Re - throw the exception to the caller
          throw;
        }
    }

    public static CPEServiceReference.PropertyType GetPropertyByID(ObjectValue objectValue, string propertyId)
    {
      try
      {
        if ((objectValue != null) && (objectValue.Property != null))
        {
          foreach (CPEServiceReference.PropertyType propertyType in objectValue.Property)
          {
            if (propertyType.propertyId == propertyId)
            {
              return propertyType;
            }
          }

          //  We did not find it
          return null;
          }       
        else 
        {
          //  We did not recieve a valid ObjectValue
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

    public static CPEServiceReference.PropertyType GetPropertyByName(string propertyName, ObjectValue item)
    {
      try
      {
        if (item == null) { throw new ArgumentNullException(nameof(item)); }
        if (item.Property == null) { throw new ArgumentNullException(nameof(item), "item.Property is null."); }

        foreach (var property in item.Property) { if (property.propertyId == propertyName) { return property; } }

        return null;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public static string GetRetrievalName(DependentObjectType contentElement)
    {
      try
      {
        //  Get the Content "Retrieval Name"s
        SingletonString retrievalName = (SingletonString)GetPropertyByID((ObjectValue)contentElement, "RetrievalName");

        if (retrievalName != null)
        {
          return retrievalName.Value;
        }
        else { return string.Empty; }

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public DocumentClasses GetAllDocumentClassDefinitions(List<string> excludeProperties = null)
    {
      try
      {

        DocumentClasses documentClasses = new DocumentClasses();

        string sql = "SELECT [Id] FROM [DocumentClassDefinition]";

        Debug.WriteLine("Performing an ExecuteSearch, SQL:");
        Debug.WriteLine($" {sql}");

        ObjectSetType objectSet = ExecuteSearch(sql, 500);

        if (objectSet != null)
        {
          ClassificationProperties properties = new ClassificationProperties();
          //ClassificationProperty property;
          ObjectValue objectValue;
          ModifiablePropertyType cewsProperty;

          DocumentClass documentClass;

          for (int templateCounter = 0; templateCounter < objectSet.Object.Length - 1; templateCounter++)
          {
            objectValue = objectSet.Object[templateCounter];
            for (int propertyCounter = 0; propertyCounter < objectValue.Property.Length; propertyCounter++)
            {
              cewsProperty = (ModifiablePropertyType)objectValue.Property[propertyCounter];
              if (cewsProperty.propertyId == "Id")
              {
                SingletonId idProperty = (SingletonId)cewsProperty;
                string id = idProperty.Value;
                documentClass = GetClassDefinition(id, excludeProperties, false);
                documentClasses.Add(documentClass);
                break;
              }
            }
          }
         return documentClasses;
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

    public ClassificationProperties GetAllPropertyTemplates([Optional][DefaultParameterValue("")] ref string errorMessage)
    {
      try
      {
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.Append("SELECT [This], [Cardinality], [DataType], [Id], ");
        stringBuilder.Append("[IsHidden], [IsValueRequired], [SymbolicName], ");
        stringBuilder.Append("[PropertyDefaultBoolean], [PropertyDefaultDateTime], ");
        stringBuilder.Append("[PropertyDefaultFloat64], [PropertyDefaultId], ");
        stringBuilder.Append("[PropertyDefaultInteger32],  [PropertyMinimumInteger32], ");
        stringBuilder.Append("[PropertyMaximumInteger32], [PropertyDefaultString], ");
        stringBuilder.Append("[PropertyMinimumDateTime], [PropertyMaximumDateTime], ");
        stringBuilder.Append("[PropertyMinimumFloat64], [PropertyMaximumFloat64], ");
        stringBuilder.Append("[Settability], [MaximumLengthString] FROM [PropertyTemplate]");

        ObjectSetType objectSet = ExecuteSearch(stringBuilder.ToString());

        if (objectSet != null)
        {
          ClassificationProperties classificationProperties = new ClassificationProperties();
          ClassificationProperty classificationProperty = null;
          ObjectValue value = null;

          for (int propertyTemplateCounter = 0; propertyTemplateCounter < objectSet.Object.Length; propertyTemplateCounter++)
          {
            value = (ObjectValue)objectSet.Object.GetValue(propertyTemplateCounter);
            classificationProperty = GetClassificationProperty(value);

            if (_contentExportPropertyExclusions == null)
            {
              classificationProperties.Add(classificationProperty);
            }
            else
            {
              if (!_contentExportPropertyExclusions.Contains(classificationProperty.PackedName))
              {
                classificationProperties.Add(classificationProperty);
              }
            }
          }

          // We need to make sure we also have Id
          if (!classificationProperties.Contains("Id"))
          {
            ClassificationProperty idProperty = (ClassificationProperty)ClassificationPropertyFactory.Create(Core.PropertyType.ecmGuid, "Id", Cardinality.ecmSingleValued);
            idProperty.IsSystemProperty = true;
            idProperty.IsRequired = true;
            idProperty.Settability = ClassificationProperty.SettabilityEnum.SETTABLE_ONLY_ON_CREATE;
            classificationProperties.Add(idProperty);
          }

          // Hard code an add for Rick
          if (!classificationProperties.Contains("IsCurrentVersion"))
          {
            ClassificationProperty isCurrentProperty = (ClassificationProperty)ClassificationPropertyFactory.Create(Core.PropertyType.ecmBoolean, "IsCurrentVersion", Cardinality.ecmSingleValued);
            isCurrentProperty.IsSystemProperty = true;
            isCurrentProperty.IsRequired = true;
            isCurrentProperty.Settability = ClassificationProperty.SettabilityEnum.READ_ONLY;
            classificationProperties.Add(isCurrentProperty);
          }

          return classificationProperties;
        }
        else { return null; }

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        errorMessage = Helper.FormatCallStack(ex);
        return null;
      }
    }

    public DocumentClass GetClassDefinition(string id, List<string> excludeProperties, bool includeSuperClassDefinitions)
    {
      try
      {
        DocumentClass documentClass;

        ObjectReference objectReference = Factory.ObjectReference("DocumentClassDefinition", id, _objectStoreName);
        PropertyFilterType propertyFilter = Factory.PropertyFilterType(0);
        FilterElementType[] includeProperties = Array.Empty<FilterElementType>();

        FilterElementType includeFilterTypes0 = new FilterElementType() { Value = "SingletonObject EnumOfObject", maxRecursion = 0, maxRecursionSpecified = true };
        FilterElementType includeFilterTypes1 = new FilterElementType() { Value = "Singleton* List*", maxRecursion = 3, maxRecursionSpecified = true };

        int filterElementCount = 24;
        if (!includeSuperClassDefinitions) { filterElementCount -= 1; }

        propertyFilter.IncludeProperties = new FilterElementType[filterElementCount];
        propertyFilter.IncludeProperties[0] = Factory.FilterElementType("Id", 5);
        propertyFilter.IncludeProperties[1] = Factory.FilterElementType("PropertyDefinitions", 5);
        propertyFilter.IncludeProperties[2] = Factory.FilterElementType("PropertyDefinition", 5);
        propertyFilter.IncludeProperties[3] = Factory.FilterElementType("Property", 5);
        propertyFilter.IncludeProperties[4] = Factory.FilterElementType("objectId", 5);
        propertyFilter.IncludeProperties[5] = Factory.FilterElementType("Name", 5);
        propertyFilter.IncludeProperties[6] = Factory.FilterElementType("DisplayName", 5);
        propertyFilter.IncludeProperties[7] = Factory.FilterElementType("DescriptiveText", 5);
        propertyFilter.IncludeProperties[8] = Factory.FilterElementType("SymbolicName", 5);
        propertyFilter.IncludeProperties[9] = Factory.FilterElementType("ChoiceList", 5);
        propertyFilter.IncludeProperties[10] = Factory.FilterElementType("MaximumLengthString", 5);
        propertyFilter.IncludeProperties[11] = Factory.FilterElementType("PropertyDefaultString", 5);
        propertyFilter.IncludeProperties[12] = Factory.FilterElementType("PropertyDefaultBoolean", 5);
        propertyFilter.IncludeProperties[13] = Factory.FilterElementType("PropertyDefaultDateTime", 5);
        propertyFilter.IncludeProperties[14] = Factory.FilterElementType("PropertyDefaultFloat64", 5);
        propertyFilter.IncludeProperties[15] = Factory.FilterElementType("PropertyDefaultId", 5);
        propertyFilter.IncludeProperties[16] = Factory.FilterElementType("PropertyDefaultInteger32", 5);
        propertyFilter.IncludeProperties[17] = Factory.FilterElementType("PropertyMinimumInteger32", 5);
        propertyFilter.IncludeProperties[18] = Factory.FilterElementType("PropertyMaximumInteger32", 5);
        propertyFilter.IncludeProperties[19] = Factory.FilterElementType("PropertyMinimumDateTime", 5);
        propertyFilter.IncludeProperties[20] = Factory.FilterElementType("PropertyMaximumDateTime", 5);
        propertyFilter.IncludeProperties[21] = Factory.FilterElementType("PropertyMinimumFloat64", 5);
        propertyFilter.IncludeProperties[22] = Factory.FilterElementType("PropertyMaximumFloat64", 5);

        if (includeSuperClassDefinitions) 
        {
          propertyFilter.IncludeProperties[23] = Factory.FilterElementType("SuperclassDefinition", 5);
        }

        propertyFilter.IncludeTypes = new FilterElementType[2] { includeFilterTypes0, includeFilterTypes1 };

        if (excludeProperties != null) { propertyFilter.ExcludeProperties = excludeProperties.ToArray(); }
        ObjectRequestType[] objectRequest = new ObjectRequestType[1] { new ObjectRequestType() { SourceSpecification = objectReference, PropertyFilter = propertyFilter } };

        GetObjectsResponse objectResponse;

        objectResponse = _client.GetObjectsAsync(_localization, objectRequest).Result;


        CPEServiceReference.PropertyType propertyType;
        SingletonId idPropertyType;
        SingletonString stringPropertyType;


        DependentObjectType propertyDefinitionStub;

        ClassificationProperties classificationProperties = new ClassificationProperties();
        ClassificationProperty classificationProperty;

        string classId = string.Empty;
        string className = string.Empty;
        string classSymbolicName = string.Empty;

        SingleObjectResponse firstObjectResponse = (SingleObjectResponse)objectResponse.GetObjectsResponse1[0];

        for (int propertyCounter = 0; propertyCounter < firstObjectResponse.Object.Property.Length; propertyCounter++)
        {
          propertyType = firstObjectResponse.Object.Property[propertyCounter];
          Debug.Print($"PropertyCounter: {propertyCounter} {propertyType.propertyId}");

          switch (propertyType.propertyId)
          {
            case "Id":
              idPropertyType = (SingletonId)propertyType;
              classId = idPropertyType.Value.ToString();
              break;
            
            case "Name":
              stringPropertyType = (SingletonString)propertyType;
              className = stringPropertyType.Value;
              break;
            
            case "SymbolicName":
              stringPropertyType = (SingletonString)propertyType;
              classSymbolicName = stringPropertyType.Value;
              break;

            case "PropertyDefinitions":
              ListOfObject propertyTypes = (ListOfObject)propertyType;
              for (int propertyDefinitionCounter = 0; propertyDefinitionCounter < propertyTypes.Value.Length; propertyDefinitionCounter++)
              {
                propertyDefinitionStub = propertyTypes.Value[propertyDefinitionCounter];
                classificationProperty = GetClassificationProperty(propertyDefinitionStub);

                if (excludeProperties == null)
                {
                  classificationProperties.Add(classificationProperty);
                }
                else
                {
                  if (!excludeProperties.Contains(classificationProperty.PackedName))
                  {
                    classificationProperties.Add(classificationProperty);
                  }
                }
              }
              break;
                  
            default:
              break;
          }
        }

        documentClass = new DocumentClass(classSymbolicName, classificationProperties, classId, className);

        return documentClass;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public DocumentClass GetDocumentClassDefinition(string documentClassName)
    {
      try
      {
        string sql = "SELECT [Id] FROM [DocumentClassDefinition] WHERE [SymbolicName] = '" + documentClassName + "'";

        ObjectSetType objectSet = ExecuteSearch(sql);

        if (objectSet != null)
        {
          if (objectSet.Object != null)
          {
            ClassificationProperties contentProperties = new ClassificationProperties();
            //ClassificationProperty classificationProperty;
            ObjectValue objectValue;
            ModifiablePropertyType cewsProperty;
            DocumentClass classDefinition;

            for (int propertyTemplateCounter = 0; propertyTemplateCounter < objectSet.Object.Length; propertyTemplateCounter++)
            {
              objectValue = objectSet.Object[propertyTemplateCounter];
              for (int propertyTemplatePropertyCounter = 0; propertyTemplatePropertyCounter < objectValue.Property.Length; propertyTemplatePropertyCounter++)
              {
                cewsProperty = (ModifiablePropertyType)objectValue.Property[propertyTemplatePropertyCounter];
                if (cewsProperty.propertyId == "Id")
                {
                  SingletonId idProperty = (SingletonId)cewsProperty;
                  string id = idProperty.Value;
                  classDefinition = GetClassDefinition(id, _contentExportPropertyExclusions, false);
                  return classDefinition;
                }
              }
            }

          }
        }
        return null;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public Values GetDocumentIDs(ISearch search)
    {
      try
      {
        Values values = new Values();
        string error = string.Empty;

        //  Specify the scope of the search
        ObjectSetType searchResult = ExecuteSearch(search.DataSource.SQLStatement);

        int hitCount = 0;
        if ((searchResult != null) && (searchResult.Object != null)) { hitCount = searchResult.Object.Length; }

        for (int i = 0; i < hitCount; i++)
        {
          SingletonId propDocumentId = (SingletonId)searchResult.Object[i].Property[0];
          values.Add(propDocumentId.Value);
        }

        return values;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public SearchResultSet GetDocumentIDSet(ISearch search)
    {
      SearchResults results = new SearchResults();
      try
      {
        Values values = GetDocumentIDs(search);

        foreach (string value in values.Cast<string>())
        {
          results.Add(new SearchResult(value));
        }

        return new SearchResultSet(results);

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        return new SearchResultSet(results, ex);
      }
    }

    public ObjectValue GetDocumentProperties(string id, RequestType requestType, string[] excludeProperties, bool exportContent = true, bool exportAnnotations = true, string classId = "Document")
    {
      try
      {

        ObjectRequestType request = CreateDocumentRequest(id, requestType, ref excludeProperties, ref exportContent, ref exportAnnotations, classId);

        //  Create the request array
        ObjectRequestType[] requestArray = new ObjectRequestType[1] { request };

        //  Send off the request
        ObjectResponseType[] responses = GetObjects(requestArray);

        //  Did we get a document back?
        if (responses.Length < 1)
        {
          throw new DocumentException(id, $"No document found for ID '{id}'.");
        }

        //  Return the Document
        return ((SingleObjectResponse)responses[0]).Object;

      }
      catch (OutOfMemoryException memEx)
      {
        ApplicationLogging.LogException(memEx, MethodBase.GetCurrentMethod());
        throw new ContentTooLargeException("Unknown", $"Content is too large to retrieve for document with id '{id}'", memEx);
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        throw new Exception($"An error occurred while retrieving document '{id}'", ex);
      }
    }

    public static string[] GetFoldersFiledIn(EnumOfObject folderEnum)
    {
      try
      {
        SingletonString pathName;
        string folderPath;
        string[] folders = Array.Empty<string>();
        int folderCounter = 0;

        if ((folderEnum != null) && (folderEnum.Value != null))
        {
          foreach (ObjectValue folder in folderEnum.Value)
          {
            try
            {
              folderCounter++;
              pathName = (SingletonString)GetPropertyByID(folder, "PathName");
              if (pathName == null)
              {
                //  We could not find the path name for this folder
                continue;
              }
              folderPath = pathName.Value;
              Array.Resize(ref folders, folderCounter);
              folders[folderCounter - 1] = folderPath;
            }
            catch (Exception ex)
            {
              ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
              //  Skip it
            }
          }
          return folders;
        }
        return null;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }


    public SearchResultSet ExecuteSearch(ISearch search)
    {
      try
      {

        if (search == null) { throw new ArgumentNullException(nameof(search)); }
        if (search.DataSource == null) { throw new ArgumentException("The search datasource is not initialized", nameof(search)); }
        if (String.IsNullOrEmpty(search.DataSource.SQLStatement)) { throw new ArgumentException("The SQL statement is not initialized", nameof(search)); }

        ObjectSetType objectSet = ExecuteSearch(search.DataSource.SQLStatement);

        return BuildSearchResultSet(objectSet);

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public bool GetHasPriviledgedAccess()
    {
      try
      {
        if (ObjectStoreReference == null) { throw new RepositoryNotConnectedException("No object store reference available"); }

        string errorMessage = string.Empty;
        bool hasPriviledgedAccess = false;
        string grantee = string.Empty;

        ListOfObject objectStoreACEs = GetObjectStoreACEs();

        foreach (DependentObjectType accessPermission in objectStoreACEs.Value)
        {
          foreach (CPEServiceReference.PropertyType property in accessPermission.Property)
          {
          //  See if this is the current logged on user.
          //  Problem: What if the priveleged access is granted to a group.  
          //  Just skip this user specific check for now and assume that if 
          //  anyone has priveleged access that we can proceed.
          switch (property.propertyId)
            {
              case "GranteeName":
                {
                  SingletonString propertyObject = (SingletonString)property;
                  if (propertyObject.Value.ToString().Contains(_provider.UserName))
                  {
                    grantee = propertyObject.Value.ToString();
                  }
                  break;
                }

              case "AccessMask":
                {
                  SingletonInteger32 propertyInt = (SingletonInteger32)property;
                  if ((PRIVILEGED_WRITE_AS_INT & propertyInt.Value) != 0)
                    hasPriviledgedAccess = true;
                  break;
                }
            }
          }
          if (hasPriviledgedAccess)
          {
            ApplicationLogging.WriteLogEntry(String.Format("Grantee '{0}' has priveleged access.", grantee), TraceEventType.Information, 65123);
            return true;
          }
        }

        return false;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public ObjectValue GetObject(string id, string classId)
    {
      try
      {
        SingleObjectResponse singleObjectResponse;
        ObjectSpecification objSpec = new ObjectSpecification();

        //  Check to make sure we got an id
        if (id == null) { throw new ArgumentNullException(nameof(id)); }
        if (id.Length == 0) { throw new ArgumentException("The specified ID is a zero length string.  Please provide a valid document identifier.", nameof(id)); }

        objSpec.objectId = id;
        objSpec.classId = classId;

        switch (classId.ToLower())
        {
          case "objectstore":
          case "domain":
          case "entirenetwork":
            {
              //  Do nothing
              break;
            }

          default:
            {
              objSpec.objectStore = _objectStoreName;
              break;
            }
        }

        ObjectRequestType request = new ObjectRequestType() { SourceSpecification = objSpec, id = "1" };

        //  Create the request array
        ObjectRequestType[] requestArray = new ObjectRequestType[1] { request };

        //  Send off the request
        ObjectResponseType[] responses = GetObjects(requestArray);  

        if (responses[0].GetType().Name == "ErrorStackResponse")
        {
          ErrorStackResponse errResp = (ErrorStackResponse)responses[0];
          ErrorStackType objStack = errResp.ErrorStack;
          ErrorRecordType objErr = objStack.ErrorRecord[0];
          ApplicationLogging.WriteLogEntry($"Error [{objErr.Description}] occurred.  Error source is [{objErr.Source}]", MethodBase.GetCurrentMethod(),TraceEventType.Error, 9348);
          return null;
        }

        //  Extract the object from the response
        if (responses[0].GetType().Name == "SingleObjectResponse")
        {
          singleObjectResponse = (SingleObjectResponse)responses[0];
        }
        else
        {
          ApplicationLogging.WriteLogEntry($"Unknown data type returned in ObjectReponse: [{responses[0].GetType()}] while querying for object '{id}'", MethodBase.GetCurrentMethod(), TraceEventType.Error, 9349);
          return null;
        }

        return singleObjectResponse.Object;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    #endregion

    #region Private Methods

    private ChangeResponseType[] ExecuteChanges(ChangeRequestType changeRequest, bool refresh = true)
    {
      try
      {
        //  Create array of ChangeRequestType elements and assign ChangeRequestType element to it 
        ChangeRequestType[] elemChangeRequestTypeArray = new ChangeRequestType[1] { changeRequest };
        return ExecuteChanges(elemChangeRequestTypeArray, refresh);
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    private ChangeResponseType[] ExecuteChanges(ChangeRequestType[] changesRequest, bool refresh = true)
    {
      try
      {
        ExecuteChangesRequest executeChangesRequest;

        //  Build ExecuteChangesRequest element and assign ChangeRequestType element array to it
        if (refresh)
        {
          executeChangesRequest = new ExecuteChangesRequest() { ChangeRequest = changesRequest, refresh = true, refreshSpecified = true };
        }
        else
        {
          executeChangesRequest = new ExecuteChangesRequest() { ChangeRequest = changesRequest };
        }

        Task<ExecuteChangesResponse> response = _client.ExecuteChangesAsync(_localization, executeChangesRequest);
        response.Wait();

        return response.Result.ExecuteChangesResponse1;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    public Contents GetContents(string versionId, ListOfObject contentElements)//, ObjectSpecification objDocumentSpec)
    {

      try
      {
        NamedStream namedStream = null;
        Content content = null;
        Contents contents = new Contents();
        string contentRetrievalName;
        GetContentResponse response;
        // Get number of content elements
        int intElementCount = (contentElements.Value == null) ? 0 : contentElements.Value.Length;
        if (intElementCount == 0)
        {
          Console.WriteLine("The selected document has no content elements");
          Console.WriteLine("Press Enter to end");
          Console.ReadLine();
          return null;
        }

        // Set a reference to the document to retrieve
        ObjectSpecification objDocumentSpec = new ObjectSpecification();
        objDocumentSpec.classId = "Document";
        objDocumentSpec.objectId = versionId; //.path = strDocPath;
        objDocumentSpec.objectStore = _objectStoreName;


        // Get the content from each content element of the document
        for (int intElem = 0; intElem < intElementCount; intElem++)
        {
          // Get a ContentTransfer object from the ContentElements property collection
          DependentObjectType objContentTransfer = contentElements.Value[intElem];

          //  Get the Retrieval Name
          contentRetrievalName = GetRetrievalName(contentElements.Value[intElem]);

          // Construct element specification for GetContent request
          ElementSpecificationType objElemSpecType = new ElementSpecificationType();
          objElemSpecType.itemIndex = intElem;
          objElemSpecType.itemIndexSpecified = true;
          objElemSpecType.elementSequenceNumber = 0;
          objElemSpecType.elementSequenceNumberSpecified = false;

          // Construct the GetContent request
          ContentRequestType objContentReqType = new ContentRequestType();
          objContentReqType.cacheAllowed = true;
          objContentReqType.cacheAllowedSpecified = true;
          objContentReqType.id = "1";
          objContentReqType.maxBytes = 100 * 1024;
          objContentReqType.maxBytesSpecified = true;
          objContentReqType.startOffset = 0;
          objContentReqType.startOffsetSpecified = true;
          objContentReqType.continueFrom = null;
          objContentReqType.ElementSpecification = objElemSpecType;
          objContentReqType.SourceSpecification = objDocumentSpec;
          ContentRequestType[] objContentReqTypeArray = new ContentRequestType[1];
          objContentReqTypeArray[0] = objContentReqType;
          GetContentRequest objGetContentReq = new GetContentRequest();
          objGetContentReq.ContentRequest = objContentReqTypeArray;
          objGetContentReq.validateOnly = false;
          objGetContentReq.validateOnlySpecified = true;

          // Call the GetContent operation
          ContentResponseType[] objContentRespTypeArray = null;
          try
          {
            //objContentRespTypeArray = wseService.GetContent(objGetContentReq);
            response = GetContent(objGetContentReq);
          }
          catch (System.Net.WebException ex)
          {
            //Console.WriteLine("An exception occurred while fetching content from a content element: [" + ex.Message + "]");
            //Console.WriteLine("Press Enter to end");
            Console.ReadLine();
            return null;
          }
          catch (Exception allEx)
          {
            //Console.WriteLine("An exception occurred: [" + allEx.Message + "]");
            //Console.WriteLine("Press Enter to end");
            //Console.ReadLine();
            return null;
          }

          // Process GetContent response
          ContentResponseType objContentRespType = response.GetContentResponse1[0];
          if (objContentRespType is ContentErrorResponse)
          {
            ContentErrorResponse objContentErrorResp = (ContentErrorResponse)objContentRespType;
            ErrorStackType objErrorStackType = objContentErrorResp.ErrorStack;
            ErrorRecordType objErrorRecordType = objErrorStackType.ErrorRecord[0];
            //Console.WriteLine("Error [" + objErrorRecordType.Description + "] occurred. " + " Err source is [" + objErrorRecordType.Source + "]");
            //Console.WriteLine("Press Enter to end");
            //Console.ReadLine();
            return null;
          }
          else if (objContentRespType is ContentElementResponse)
          {
            ContentElementResponse objContentElemResp = (ContentElementResponse)objContentRespType;
            InlineContent objInlineContent = (InlineContent)objContentElemResp.Content;

            // Write content to file
            //Stream outputStream = File.OpenWrite(strDocContentPath);
            MemoryStream outputStream = new MemoryStream();
            outputStream.Write(objInlineContent.Binary, 0, objInlineContent.Binary.Length);
            //outputStream.Close();
            //Console.WriteLine("Document content has been written");
            //Console.WriteLine("Press Enter to end");
            //Console.ReadLine();

            namedStream = new NamedStream(outputStream, contentRetrievalName);
            content = new Content(namedStream);
            contents.Add(content);

          }
          else
          {
            //Console.WriteLine("Unknown data type returned in content response: [" + objContentRespType.GetType().ToString() + "]");
            //Console.WriteLine("Press Enter to end");
            //Console.ReadLine();
            return null;
          }
        }
        return contents;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private static void BuildContentElementList(Core.Version version, ref ModifiablePropertyType[] inputProperties)
    {
      try
      {
        int contentCounter = 0;

        //  Create an object reference to dependently persistable ContentTransfer object
        DependentObjectType contentTransfer;

        //  Create reference to the object set of ContentTransfer objects returned by the Document.ContentElements property
        ListOfObject contentElements = Factory.ListOfObject("ContentElements", true, true);

        //  Create the array to hold the content objects
        contentElements.Value = new DependentObjectType[version.Contents.Count];

        //  Loop through all available content elements
        foreach (Content content in version.Contents)
        {
          //  Create an object reference to dependently persistable ContentTransfer object
          contentTransfer = Factory.DependentObjectType("ContentTransfer", DependentObjectTypeDependentAction.Insert, true);
          contentTransfer.Property = new CPEServiceReference.PropertyType[3];

          contentElements.Value[contentCounter] = contentTransfer;

          Stream attachmentStream = content.ToStream();
          InlineContent inlineContent = new InlineContent();

          inlineContent.Binary = new byte[attachmentStream.Length];
          attachmentStream.Read(inlineContent.Binary, 0, (int)attachmentStream.Length);
          attachmentStream.Close();

          // Create reference to Content pseudo-property
          CPEServiceReference.ContentData contentData = new CPEServiceReference.ContentData();
          contentData.Value = (ContentType)inlineContent;
          contentData.propertyId = "Content";

          // Assign Content property to ContentTransfer object
          contentTransfer.Property[0] = contentData;

          //  Create and assign ContentType string-valued property to ContentTransfer object
          SingletonString contentTypeProperty = Factory.SingletonString("ContentType", content.MIMEType, true);
          contentTransfer.Property[1] = contentTypeProperty;

          //  Create and assign RetrievalName string-valued property to ContentTransfer object
          SingletonString retrievalNameProperty = Factory.SingletonString("RetrievalName", content.FileName, true);
          contentTransfer.Property[2] = retrievalNameProperty;

          contentCounter++;

        }

        Array.Resize(ref inputProperties, inputProperties.Length + 1);

        inputProperties[inputProperties.Length - 1] = contentElements;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private ObjectResponseType[] GetObjects(ObjectRequestType[] request)
    {
      try
      {
        Task<GetObjectsResponse> objectsResponse = _client.GetObjectsAsync(_localization, request);
        if (objectsResponse != null)
        {
          objectsResponse.Wait();
          return objectsResponse.Result.GetObjectsResponse1;
        }
        return null;
      }
      catch (WebException webEx)
      {
        ApplicationLogging.LogException(webEx, MethodBase.GetCurrentMethod());
        _provider.SetState(ProviderConnectionState.Unavailable);
        throw new RepositoryNotConnectedException(webEx.Message, webEx);
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private GetContentResponse GetContent(GetContentRequest request)
    {
      try
      {
        Task<GetContentResponse> response = _client.GetContentAsync(_localization, request);
        if (response != null)
        {
          response.Wait();
          return response.Result;
        }
        return null;
      }
      catch (WebException webEx)
      {
        ApplicationLogging.LogException(webEx, MethodBase.GetCurrentMethod());
        _provider.SetState(ProviderConnectionState.Unavailable);
        throw new RepositoryNotConnectedException(webEx.Message, webEx);
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    /// <summary>
    /// Ask FileNet for permissions
    /// </summary>
    /// <param name="lpFolderId"></param>
    /// <returns></returns>
    /// <remarks></remarks>
    private ListOfObject GetObjectStoreACEs()
    {
      try
      {
        ObjectSpecification objSpec = new ObjectSpecification();
        // objSpec.path = lpDocumentPath
        // objSpec.objectId = ObjectStoreName

        // Retrieve the Document
        ObjectRequestType objRequest = new ObjectRequestType();
        objSpec.classId = "ObjectStore";
        objSpec.objectStore = _objectStoreScope.objectStore;
        objRequest.SourceSpecification = objSpec;
        objRequest.id = "1";

        // Set up which properties are returned for the document
        FilterElementType[] incProps;
        objRequest.PropertyFilter = new PropertyFilterType();

        int lpNumberOfPropertyElements = 8;

        incProps = new FilterElementType[lpNumberOfPropertyElements + 1];

        // incProps(0) = New FilterElementType
        // incProps(0).Value = "FoldersFiledIn"

        // PathName
        incProps[0] = new FilterElementType();
        {
          var withBlock = incProps[0];
          withBlock.Value = "PathName";
          withBlock.maxRecursion = 1;
          withBlock.maxRecursionSpecified = true;
        }


        incProps[1] = new FilterElementType();
        incProps[1].Value = "Permissions";

        incProps[2] = new FilterElementType();
        incProps[2].Value = "GranteeName";

        incProps[3] = new FilterElementType();
        incProps[3].Value = "AccessType";

        incProps[4] = new FilterElementType();
        incProps[4].Value = "GranteeType";

        incProps[5] = new FilterElementType();
        incProps[5].Value = "InheritableDepth";

        incProps[6] = new FilterElementType();
        incProps[6].Value = "PermissionSource";

        incProps[7] = new FilterElementType();
        incProps[7].Value = "AccessMask";


        objRequest.PropertyFilter.IncludeProperties = incProps;
        objRequest.PropertyFilter.maxRecursion = 1;
        objRequest.PropertyFilter.maxRecursionSpecified = true;

        // Create the request array
        ObjectRequestType[] objRequestArray = new ObjectRequestType[2];
        objRequestArray[0] = objRequest;

        // Send off the request
        ObjectResponseType[] objResponseArray;
        objResponseArray = Array.Empty<ObjectResponseType>(); //new ObjectResponseType[] { };
        try
        {
          objResponseArray = GetObjects(objRequestArray);
        }
        catch (Exception Ex)
        {
          throw new Exception("An error occurred while retrieving object store'" + _objectStoreName + "'", Ex);
        }

        // Did we get a document back?
        if (objResponseArray.Length < 1)
          throw new Exception("No object store found for  '" + _objectStoreName + "'");

        if (objResponseArray[0].GetType().Name == "ErrorStackResponse")
        {
          ErrorStackResponse objErrResp = (ErrorStackResponse)objResponseArray[0];
          ErrorStackType objStack = objErrResp.ErrorStack;
          ErrorRecordType objErr = objStack.ErrorRecord[0];
          throw new Exception("Error [" + objErr.Description + "] occurred." + "  Err source is [" + objErr.Source + "]");
        }

        // Return the Document
        ObjectValue lobjObjectValue;
        SingleObjectResponse responseArray = (SingleObjectResponse)objResponseArray[0];
        lobjObjectValue = responseArray.Object;

        // Get each access permissions object
        CPEServiceReference.PropertyType lobjPropertyType;


        for (Int16 lintPropertyCounter = 0; lintPropertyCounter <= lobjObjectValue.Property.Length - 1; lintPropertyCounter++)
        {
          lobjPropertyType = lobjObjectValue.Property[lintPropertyCounter];
          switch (lobjPropertyType.propertyId)
          {
            case "Permissions":
              {
                return (ListOfObject)lobjPropertyType;
              }
          }
        }
      }
      catch (Exception)
      {
        throw;
      }

      return null/* TODO Change to default(_) if this is not a reference type */;
    }

    private ObjectSetType ExecuteSearch(string sql, int maxElements = 1, bool maxElementsSpecified = true)
    {
      try
      {
        RepositorySearch repositorySearch = (RepositorySearch)Factory.RepositorySearch(_objectStoreName, sql, RepositorySearchModeType.Rows);
        if (maxElementsSpecified)
        {
          repositorySearch.maxElements = maxElements;
          repositorySearch.maxElementsSpecified = maxElementsSpecified;
        }

        ObjectStoreScope objectStoreScope = new ObjectStoreScope() { objectStore = _objectStoreName };
        repositorySearch.SearchScope = objectStoreScope;

        Task<ExecuteSearchResponse> results = _client.ExecuteSearchAsync(_localization, repositorySearch);
        results.Wait();

        if (results.Result != null)
        {
          ExecuteSearchResponse searchResponse = results.Result;
          ObjectSetType objectSet = searchResponse.ExecuteSearchResponse1;
          return objectSet;
        }
        return null;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private static SearchResultSet BuildSearchResultSet(ObjectSetType objectSet)
    {
      //  Make sure we got a valid object set to work with
      if (objectSet == null) { throw new ArgumentNullException(nameof(objectSet)); }

      try
      {
        SearchResultSet resultSet = new SearchResultSet();

        //  Make sure we have some actual data to return
        if (objectSet.Object == null)
        {
          //  Simply return the empty result set
          return resultSet;
        }

        ModifiablePropertyType resultProperty;
        SearchResult searchResult;

        //  Iterate through each object returned from the search
        for (int resultCounter = 0; resultCounter < objectSet.Object.Length; resultCounter++)
        {
          searchResult = new SearchResult();

          //  Iterate through each property returned for each object
          for (int propertyCounter = 0; propertyCounter < objectSet.Object[resultCounter].Property.Length; propertyCounter++)
          {
            resultProperty = (ModifiablePropertyType)objectSet.Object[resultCounter].Property[propertyCounter];

            if (resultProperty.GetType().Name == "SingletonId")
            {
              SingletonId idProperty = (SingletonId)resultProperty;
              if (idProperty.propertyId == "Id")
              {
                searchResult.ID = idProperty.Value;
              }
            }
            //  Add the property to the search result
            searchResult.Values.Add(GetDataItem(resultProperty));
          }
          resultSet.Results.Add(searchResult);
        }

        return resultSet;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private static DataItem GetDataItem(ModifiablePropertyType resultProperty)
    {
      DataItem dataItem = null;
      try
      {
        switch (resultProperty.GetType().Name)
        {
          case "SingletonId":
            {
              if (resultProperty.propertyId == "Id")
              {
                dataItem = new DataItem(resultProperty.propertyId, Core.PropertyType.ecmGuid, ((SingletonId)resultProperty).Value, true);
              }
              else
              {
                dataItem = new DataItem(resultProperty.propertyId, Core.PropertyType.ecmGuid, ((SingletonId)resultProperty).Value, false);
              }
              break;
            }

          case "SingletonString":
            {
              dataItem = new DataItem(resultProperty.propertyId, Core.PropertyType.ecmString, ((SingletonString)resultProperty).Value, false);
              break;
            }

          case "SingletonBoolean":
            {
              dataItem = new DataItem(resultProperty.propertyId, Core.PropertyType.ecmBoolean, ((SingletonBoolean)resultProperty).Value, false);
              break;
            }

          case "SingletonDateTime":
            {
              dataItem = new DataItem(resultProperty.propertyId, Core.PropertyType.ecmDate, ((SingletonDateTime)resultProperty).Value, false);
              break;
            }

          case "SingletonFloat64":
            {
              dataItem = new DataItem(resultProperty.propertyId, Core.PropertyType.ecmDouble, ((SingletonFloat64)resultProperty).Value, false);
              break;
            }

          case "SingletonInteger32":
            {
              dataItem = new DataItem(resultProperty.propertyId, Core.PropertyType.ecmLong, ((SingletonInteger32)resultProperty).Value, false);
              break;
            }

          case "SingletonObject":
            {
              dataItem = new DataItem(resultProperty.propertyId, Core.PropertyType.ecmObject, ((SingletonObject)resultProperty).Value, false);
              break;
            }

          case "SingletonBinary":
            {
              dataItem = new DataItem(resultProperty.propertyId, Core.PropertyType.ecmBinary, ((SingletonBinary)resultProperty).Value, false);
              break;
            }

        }

        return dataItem;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private ObjectReference CreateObjectStoreReference()
    {
      try
      {
        return new ObjectReference() { classId = "ObjectStore", objectStore = _objectStoreName };
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    private ClassificationProperty GetClassificationProperty(ObjectValue propertyDefinition)
    {
      try
      {
        string id = string.Empty;
        Cardinality cardinality = Cardinality.ecmSingleValued;
        Core.PropertyType propertyType = Core.PropertyType.ecmUndefined;
        bool isSystemOwned = false;
        bool isValueRequired = false;
        string name = string.Empty;
        string symbolicName = string.Empty;
        bool isHidden = true;
        Core.ClassificationProperty.SettabilityEnum settability = ClassificationProperty.SettabilityEnum.READ_WRITE;
        string choiceListId = string.Empty;
        ChoiceList choiceList = null/* TODO Change to default(_) if this is not a reference type */;
        Nullable<int> maximumLengthString = null;
        Nullable<DateTime> minimumDateTime = null;
        Nullable<DateTime> maximumDateTime = null;
        Nullable<int> minimumInteger32 = null;
        Nullable<int> maximumInteger32 = null;
        Nullable<double> minimumFloat64 = null;
        Nullable<double> maximumFloat64 = null;
        string errorMessage = string.Empty;

        Nullable<bool> propertyDefaultBoolean = null;
        Nullable<DateTime> propertyDefaultDateTime = null;
        Nullable<double> ldblPropertyDefaultFloat64 = null;
        Nullable<Guid> propertyDefaultId = null;
        Nullable<int> propertyDefaultInteger32 = null;
        string propertyDefaultString = string.Empty;

        ModifiablePropertyType cewsProperty = null/* TODO Change to default(_) if this is not a reference type */;

        ClassificationProperty classificationProperty;

        for (Int16 propDefPropertyCounter = 0; propDefPropertyCounter <= propertyDefinition.Property.Length - 1; propDefPropertyCounter++)
        {
          cewsProperty = (ModifiablePropertyType)propertyDefinition.Property[propDefPropertyCounter];
          switch (cewsProperty.propertyId)
          {
            case "This":
              {
                SingletonObject propertyObject = (SingletonObject)cewsProperty;
                ObjectReference propertyObjectReference = (ObjectReference)propertyObject.Value;
                string thisID = propertyObjectReference.objectId;
                
                ObjectReference objectReference = Factory.ObjectReference("PropertyTemplateString", thisID, _objectStoreName);

                PropertyFilterType propertyFilter = Factory.PropertyFilterType(0);

                FilterElementType[] includeProperties = Array.Empty<FilterElementType>(); // new FilterElementType[] { };

                FilterElementType objectNameFilter = new FilterElementType();
                {
                  var withBlock = objectNameFilter;
                  withBlock.Value = "Name";
                }

                FilterElementType objectDisplayNameFilter = new FilterElementType();
                {
                  var withBlock = objectDisplayNameFilter;
                  withBlock.Value = "DisplayName";
                }

                FilterElementType objectSettabilityFilter = new FilterElementType();
                {
                  var withBlock = objectSettabilityFilter;
                  // .maxRecursion = 5
                  // .maxRecursionSpecified = True
                  withBlock.Value = "Settability";
                }

                FilterElementType objectChoiceListFilter = new FilterElementType();
                {
                  var withBlock = objectChoiceListFilter;
                  withBlock.maxRecursion = 5;
                  withBlock.maxRecursionSpecified = true;
                  withBlock.Value = "ChoiceList";
                }

                {
                  var withBlock = propertyFilter;
                  withBlock.IncludeProperties = new FilterElementType[4];
                  withBlock.IncludeProperties[0] = objectNameFilter;
                  withBlock.IncludeProperties[1] = objectDisplayNameFilter;
                  withBlock.IncludeProperties[2] = objectSettabilityFilter;
                  withBlock.IncludeProperties[3] = objectChoiceListFilter;
                }

                ObjectRequestType[] objectRequest = new ObjectRequestType[2];
                objectRequest[0] = new ObjectRequestType();
                {
                  var withBlock = objectRequest[0];
                  withBlock.SourceSpecification = objectReference;
                  withBlock.PropertyFilter = propertyFilter;
                }

                // WSEService.GetObjects(lobjObjectRequest)
                //ObjectResponseType[] lobjObjectResponse = WSEService.GetObjects(lobjObjectRequest);
                Task<GetObjectsResponse> objectsResponse = _client.GetObjectsAsync(_localization, objectRequest);

                if (objectsResponse != null)
                {
                  //lstrName = GetPropertyValueByName("Name", (SingleObjectResponse)objectsResponse[0].Object);
                  // lstrName = CType(CType(lobjObjectResponse(0), SingleObjectResponse).Object.Property(0), SingletonString).Value

                  // Try to get the ChoiceList if it is available
                  //SingleObjectResponse firstObjectResponse = (SingleObjectResponse)objectsResponse.GetObjectsResponse1[0];
                  //lobjChoiceList = GetChoiceList(objectsResponse[0], lstrErrorMessage);
                }

                break;
              }

            case "Cardinality":
              {
                SingletonInteger32 singletonInteger32 = (SingletonInteger32)cewsProperty;
                cardinality = GetECMCardinality(singletonInteger32.Value);
                break;
              }

            case "DataType":
              {
                SingletonInteger32 singletonInteger32 = (SingletonInteger32)cewsProperty;
                propertyType = GetECMPropertyType(singletonInteger32.Value);
                break;
              }

            case "Id":
              {
                SingletonId singletonId = (SingletonId)cewsProperty;
                id = singletonId.Value;
                break;
              }

            case "IsHidden":
              {
                SingletonBoolean singletonBoolean = (SingletonBoolean)cewsProperty;
                isHidden = singletonBoolean.Value;
                break;
              }

            case "IsSystemOwned":
              {
                SingletonBoolean singletonBoolean = (SingletonBoolean)cewsProperty;
                isSystemOwned = singletonBoolean.Value;
                break;
              }

            case "IsValueRequired":
              {
                SingletonBoolean singletonBoolean = (SingletonBoolean)cewsProperty;
                isValueRequired = singletonBoolean.Value;
                break;
              }

            case "Settability":
              {
                SingletonInteger32 singletonInteger32 = (SingletonInteger32)cewsProperty;
                settability = (ClassificationProperty.SettabilityEnum)singletonInteger32.Value;
                break;
              }

            case "ChoiceList":
              {
                SingletonObject singletonObject = (SingletonObject)cewsProperty;
                if (singletonObject.Value != null)
                {
                  ObjectValue objectValue = (ObjectValue)singletonObject.Value;
                  choiceListId = objectValue.objectId;
                  Core.ObjectIdentifier lobjIdentifier = new ObjectIdentifier(choiceListId, Core.ObjectIdentifier.IdTypeEnum.ID);
                  choiceList = GetChoiceList(lobjIdentifier);
                }

                break;
              }

            case "MaximumLengthString":
              {
                SingletonInteger32 singletonInteger32 = (SingletonInteger32)cewsProperty;
                if (singletonInteger32.ValueSpecified == true)
                  maximumLengthString = singletonInteger32.Value;
                else
                  maximumLengthString = null;
                break;
              }

            case "PropertyMinimumDateTime":
              {
                SingletonDateTime singletonDateTime = (SingletonDateTime)cewsProperty;
                if (singletonDateTime.ValueSpecified == true)
                  minimumDateTime = singletonDateTime.Value;
                else
                  minimumDateTime = null;
                break;
              }

            case "PropertyMaximumDateTime":
              {
                SingletonDateTime singletonDateTime = (SingletonDateTime)cewsProperty;
                if (singletonDateTime.ValueSpecified == true)
                  maximumDateTime = singletonDateTime.Value;
                else
                  maximumDateTime = null;
                break;
              }

            case "PropertyMinimumInteger32":
              {
                SingletonInteger32 singletonInteger32 = (SingletonInteger32)cewsProperty;
                if (singletonInteger32.ValueSpecified == true)
                  minimumInteger32 = singletonInteger32.Value;
                else
                  minimumInteger32 = null;
                break;
              }

            case "PropertyMaximumInteger32":
              {
                SingletonInteger32 singletonInteger32 = (SingletonInteger32)cewsProperty;
                if (singletonInteger32.ValueSpecified == true)
                  maximumInteger32 = singletonInteger32.Value;
                else
                  maximumInteger32 = null;
                break;
              }

            case "PropertyMinimumFloat64":
              {
                SingletonFloat64 singletonFloat64 = (SingletonFloat64)cewsProperty;
                if (singletonFloat64.ValueSpecified == true)
                  minimumFloat64 = singletonFloat64.Value;
                else
                  minimumFloat64 = null;
                break;
              }

            case "PropertyMaximumFloat64":
              {
                SingletonFloat64 singletonFloat64 = (SingletonFloat64)cewsProperty;
                if (singletonFloat64.ValueSpecified == true)
                  maximumFloat64 = singletonFloat64.Value;
                else
                  maximumFloat64 = null;
                break;
              }

            case "PropertyDefaultBoolean":
              {
                SingletonBoolean singletonBoolean = (SingletonBoolean)cewsProperty;
                if (singletonBoolean.ValueSpecified == true)
                  propertyDefaultBoolean = singletonBoolean.Value;
                else
                  propertyDefaultBoolean = null;
                break;
              }

            case "PropertyDefaultDateTime":
              {
                SingletonDateTime singletonDateTime = (SingletonDateTime)cewsProperty;
                if (singletonDateTime.ValueSpecified == true)
                  propertyDefaultDateTime = singletonDateTime.Value;
                else
                  propertyDefaultDateTime = null;
                break;
              }

            case "PropertyDefaultFloat64":
              {
                SingletonFloat64 singletonFloat64 = (SingletonFloat64)cewsProperty;
                if (singletonFloat64.ValueSpecified == true)
                  ldblPropertyDefaultFloat64 = singletonFloat64.Value;
                else
                  ldblPropertyDefaultFloat64 = null;
                break;
              }

            case "PropertyDefaultId":
              {
                SingletonId singletonId = (SingletonId)cewsProperty;
                if (singletonId.Value != null)
                  propertyDefaultId = new Guid(singletonId.Value);
                else
                  propertyDefaultId = null;
                break;
              }

            case "PropertyDefaultInteger32":
              {
                SingletonInteger32 singletonInteger32 = (SingletonInteger32)cewsProperty;
                if (singletonInteger32.ValueSpecified == true)
                  propertyDefaultInteger32 = singletonInteger32.Value;
                else
                  propertyDefaultInteger32 = null;
                break;
              }

            case "PropertyDefaultString":
              {
                SingletonString singletonString = (SingletonString)cewsProperty;
                propertyDefaultString = singletonString.Value;
                break;
              }

            case "Name":
              {
                SingletonString singletonString = (SingletonString)cewsProperty;
                name = singletonString.Value;
                break;
              }

            case "SymbolicName":
              {
                SingletonString singletonString = (SingletonString)cewsProperty;
                symbolicName = singletonString.Value;
                break;
              }
          }
        }

        classificationProperty = (ClassificationProperty)ClassificationPropertyFactory.Create(propertyType, name, symbolicName, cardinality);

        {
          var withBlock = classificationProperty;
          withBlock.IsSystemProperty = isSystemOwned;
          withBlock.IsHidden = isHidden;
          withBlock.IsRequired = isValueRequired;
          withBlock.SetID(id);
          withBlock.SetPackedName(symbolicName);
          withBlock.Settability = settability;
          if (choiceList != null)
            withBlock.ChoiceList = choiceList;
        }

        switch (propertyType)
        {
          case Core.PropertyType.ecmString:
            {
              {
                var withBlock = (ClassificationStringProperty)classificationProperty;
                withBlock.DefaultValue = propertyDefaultString;
                withBlock.MaxLength = maximumLengthString;
              }

              break;
            }

          case Core.PropertyType.ecmLong:
            {
              {
                var withBlock = (ClassificationLongProperty)classificationProperty;
                withBlock.DefaultValue = propertyDefaultInteger32;
                withBlock.MinValue = minimumInteger32;
                withBlock.MaxValue = maximumInteger32;
              }

              break;
            }

          case Core.PropertyType.ecmBoolean:
            {
              {
                var withBlock = (ClassificationBooleanProperty)classificationProperty;
                withBlock.DefaultValue = propertyDefaultBoolean;
              }

              break;
            }

          case Core.PropertyType.ecmDate:
            {
              {
                var withBlock = (ClassificationDateTimeProperty)classificationProperty;
                withBlock.DefaultValue = propertyDefaultDateTime;
                withBlock.MinValue = minimumDateTime;
                withBlock.MaxValue = maximumDateTime;
              }

              break;
            }

          case Core.PropertyType.ecmDouble:
            {
              {
                var withBlock = (ClassificationDoubleProperty)classificationProperty;
                withBlock.DefaultValue = ldblPropertyDefaultFloat64;
                withBlock.MinValue = minimumFloat64;
                withBlock.MaxValue = maximumFloat64;
              }

              break;
            }

          case Core.PropertyType.ecmGuid:
            {
              {
                var withBlock = (ClassificationGuidProperty)classificationProperty;
                withBlock.DefaultValue = propertyDefaultId;
              }

              break;
            }
        }

        return classificationProperty;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, System.Reflection.MethodBase.GetCurrentMethod());
        // Re-throw the exception to the caller
        throw;
      }
    }

    private string GetChoiceListObjectId(ObjectIdentifier id)
    {
      try
      {
        string sql;
        string objectId = string.Empty;
        SingletonId singletonId = null;

        //  We will execute a query to find the object by name.
        //  If we find the object, we will return it's object id.

        if (id.IdentifierType == ObjectIdentifier.IdTypeEnum.ID)
        {
          sql = $"SELECT [Id] FROM [ChoiceList] WHERE ([Id] = {id.IdentifierValue})";
        }
        else
        {
          sql = $"SELECT [Id] FROM [ChoiceList] WHERE ([DisplayName] = {id.IdentifierValue})";
        }

        Debug.WriteLine("Performing an ExecuteSearch, SQL:");
        Debug.WriteLine($" {sql}");
        
        ObjectSetType objectSet = ExecuteSearch(sql);

        if (objectSet != null)
        {
          ObjectValue value = (ObjectValue)objectSet.Object.GetValue(0);
          singletonId = (SingletonId)value.Property.GetValue(0);
          objectId = singletonId.Value;
        }

        return objectId;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private ObjectValue GetCEChoiceList(ObjectIdentifier id)
    {
      try
      {

        if (id == null) { throw new ArgumentNullException(nameof(id)); }

        string objectId;

        objectId = id.IdentifierValue;

        if (string.IsNullOrEmpty(objectId)) { return null; }

        //  Create an object reference
        ObjectReference objectReference = Factory.ObjectReference("ChoiceList", objectId, _objectStoreName);

        //  Add an include filter
        PropertyFilterType propertyFilter = Factory.PropertyFilterType(0);

        //  Add an exclude filter
        string[] excludePropertyFilter = { "ClassDescription", "ObjectStore", "Creator", "LastModifier", "DateCreated", "LastModified", "DateLastModified", "AuditedEvents", "Owner", "Permissions" };

        FilterElementType propertyDefinitionFilter = new FilterElementType() { maxRecursion = 10, maxRecursionSpecified = true, Value = "ChoiceList" };

        FilterElementType filterElementChoiceValues = new FilterElementType() { Value = "ChoiceValues" };

        FilterElementType filterElementChoice = new FilterElementType() { Value = "Choice" };

        FilterElementType includeTypesFilter_0 = new FilterElementType() { Value = "EnumOfObject", maxRecursion = 0, maxRecursionSpecified = true };

        FilterElementType includeTypesFilter_1 = new FilterElementType() { Value = "Singleton* List*", maxRecursion = 5, maxRecursionSpecified = true };

        propertyFilter.IncludeTypes = new FilterElementType[] { includeTypesFilter_0, includeTypesFilter_1 };
        propertyFilter.ExcludeProperties = excludePropertyFilter;

        //  Create the object request
        ObjectRequestType[] objectRequest = new ObjectRequestType[1] { new ObjectRequestType() { SourceSpecification = objectReference, PropertyFilter = propertyFilter } };

        //  Send off the request
        Task<GetObjectsResponse> getObjects = _client.GetObjectsAsync(_localization, objectRequest);

        getObjects.Wait();

        var objectResponse = getObjects.Result.GetObjectsResponse1;
        SingleObjectResponse singleObject = (SingleObjectResponse)objectResponse[0];
        return singleObject.Object;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private ChoiceList GetChoiceList(SingleObjectResponse property)
    {
      try
      {
        ObjectValue choiceListObjectValue;
        string choiceListId = string.Empty;
        ObjectIdentifier objectIdentifier;

        foreach (ModifiablePropertyType propertyType in property.Object.Property.Cast<ModifiablePropertyType>())
        {
          if (propertyType.propertyId == "ChoiceList")
          {
            SingletonObject singletonObject = (SingletonObject)propertyType;
            choiceListObjectValue = (ObjectValue)singletonObject.Value;
            if (choiceListObjectValue == null) { return null; }
            choiceListId = choiceListObjectValue.objectId;
            break;
          }
        }

        if (choiceListId.Length > 1)
        {
          objectIdentifier = new ObjectIdentifier(choiceListId, ObjectIdentifier.IdTypeEnum.ID);
          return GetChoiceList(objectIdentifier);
        }

        return null;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private ChoiceList GetChoiceList(ObjectIdentifier id)
    {
      try
      {

        ObjectValue choiceListObject = (ObjectValue)GetCEChoiceList(id);

        if (choiceListObject == null) { return null; }

        Debug.WriteLine(choiceListObject.objectId);

        ChoiceList choiceList = new ChoiceList();
        ListOfObject choiceValues = null;

        SingletonString singletonString;

        foreach (CPEServiceReference.PropertyType property in choiceListObject.Property)
        {
          switch (property.propertyId)
          {
            case "Name":
              singletonString = (SingletonString)property;
              choiceList.Name = singletonString.Value;
              break;

            case "DisplayName":
              singletonString = (SingletonString)property;
              choiceList.DisplayName = singletonString.Value;
              break;

            case "DescriptiveText":
              singletonString = (SingletonString)property;
              choiceList.DescriptiveText = singletonString.Value;
              break;

            case "ChoiceValues":
              choiceValues = (ListOfObject)property;
              break;
          }
        }

        choiceList.Id = choiceListObject.objectId;
        choiceList.ChoiceValues = ChoiceValues(choiceValues);

        return choiceList;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private ChoiceValues ChoiceValues(ListOfObject cewsChoiceValues, ChoiceValues parent = null)
    {
      try
      {
        DependentObjectType choiceValue;
        ChoiceValues ctsChoiceValues;

        if (parent == null)
        {
          ctsChoiceValues = new ChoiceValues();
        }
        else
        {
          ctsChoiceValues = parent;
        }

        for (int valueCounter = 0; valueCounter < cewsChoiceValues.Value.Length; valueCounter++)
        {
          choiceValue = cewsChoiceValues.Value[valueCounter];
          ctsChoiceValues.Add(ChoiceItem(choiceValue));
        }

        return ctsChoiceValues;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private ChoiceItem ChoiceItem(DependentObjectType choice)
    {
      try
      {
        CPEServiceReference.PropertyType choiceProperty;
        string name = string.Empty;
        string displyName = string.Empty;
        string id = string.Empty;
        int choiceType = -1;
        int choiceIntegerValue = -1;
        string choiceStringValue = null;
        //ChoiceValues ctsChoiceValues;
        ChoiceItem choiceItem = null;
        ListOfObject ceChoiceValues = null;

        SingletonString singletonString;
        SingletonId singletonId;
        SingletonInteger32 singletonInteger32;
        ListOfObject listOfObject;

        for (int objectCounter = 0; objectCounter < choice.Property.Length; objectCounter++)
        {
          choiceProperty = choice.Property[objectCounter];
          switch (choiceProperty.propertyId)
          {
            case "Name":
              singletonString = (SingletonString)choiceProperty;
              name = singletonString.Value;
              break;

            case "DisplayName":
              singletonString = (SingletonString)choiceProperty;
              displyName = singletonString.Value;
              break;

            case "Id":
              singletonId = (SingletonId)choiceProperty;
              id = singletonId.Value;
              break;

            case "ChoiceType":
              singletonInteger32 = (SingletonInteger32)choiceProperty;
              choiceType = singletonInteger32.Value;
              break;

            case "ChoiceIntegerValue":
              singletonInteger32 = (SingletonInteger32)choiceProperty;
              choiceIntegerValue = singletonInteger32.Value;
              break;

            case "ChoiceStringValue":
              singletonString = (SingletonString)choiceProperty;
              choiceStringValue = singletonString.Value;
              break;

            case "ChoiceValues":
              listOfObject = (ListOfObject)choiceProperty;
              ceChoiceValues = listOfObject;
              break;
          }
        }

        switch (choiceType)
        {
          case 0:
            //  This is a ChoiceValue of type Integer
            choiceItem = new ChoiceValue(choiceIntegerValue) { Name = name, DisplayName = displyName, Id = id };
            break;

          case 1:
            //  This is a ChoiceValue of type String
            choiceItem = new ChoiceValue(choiceStringValue) { Name = name, DisplayName = displyName, Id = id }; 
            break;

          case 2:
          case 3:
            //  This is a ChoiceGroup
            choiceItem = new ChoiceGroup(name) { Id = id, ChoiceValues = ChoiceValues(ceChoiceValues) };
            break;
        }

        return choiceItem;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    /// <summary>
    /// Blocks nonword characters, accented letters and other alphabetical symbols
    /// </summary>
    /// <param name="originalName">The containment name to clean.</param>
    /// <returns>The original containment name scrubbed of known illegal characters.</returns>
    /// <remarks>
    /// Emulates the behavior of FileNet Enterprise Manager for Containment Name when creating objects.
    /// </remarks>
    /// <exception cref="ArgumentNullException">Thrown if the original name is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown if the original name is an empty string.</exception>
    private static string CleanContainmentName(string originalName)
    {
      try
      {
        //  The name cannot be an empty string and 
        //  cannot contain any of the following characters: 
        //  \ / : * ? " < > |

        if (originalName == null) { throw new ArgumentNullException(nameof(originalName), "Cannot clean null containment name."); }
        if (originalName.Length == 0) { throw new ArgumentOutOfRangeException(nameof(originalName), "Cannot clean empty containment name."); }

        Regex regex = new Regex("[\\W-[\\s-]]*", RegexOptions.Compiled);
        return regex.Replace(originalName, string.Empty);
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    private static Cardinality GetECMCardinality(int cardinality)
    {
      try
      {
        switch (cardinality)
        {
          case 0:
          case 1:
            return Cardinality.ecmSingleValued;
          case 2:
            return Cardinality.ecmMultiValued;
        default:
            throw new ArgumentException("Invalid cardinality value", nameof(cardinality));
        }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private static Core.PropertyType GetECMPropertyType(int dataType)
    {
      try
      {
        switch (dataType)
        {
          case 1:
            return Core.PropertyType.ecmBinary;
          case 2:
            return Core.PropertyType.ecmBoolean;
          case 3:
            return Core.PropertyType.ecmDate;
          case 4:
            return Core.PropertyType.ecmDouble;
          case 5:
            return Core.PropertyType.ecmGuid;
          case 6:
            return Core.PropertyType.ecmLong;
          case 7:
            return Core.PropertyType.ecmObject;
          case 8:
            return Core.PropertyType.ecmString;
          default:
            throw new ArgumentOutOfRangeException(nameof(dataType), "Invalid dataType value");
        }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private ModifiablePropertyType[] BuildAnnotationPropertyList(string annotationId, string parentObjectId)
    {
      try
      {
        ObjectValue annotationReference = GetObject(parentObjectId, "Document");
        ModifiablePropertyType[] inputProps = new ModifiablePropertyType[2];
        inputProps[0] = Factory.SingletonObject("AnnotatedObject", annotationReference, true);
        inputProps[1] = Factory.SingletonId("ID", annotationId, true);
        return inputProps;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private ModifiablePropertyType[] BuildPropertyList(Core.Version version)
    {
      //  Build a list of properties to set in the new document 
      try
      {
        return BuildPropertyList(version.Properties);
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private ModifiablePropertyType[] BuildPropertyList(IProperties properties)
    {
      try
      {
        int propertyCounter = 0;
        ModifiablePropertyType[] inputProps;
        object ceProperty;
        bool setSystemProperties;
        List<string> exclusions = _provider.GetAllContentExportPropertyExclusions();
        inputProps = new ModifiablePropertyType[properties.Count];

        //  Add additional exclusions
        exclusions.Add("ComponentBindingLabel");
        exclusions.Add("CompoundDocumentState");
        exclusions.Add("Creator");
        exclusions.Add("EntryTemplateId");
        exclusions.Add("EntryTemplateObjectStoreName");
        exclusions.Add("LastModifier");
        exclusions.Add("Owner");
        exclusions.Add("VersionId");
        exclusions.Add("VersionSeriesId");

        try
        {
          setSystemProperties = (bool)_provider.ActionProperties[CProvider.ACTION_SET_SYSTEM_PROPERTIES].Value;
        }
        catch (Exception ex)
        {
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
          // Just treat it as false
          setSystemProperties = false;
        }

        foreach (IProperty property in properties)
        {
          if (property.Name.IsLike(true,
                            PROP_CREATOR,
                            PROP_CHECK_IN_DATE,
                            PROP_CREATE_DATE,
                            PROP_MODIFY_USER,
                            PROP_MODIFY_DATE,
                            PROP_MIME_TYPE))
          {
            if (setSystemProperties)
            {
              if (HasPrivelegedAccess)
              {
                ceProperty = BuildCEProperty(property);
                continue;
              }
            }
          }

          if (property.Name.IsLike(true, _provider.ActionProperties.NameArray()))
          {
            //  Skip over these, they are action properties
            continue;
          }

          if (exclusions.Contains(property.Name)) { continue; }

          ceProperty = BuildCEProperty(property);

          if (ceProperty != null)
          {
            inputProps[propertyCounter] = (ModifiablePropertyType)ceProperty;
            propertyCounter++;
          }

        }

        return inputProps;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private static ModifiablePropertyType BuildCEProperty(IProperty property)
    {

      //SingletonString singletonStringProperty;
      //SingletonBinary singletonBinaryProperty;
      //SingletonBoolean singletonBooleanProperty;
      //SingletonDateTime singletonDateProperty;
      //SingletonFloat64 singletonFloatProperty;
      //SingletonId singletonIdProperty;
      //SingletonInteger32 singletonIntegerProperty;
      //SingletonObject singletonObjectProperty;
      ListOfString listOfStringProperty;
      ListOfInteger32 listOfIntegerProperty;
      string[] strings = { string.Empty };
      int[] integers = { 1 };
      //bool valueSpecified = false;
      string cleanPropertyName;

      ModifiablePropertyType ceProperty = null;

      try
      {
        cleanPropertyName = CleanSymbolicName(property.Name);

        switch (property.Cardinality)
        {
          case Cardinality.ecmSingleValued:
            {
              switch (property.Type)
              {
                case Core.PropertyType.ecmBinary:
                  {
                    ceProperty = Factory.SingletonBinary(cleanPropertyName, true, true, (byte[])property.Value);
                    break;
                  }

                case Core.PropertyType.ecmBoolean:
                  {
                    ceProperty = Factory.SingletonBoolean(cleanPropertyName, (bool)property.Value, true);
                    break;
                  }

                case Core.PropertyType.ecmDate:
                  {
                    if ((property.Value.GetType() == typeof(string)) && (string.IsNullOrEmpty((string)property.Value)))
                    {
                      return ceProperty;
                    }
                    else
                    {
                      ceProperty = Factory.SingletonDateTime(cleanPropertyName, (DateTime)property.Value, true);
                      break;
                    }
                  }

                case Core.PropertyType.ecmDouble:
                  {
                    ceProperty = Factory.SingletonFloat64(cleanPropertyName, (double)property.Value, true);
                    break;
                  }

                case Core.PropertyType.ecmGuid:
                  {
                    ceProperty = Factory.SingletonId(cleanPropertyName, (string)property.Value, true);
                    break;
                  }

                case Core.PropertyType.ecmLong:
                  {
                    long propValue;
                    if ((property.Value.GetType() == typeof(string)))
                    {
                      if (long.TryParse((string)property.Value, out propValue))
                      {
                        ceProperty = Factory.SingletonInteger32(cleanPropertyName, propValue, true);
                      }
                    }
                    break;
                  }

                //case Core.PropertyType.ecmObject:
                //  {
                //    if (property.HasValue) { ceProperty = Factory.SingletonObject(cleanPropertyName, property.Value, true); }                    
                //    break;
                //  }

                case Core.PropertyType.ecmString:
                  {
                    ceProperty = Factory.SingletonString(cleanPropertyName, (string)property.Value, true);
                    break;
                  }

              }
              break;              
            }

          case Cardinality.ecmMultiValued:
            {
              switch (property.Type)
              {

                // NOTE: We have only implemented multi-valued 
                // properties here for integer and string properties.

                // If there is a need we can implement for the other types, 
                // although for the life of me I can't figure out why
                // anyone would ever use a multi-valued boolean.

                // Ernie Bahr

                case Core.PropertyType.ecmLong:
                  {
                    listOfIntegerProperty = Factory.ListOfInteger32(cleanPropertyName, true, true);
                    Values values = (Values)property.Values;
                    Array.Resize(ref integers, values.Count());
                    for (Int16 lintValueCounter = 0; lintValueCounter < values.Count(); lintValueCounter++)
                      integers[lintValueCounter] = (int)values.GetItemByIndex(lintValueCounter);
                    ceProperty = listOfIntegerProperty;
                    break;
                  }

                case Core.PropertyType.ecmString:
                  {
                    listOfStringProperty = Factory.ListOfString(cleanPropertyName, true, true);
                    Values values = (Values)property.Values;
                    Array.Resize(ref strings, values.Count());
                    for (Int16 lintValueCounter = 0; lintValueCounter < values.Count(); lintValueCounter++)
                      strings[lintValueCounter] = (string)values.GetItemByIndex(lintValueCounter);
                    ceProperty = listOfStringProperty;
                    break;
                  }

              }

              break;
            }
        }

        return ceProperty;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    /// <summary>
    /// Blocks nonword characters, accented letters and other alphabetical symbols
    /// </summary>
    /// <param name="originalName"></param>
    /// <returns></returns>
    /// <remarks>Emulates the behavior of FileNet Enterprise Manager for Symbolic Name when creating objects.</remarks>
    private static string CleanSymbolicName(string originalName)
    {
      try
      {
        //  The name cannot be an empty string and cannot contain any of the following characters: \ / : * ? " < > |
        return new Regex("\\b\\W*", RegexOptions.Compiled).Replace(originalName, string.Empty);
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private static bool IsMajorVersion(Core.Version version, bool deleteProperty = true)
    {
      try
      {
        if ((version.Properties.PropertyExists(MAJOR_VERSION)) && ((bool)version.Properties[MAJOR_VERSION].Value == true))
        {
          if (deleteProperty)
          {
            version.Properties.Delete(MAJOR_VERSION);            
          }
          return true;
        }
        else
        {
          return false;
        }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public ObjectValue GetVersionSeries(string versionId)
    {
      try
      {
        ObjectSpecification spec = new ObjectSpecification();
        ObjectRequestType request = new ObjectRequestType();
        spec.objectId = versionId;
        spec.classId = "Document";
        spec.objectStore = _objectStoreName;
        request.SourceSpecification = spec;
        request.id = "1";

        FilterElementType[] incProps = new FilterElementType[4] { 
          new FilterElementType() {  Value = "VersionSeries" }, 
          new FilterElementType() { Value = "Versions" }, 
          new FilterElementType() { Value = "MajorVersionNumber" }, 
          new FilterElementType() { Value = "MinorVersionNumber" } };

        request.PropertyFilter = new PropertyFilterType() { IncludeProperties = incProps, maxRecursion = 2, maxRecursionSpecified = true };

        ObjectRequestType[] lpRequest = new ObjectRequestType[2] { request, null };
        ObjectResponseType[] array2 = Array.Empty<ObjectResponseType>(); //new ObjectResponseType[0];

        //NOTE: Setting the recursion level to 2 fails to get the version numbers for the  
        //      Version objects in the VersionSeries, it is too low. 

        //      Setting the number to 3 or higher results in an exception. 

        //      We will need to take a different approach.  
        //      Recommend two possible approaches

        //      1. Do a search for documents using the VersionSeriesID as a search criteria
        //         with the version numbers as some of the result columns

        //      2. Individually get the version objects from the version ids and then get 
        //         the version numbers one at a time 

        //      We need the version numbers in order to properly sequence the versions in the document object.  
        //      The other option is to add version sorting to the document object.  Then we could add the   
        //      versions in an arbitrary order and then sort the versions once we have completed building the stack.

        //  Create the request array
        ObjectRequestType[] objRequestArray = new ObjectRequestType[1] { request };

        //  Send off the request
        ObjectResponseType[] objResponseArray;

        try
        {
          objResponseArray = GetObjects(objRequestArray);
        }
        catch (Exception ex)
        {
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
          throw new Exception($"An error occurred while retrieving version series for document '{versionId}'", ex);
        }

        //  Did we get a document back?
        if (objResponseArray.Length < 1) { throw new Exception($"No document found for ID '{versionId}'"); }

        if (objResponseArray[0].GetType().Name == "ErrorStackResponse")
        {
          ErrorStackResponse errorStackResponse = (ErrorStackResponse)objResponseArray[0];
          ErrorStackType errorStack = errorStackResponse.ErrorStack;
          ErrorRecordType errorRecordType = errorStack.ErrorRecord[0];
          throw new Exception($"Error [{errorRecordType.Description}] occurred.  Err source is [{errorRecordType.Source}]");
        }

        return ((SingleObjectResponse)objResponseArray[0]).Object;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public string GetVersionSeriesID(string versionId)
    {
      try
      {
        //  Get the Version Series
        ObjectValue versionSeries = GetVersionSeries(versionId);
        foreach (CPEServiceReference.PropertyType propertyType in versionSeries.Property)
        {
          if (propertyType.propertyId == "VersionSeries")
          {
            return ((ObjectValue)((SingletonObject)propertyType).Value).objectId;
          }
        }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }

      return default;
    }

    private ObjectRequestType CreateDocumentRequest(string id, RequestType requestType, [Optional][DefaultParameterValue(null)] ref string[] excludeProperties, [Optional][DefaultParameterValue(true)] ref bool exportContent, [Optional][DefaultParameterValue(true)] ref bool exportAnnotations, string classId = "Document")
    {
      checked
      {
        try
        {
          ObjectSpecification spec = new ObjectSpecification();
          ObjectRequestType request = new ObjectRequestType();
          switch (requestType)
          {
            case RequestType.BasicProperties:
              spec.objectId = id;

              //  Retrieve the document
              spec.classId = classId;
              spec.objectStore = _objectStoreName;
              request.SourceSpecification = spec;
              request.id = "1";

              if (excludeProperties != null) { request.PropertyFilter = new PropertyFilterType() { ExcludeProperties = excludeProperties }; }

              return request;

            case RequestType.ExtendedProperties:
              {
                spec.objectId = id;

                //  Retrieve the document
                spec.classId = classId;
                spec.objectStore = _objectStoreName;
                request.SourceSpecification = spec;
                request.id = "1";

                FilterElementType[] incProps;
                request.PropertyFilter = new PropertyFilterType();
                int numberOfPropertyElements = 10;
                if (!exportContent) { numberOfPropertyElements = 5; }
                if (exportAnnotations) { numberOfPropertyElements++; }

                incProps = new FilterElementType[numberOfPropertyElements];

                incProps[0] = new FilterElementType() { Value = "FoldersFiledIn" };
                incProps[1] = new FilterElementType() { Value = "PathName", maxRecursion = 1, maxRecursionSpecified = true };
                incProps[2] = new FilterElementType() { Value = "VersionSeries" };
                incProps[3] = new FilterElementType() { Value = "Versions" };

                if (exportContent)
                {
                  //  Ask for the content properties...
                  ulong maxSize = Convert.ToUInt64(10000000);
                  incProps[4] = new FilterElementType() { Value = "ContentElements" };
                  incProps[5] = new FilterElementType() { Value = "ContentData", maxSize = maxSize, maxSizeSpecified = true };
                  incProps[6] = new FilterElementType() { Value = "Content" };
                  incProps[7]= new FilterElementType() { Value = "RetrievalName", maxRecursion = 5, maxRecursionSpecified = true };
                  incProps[8] = new FilterElementType() { Value = "ContentLocation", maxRecursion = 4, maxRecursionSpecified = true };
                  incProps[9] = new FilterElementType() { Value = "ContentType", maxRecursion = 4, maxRecursionSpecified = true };
                }

                if (exportAnnotations)
                {
                  incProps[10] = new FilterElementType() { Value = "Annotations", maxRecursion = 5, maxRecursionSpecified = true };
                }

                request.PropertyFilter.IncludeProperties =  incProps;
                request.PropertyFilter.maxRecursion = 1;
                request.PropertyFilter.maxRecursionSpecified = true;

                return request;
              }
            case RequestType.FoldersFiledIn:
              {
                spec.objectId = id;

                //  Retrieve the document
                spec.classId = "Document";
                spec.objectStore = _objectStoreName;
                request.SourceSpecification = spec;
                request.id = "1";

                FilterElementType[] incProps;
                request.PropertyFilter = new PropertyFilterType();
                incProps = new FilterElementType[2];
                incProps[0] = new FilterElementType() { Value = "FoldersFiledIn" };
                incProps[1] = new FilterElementType() { Value = "PathName", maxRecursion = 1, maxRecursionSpecified = true };

                request.PropertyFilter.IncludeProperties = incProps;
                request.PropertyFilter.maxRecursion = 1;
                request.PropertyFilter.maxRecursionSpecified = true;

                return request;

              }

            default:
              return request;
          }
        }
        catch (Exception ex)
        {
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
          //  Re - throw the exception to the caller
          throw;
        }
      }
    }

    //private void GetPropertyDescription(string classId, string objectId)
    //{
    //  try
    //  {
    //    ObjectReference sourceSpecification = Factory.ObjectReference("PropertyDefinition", objectId, _objectStoreName);
    //    PropertyFilterType propertyFilter = Factory.PropertyFilterType(0);
    //    FilterElementType[] incProps = new FilterElementType[9];
    //    {
    //      new FilterElementType() { Value = "PropertyDescription" }; 
    //    new FilterElementType() { Value = "PropertyDescriptionBinary" }; 
    //    new FilterElementType() { Value = "PropertyDescriptionBoolean" };
    //    new FilterElementType() { Value = "PropertyDescriptionDateTime" };
    //    new FilterElementType() { Value = "PropertyDescriptionFloat64" };
    //    new FilterElementType() { Value = "PropertyDescriptionId" };
    //    new FilterElementType() { Value = "PropertyDescriptionInteger32" };
    //    new FilterElementType() { Value = "PropertyDescriptionObject" };
    //    new FilterElementType() { Value = "PropertyDescriptionString" };
    //    };

    //    propertyFilter.IncludeProperties = incProps;

    //    ObjectRequestType[] request = new ObjectRequestType[1] { new ObjectRequestType () { SourceSpecification = sourceSpecification, PropertyFilter = propertyFilter } };

    //    ObjectResponseType[] responses = GetObjects(request);

    //  }
    //  catch (Exception ex)
    //  {
    //    ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
    //    //  Re - throw the exception to the caller
    //    throw;
    //  }
    //}

    //private GetSearchMetadataResponse GetMetadata(string lpClassFilter, string lpPropertyFilter = "", string lpSearchScope = "", bool lpFilterProperties = true)
    //{
    //  try
    //  {
    //    GetSearchMetadataRequest getSearchMetadataRequest = new GetSearchMetadataRequest();
    //    if (lpSearchScope.Length > 0)
    //    {
    //      ObjectStoreScope objectStoreScope = new ObjectStoreScope();
    //      objectStoreScope.objectStore = _objectStoreName;
    //      getSearchMetadataRequest.SearchScope = objectStoreScope;
    //    }

    //    getSearchMetadataRequest.ClassFilter = lpClassFilter;
    //    if (lpPropertyFilter.Length > 0)
    //    {
    //      PropertyFilterType propertyFilterType = new PropertyFilterType();
    //      propertyFilterType.maxRecursion = 1;
    //      propertyFilterType.maxRecursionSpecified = true;
    //      if (lpFilterProperties)
    //      {
    //        propertyFilterType.IncludeProperties = new FilterElementType[2];
    //        propertyFilterType.IncludeProperties[0] = new FilterElementType();
    //        propertyFilterType.IncludeProperties[0].Value = lpPropertyFilter;
    //      }

    //      getSearchMetadataRequest.PropertyFilter = propertyFilterType;
    //    }

    //    return WSEService.GetSearchMetadata(getSearchMetadataRequest);
    //  }
    //  catch (Exception ex)
    //  {
    //    ProjectData.SetProjectError(ex);
    //    Exception ex2 = ex;
    //    ApplicationLogging.LogException(ex2, MethodBase.GetCurrentMethod());
    //    throw new Exception("Search Failed [" + ex2.Message + "]", ex2);
    //  }
    //}


    private static CPEServiceReference.PropertyType GetPropertyValue(string propertyId, ChangeResponseType responseObject)
    {
      try
      {
        //  Capture value of the property in the returned doc object
        if (responseObject.Property == null) { return null; }

        foreach (CPEServiceReference.PropertyType responseProperty in responseObject.Property)
        {
          //  If property found, store its value
          if (responseProperty.propertyId == propertyId) { return responseProperty; }
        }
        return null;

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
