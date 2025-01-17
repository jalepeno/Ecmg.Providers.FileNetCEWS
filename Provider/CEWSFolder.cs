using CPEServiceReference;
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
  public class CEWSFolder : CFolder
  {

    #region Class Variables

    ObjectValue _ceFolder;
    Folders _subFolders;
    FolderContents _folderContents;
    int _subFolderCount = -1;

    #endregion

    #region Constructors

    public CEWSFolder() { }

    public CEWSFolder(ref ObjectValue ceFolder, CProvider provider)
    {
      try
      {
        string pathName = string.Empty;
        _ceFolder = ceFolder;

        pathName = (string)CEWSProvider.GetPropertyValueByName("PathName", ceFolder);

        InitializeFolderCollection(pathName);

        Provider = provider;
        InitializeFolder();
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }    
    }

    #endregion

    #region CFolder Overrides

    public override FolderContents Contents
    {
      get
      {
        try
        {
          if (_folderContents == null) { InitializeFolderContents(); }
          return _folderContents;
        }
        catch (Exception ex)
        {
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
          //  Re - throw the exception to the caller
          throw;
        }
      }
    }

    public virtual string FolderClass
    {
      get
      {
        try
        {
          if (Properties.PropertyExists("FolderClass"))
          {
            return Properties["FolderClass"].Value.ToString();
          }
          else { return "UNKNOWN"; }
        }
        catch (Exception ex)
        {
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
          //  Re - throw the exception to the caller
          throw;
        }
      }
    }

    public override Folders SubFolders
    {
      get
      {
        try
        {
            return GetSubFolders(true);
          }
        catch (Exception ex)
        {
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
          //  Re - throw the exception to the caller
          throw;
        }
      }
    }

    public override string GetID()
    {
      try
      {
        if (_ceFolder != null) { return _ceFolder.objectId; }

        CEWSServices cewsServices = new CEWSServices((CEWSProvider)Provider);
        string errorMessage = string.Empty;
        string folderId = cewsServices.GetFolderID(Path, ref errorMessage);

        if (string.IsNullOrEmpty(folderId))
        {
          throw new Exception($"Unable to get id of folder '{Path}': {errorMessage}");
        }
        else
        {
          return folderId;
        }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public override Folders GetSubFolders(bool lpGetContents)
    {
      try
      {
        if (_subFolders == null) { _subFolders = InitializeSubFolders(); }
        return _subFolders;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public override void Refresh()
    {
      throw new NotImplementedException();
    }

    protected override IFolder GetFolderByPath(string lpFolderPath, long lpMaxContentCount)
    {
      try
      {
        CEWSServices cewsServices = new CEWSServices((CEWSProvider)Provider);
        string errorMessage = string.Empty;
        ObjectValue ceFolder = cewsServices.GetFolderInfo(lpFolderPath, (int)lpMaxContentCount, ref errorMessage);
        return new CEWSFolder(ref ceFolder, (CProvider)Provider);
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    protected override string GetPath()
    {
      try
      {
        if (_ceFolder == null) { return string.Empty; }
        return ((SingletonString)_ceFolder.Property[3]).Value;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    protected override long GetSubFolderCount()
    {
      try
      {
        if (_subFolderCount == -1)
        {
          if (_subFolders == null)
          {
            _subFolderCount = 0;
          }
          else
          {
            _subFolderCount = _subFolders.Count;
          }
        }
        return _subFolderCount;
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

    protected override void InitializeFolder()
    {
      try
      {
        CEWSServices cewsServices = new CEWSServices((CEWSProvider)Provider);
        string errorMessage = string.Empty;
        ECMProperty folderProperty;
        if (_ceFolder == null)
        {
          _ceFolder = cewsServices.GetFolderInfo("/", (int)MaxContentCount, ref errorMessage);
        }

        //  Add all of the available folder properties to the folder object
        foreach (CPEServiceReference.PropertyType propertyType in _ceFolder.Property)
        {
          folderProperty = cewsServices.GetCtsProperty(propertyType);
          if (folderProperty != null) { Properties.Add(folderProperty); }

        }

        AddProperty("FolderClass", _ceFolder.classId);

        Id = _ceFolder.objectId;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private Folders InitializeSubFolders()
    {
      try
      {
        Folders folders = new Folders();
        IFolder folder;
        string folderPath;
        string errorMessage = string.Empty;

        if (_ceFolder == null) { return null; }

        EnumOfObject ceSubFolders;
        ObjectValue ceFolder;

        ceSubFolders = (EnumOfObject)CEWSProvider.GetPropertyByName("SubFolders", _ceFolder);

        if (ceSubFolders.Value == null) { return null; }

        CEWSServices cewsServices = new CEWSServices((CEWSProvider)Provider);

        foreach (ObjectValue subFolder in ceSubFolders.Value)
        {
          folderPath = (string)CEWSProvider.GetPropertyValueByName("PathName", subFolder);
          ceFolder = cewsServices.GetFolderInfo(folderPath, 0, ref errorMessage);
          folder = new CEWSFolder(ref ceFolder, (CProvider)Provider);
          folders.Add((CFolder)folder);
        }

        //  Sort the folders
        folders.Sort();

        return folders;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private void InitializeFolderContents()
    {
      try
      {
        _folderContents = new FolderContents();
        if (_ceFolder != null)
        {
          FolderContent folderContent;
          string containedDocumentName = string.Empty;
          double contentSize;
          DateTime dateLastModified;
          string retrievalName = string.Empty;
          string errorMessage = string.Empty;
          CEWSServices cewsServices;

          EnumOfObject ceFolderContents = (EnumOfObject)CEWSProvider.GetPropertyByName("ContainedDocuments", _ceFolder);
          if (ceFolderContents == null) { throw new InvalidOperationException("Unable to get folder property 'ContainedDocuments'."); }

          //  If there are no contained documents then try to get them else exit
          if (ceFolderContents.Value == null)
          {
            cewsServices = new CEWSServices((CEWSProvider)Provider);
            _ceFolder = cewsServices.GetFolderInfo(Path, 1, ref errorMessage);
            //Changed maxcontentcount from -1 to 1 because below we only care about the first document?
            ceFolderContents = (EnumOfObject)CEWSProvider.GetPropertyByName("ContainedDocuments", _ceFolder);
          }
          if (ceFolderContents == null) { return; }

          foreach (ObjectValue objectValue in ceFolderContents.Value)
          {
            //  Try to get the name of the document
            try
            {
              containedDocumentName = (string)CEWSProvider.GetPropertyValueByName("Name", objectValue);
            }
            catch (ArgumentNullException argNullEx)
            {
              ApplicationLogging.LogException(argNullEx, MethodBase.GetCurrentMethod(), 0, "Name");
              if ((objectValue.objectId != null) && (!string.IsNullOrEmpty(objectValue.objectId)))
              {
                ApplicationLogging.WriteLogEntry($"Unable to get the name of document '{objectValue.objectId}'. The Name property was null.", MethodBase.GetCurrentMethod(), System.Diagnostics.TraceEventType.Warning, 62361);
                containedDocumentName = $"Name Resolution Failed - {objectValue.objectId}";
              }
              else
              {
                ApplicationLogging.WriteLogEntry($"Unable to get the name of document. The Name property was null.", MethodBase.GetCurrentMethod(), System.Diagnostics.TraceEventType.Warning, 62362);
                containedDocumentName = $"Name Resolution Failed";
              }
            }
            catch (Exception ex)
            {
              ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod(), TraceEventType.Warning, 62363);
              containedDocumentName = $"Name Resolution Failure";
            }

            folderContent = new FolderContent(containedDocumentName);
            folderContent.Properties.Add("ID", objectValue.objectId);

            //  Just in case these props are not retrieved
            try
            {
              ListOfObject contentElements = (ListOfObject)CEWSServices.GetPropertyByID(objectValue, "ContentElements");
              foreach (DependentObjectType contentElement in contentElements.Value)
              {
                retrievalName = CEWSServices.GetRetrievalName(contentElement);
                break; // Only get the first
              }
              folderContent.Properties.Add("FileName", retrievalName);
            }
            catch (Exception) { }

            try 
            { 
              dateLastModified = (DateTime)CEWSProvider.GetPropertyValueByName("DateLastModified", objectValue);
              folderContent.Properties.Add("DateLastModified", dateLastModified);
            }
            catch (Exception) { }

            try
            {
              contentSize = (double)CEWSProvider.GetPropertyValueByName("ContentSize", objectValue);
              folderContent.Properties.Add("ContentSize", contentSize);
            }
            catch (Exception) { }

            try
            {
              bool? isReserved = (bool)CEWSProvider.GetPropertyValueByName("IsReserved", objectValue);
              if (isReserved != null)
              {
                folderContent.Properties.Add("IsReserved", isReserved);
              }
              else
              {
                folderContent.Properties.Add("IsReserved", false);
              }
            }
            catch (Exception) { }

            try
            {
              int? majorVersionNumber = (int)CEWSProvider.GetPropertyValueByName("MajorVersionNumber", objectValue);
              if (majorVersionNumber != null)
              {
                folderContent.Properties.Add("MajorVersionNumber", majorVersionNumber);
              }
              else
              {
                folderContent.Properties.Add("MajorVersionNumber", 0);
              }
            }
            catch (Exception) { }

            try
            {
              int? minorVersionNumber = (int)CEWSProvider.GetPropertyValueByName("MinorVersionNumber", objectValue);
              if (minorVersionNumber != null)
              {
                folderContent.Properties.Add("MinorVersionNumber", minorVersionNumber);
              }
              else
              {
                folderContent.Properties.Add("MinorVersionNumber", 0);
              }
            }
            catch (Exception) { }

            _folderContents.Add(folderContent);

          }
        }
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
