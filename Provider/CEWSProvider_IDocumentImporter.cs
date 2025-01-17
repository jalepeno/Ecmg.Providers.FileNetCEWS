using CPEServiceReference;
using Documents.Annotations;
using Documents.Arguments;
using Documents.Core;
using Documents.Migrations;
using Documents.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static Documents.Providers.FileNetCEWS.CEWSServices;

namespace Documents.Providers.FileNetCEWS
{
  public partial class CEWSProvider : IDocumentImporter
  {

    #region IDocumentImporter Implementation

    public bool EnforceClassificationCompliance => throw new NotImplementedException();

    public event DocumentImportedEventHandler DocumentImported;
    public event DocumentImportErrorEventHandler DocumentImportError;
    public event DocumentImportMessageEventHandler DocumentImportMessage;

    public bool ImportDocument(ref ImportDocumentArgs args)
    {
      try
      {
        if (args == null) { throw new ArgumentNullException(nameof(args)); }
        if (args.Document == null) { throw new ArgumentException("No document was referenced in ImportDocumentArgs"); }

        bool addDocumentSuccess = false;
        Object newCEObject = null;
        string errorMessage = string.Empty;
        Core.Document document = args.Document;

        //  Remove any unsettable properties.
        args.Document.RemoveUnsettableProperties(get_DocumentClass(args.Document.DocumentClass));

        switch (args.FilingMode)
        {
          case Core.FilingMode.UnFiled:
            {
              addDocumentSuccess = _cewsServices.AddDocument(ref document, null, true, args.VersionType, ref newCEObject, ref errorMessage);
              break;
            }

          case Core.FilingMode.DocumentFolderPath:
            {
              PathFactory pathFactory = args.PathFactory;
              pathFactory.BaseFolderPath = string.Empty;
              string[] folderPath = args.Document.get_FolderPathArray(pathFactory);
              folderPath = CleanFolderPath(folderPath);
              addDocumentSuccess = _cewsServices.AddDocument(ref document, folderPath, true, args.VersionType, ref newCEObject, ref errorMessage);
              break;
            }

          case Core.FilingMode.BaseFolderPathOnly:
            {
              if (args.PathFactory != null)
              {
                string[] folderPath = new string[1] { args.PathFactory.BaseFolderPath };
                //  Clean the Folder Path of any illegal characters
                folderPath = CleanFolderPath(folderPath);
                addDocumentSuccess = _cewsServices.AddDocument(ref document, folderPath, true, args.VersionType, ref newCEObject, ref errorMessage);
              }
              else
              {
                addDocumentSuccess = _cewsServices.AddDocument(ref document, null, true, args.VersionType, ref newCEObject, ref errorMessage);
              }
              break;
            }

          case Core.FilingMode.BaseFolderPathPlusDocumentFolderPath:
          case Core.FilingMode.DocumentFolderPathPlusBaseFolderPath:
            {
              if (args.PathFactory != null)
              {
                string[] folderPath = args.Document.get_FolderPathArray(args.PathFactory);
                //  Clean the Folder Path of any illegal characters
                folderPath = CleanFolderPath(folderPath);
                addDocumentSuccess = _cewsServices.AddDocument(ref document, folderPath, true, args.VersionType, ref newCEObject, ref errorMessage);
              }
              else
              {
                addDocumentSuccess = _cewsServices.AddDocument(ref document, null, true, args.VersionType, ref newCEObject, ref errorMessage);
              }
                break;
            }

        }

        if (addDocumentSuccess && args.Document.HasAnnotations() && args.SetAnnotations)
        {
          //  Only attempt to import annotations if the add document operation succeeded and annotations are present in the document.
          ImportAnnotations(ref args);
        }

        return addDocumentSuccess;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public void ImportAnnotations(ref ImportDocumentArgs args)
    {
      try
      {
        if (!args.Document.HasAnnotations()) { return; }

        //// Create a new builder for annotation content (XML)
        //AnnotationImporter annotationXmlBuilder = new AnnotationImporter();

        ////  Build a list of properties to exclude on the new folder object that will be returned
        //string[] excludeProps = new string[24];
        //excludeProps[0] = "Owner";
        //excludeProps[1] = "DateLastModified";
        //excludeProps[2] = "ClassDescription";
        //excludeProps[3] = "ObjectStore";
        //excludeProps[4] = "SecurityParent";
        //excludeProps[5] = "versionSeries";
        //excludeProps[6] = "CurrentVersion";
        //excludeProps[7] = "Reservation";
        //excludeProps[8] = "ReleasedVersion";
        //excludeProps[9] = "DocumentLifecyclePolicy";

        //excludeProps[10] = "StoragePolicy";
        //excludeProps[11] = "AuditedEvents";
        //excludeProps[12] = "Permissions";
        //excludeProps[13] = "ActiveMarkings";
        //excludeProps[14] = "Containers";
        //excludeProps[15] = "SourceDocument";
        //excludeProps[16] = "OwnerDocument";
        //excludeProps[17] = "PublicationInfo";
        //excludeProps[18] = "DestinationDocuments";
        //excludeProps[19] = "DependentDocuments";
        //excludeProps[20] = "IgnoreRedirect";
        //excludeProps[21] = "EntryTemplateObjectStoreName";
        //excludeProps[22] = "EntryTemplateLaunchedWorkflowNumber";
        //excludeProps[23] = "EntryTemplateId";

        ////  Obtain a web service object reference to the version series of the document in question
        ////ObjectValue versionSeries = _cewsServices.getv


        //PropertyFilterType propertyFilter = new PropertyFilterType();
        //propertyFilter.IncludeProperties = new FilterElementType[1];
        //propertyFilter.IncludeProperties[0] = new FilterElementType() { Value = "ContentElements" };

        //  For each Version in Document
        int versionIndex = 0;

        foreach (Core.Version ctsVersion in args.Document.Versions)
        {
          //  Get the "version id" from the version collection
          //var p8Document = _cewsServices.GetDocumentProperties(ctsVersion.ID, RequestType.BasicProperties, ContentExportPropertyExclusions.ToArray(), exportContent, exportAnnotations);
          string nativeVersionDocumentId = ctsVersion.Identifier;

          int ctsContentElementIndex = 0;

          foreach (Content content in ctsVersion.Contents)
          {
            //  If this content element does not have an Annotations collection or it is empty, skip to the next content element
            if ((content.Annotations == null) || (content.Annotations.Count == 0))
            {
              ctsContentElementIndex++;
              continue;
            }

            //  For each Annotation in the Annotations of the ContentElement
            foreach (Annotation ctsAnnotation in content.Annotations)
            {
              string annotationErrorMessage = string.Empty;
              ECMProperty CurrentVersionIdProperty = ctsVersion.Properties.GetItemByName("CurrentVersionId");
              string versionId = string.Empty;
              if (CurrentVersionIdProperty != null) 
              { versionId = (string)CurrentVersionIdProperty.Value; }
              else { versionId = args.Document.ObjectID; }


              //  The annotation to be loaded must use a unique id in the repository
              //  The annotation xml contains the GUID id of the annotation

              //  We have to create a new guid id for the annotation and replace the id inside
              //  the annotation xml with the new id before loading into FileNet

              string annotationText = Helper.CopyStreamToString(ctsAnnotation.AnnotatedContent.ToStream());
              string newAnnotationId = $"{{{Guid.NewGuid().ToString()}}}";
              string annotationId = ctsAnnotation.ID;
              annotationText = annotationText.Replace(annotationId, newAnnotationId);
              Stream annotationStream = Helper.CopyStringToStream(annotationText);

              //  Create/persist a new annotation object for the document version and content element index number.
              bool addSuccess = _cewsServices.AddAnnotation(newAnnotationId, versionId, annotationStream, ref annotationErrorMessage);

              if (addSuccess) 
              { 
                ApplicationLogging.LogInformation($"Successfully added annotation '{annotationId}' as '{newAnnotationId}'."); 
              }
              else 
              { 
                ApplicationLogging.WriteLogEntry($"Annotation {annotationId} add failed: '{annotationErrorMessage}'", MethodBase.GetCurrentMethod(), TraceEventType.Error, 34876); 
              }
            }
          }

        }

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    internal ObjectValue CreateAnnotation()
    {
      ObjectValue nativeAnnotation = new ObjectValue() { classId = "Annotation" };
      try
      {
        Factory.CreateAction("Annotation");
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
      return nativeAnnotation;
    }

    public void OnDocumentImported(ref DocumentImportedEventArgs e)
    {
      throw new NotImplementedException();
    }

    public void OnDocumentImportError(ref DocumentImportErrorEventArgs e)
    {
      throw new NotImplementedException();
    }

    #endregion

  }
}
