using CPEServiceReference;
using Documents.Core;
using Documents.Exceptions;
using Documents.Providers.FileNetCEWS.Configuration;
using Documents.Providers.FileNetCEWS.Provider;
using Documents.Utilities;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Documents.Providers.FileNetCEWS
{
  public partial class CEWSProvider : CProvider
  {

    #region Class Constants

    private const string PROVIDER_NAME = "P8 Content Engine Web Services Provider";
    private const string PROVIDER_SYSTEM_TYPE = "FileNet P8 5.2 and above";
    private const string PROVIDER_COMPANY_NAME = "IBM";
    private const string PROVIDER_PRODUCT_NAME = "FileNet P8 Content Engine";
    private const string PROVIDER_PRODUCT_VERSION = "5.5";

    private const string SERVER_NAME = "ServerName";
    private const string PORT_NUMBER = "PortNumber";
    private const string PROTOCOL = "Protocol";
    private const string TRUSTED_CONNECTION = "TrustedConnection";
    private const string USER_NAME = "UserName";
    private const string PASSWORD = "Password";
    private const string OBJECT_STORE = "ObjectStore";

    #endregion

    #region Enumerations

    public enum VersionStatus
    {
      Released = 1,
      InProcess = 2,
      Reservation = 3,
      Superceded = 4
    }

    #endregion

    #region Class Variables

    private ProviderSystem _providerSystem = new ProviderSystem(PROVIDER_NAME, PROVIDER_SYSTEM_TYPE, PROVIDER_COMPANY_NAME,
                         PROVIDER_PRODUCT_NAME, PROVIDER_PRODUCT_VERSION);

    private string _serverName;
    private string _url;
    private string _protocol = "https";
    private int? _portNumber;
    private bool _trustedConnection;
    private string _objectStoreName;
    //private FNCEWS40PortTypeClient _client;
    private CEWSServices _cewsServices;
    Settings _cewsSettings;
    private List<string> _contentExportPropertyExclusions;
    private CEWSSearch _search;

    #endregion

    #region Public Properties

    public int? PortNumber
    {
      get
      {
        return _portNumber;
      }
      set
      {
        _portNumber = value;
      }
    }

    public string Protocol
    {
      get
      {
        return _protocol;
      }
      set
      {
        _protocol = value;
      }
    }

    public string ServerName
    {
      get
      {
        return _serverName;
      }
      set
      {
        _serverName = value;
      }
    }

    public bool TrustedConnection
    {
      get
      {
        return _trustedConnection;
      }
      set
      {
        _trustedConnection = value;
      }
    }

    public string ObjectStoreName
    {
      get
      {
        return _objectStoreName;
      }
      set
      {
        _objectStoreName = value;
      }
    }

    public override ProviderSystem ProviderSystem
    {
      get
      {
        return _providerSystem;
      }
    }

    public string URL
    {
      get
      {
        try
        {
          if (string.IsNullOrEmpty(_url))
          {
            _url = CreateURL();
          }
          return _url;
        }
        catch (Exception ex)
        {
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
          //  Re - throw the exception to the caller
          throw;
        }
      }
    }

    public List<string> ContentExportPropertyExclusions
    {
      get
      {
      try
      {
          if (_contentExportPropertyExclusions == null)
          {
            _contentExportPropertyExclusions = GetAllContentExportPropertyExclusions();
          }
          return _contentExportPropertyExclusions;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
      }
    }

    #endregion

    #region Constructors

    public CEWSProvider() 
    {
      try
      {
        AddProperties();
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    public CEWSProvider(string connectionString)
    {
      try
      {
        AddProperties();
        //ParseConnectionString();
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    #endregion

    #region CProvider Implementation

    public override void Connect(ContentSource contentSource)
    {
      try
      {
        //base.Connect(ContentSource);
        InitializeProvider(contentSource);
        InitializeProperties();
        GetSettings();

        //string errorMessage = string.Empty;
        //if (!SoapServerAvailable(ref errorMessage))
        //{
        //  IsConnected = false;
        //  SetState(ProviderConnectionState.Unavailable);
        //  if (errorMessage.Length > 0) { errorMessage = $"Soap Server not available: {errorMessage}"; }
        //  ApplicationLogging.WriteLogEntry(errorMessage, MethodBase.GetCurrentMethod(), TraceEventType.Error, 404);
        //  throw new RepositoryNotAvailableException(contentSource.Name, errorMessage);
        //}

        //_client = WSIUtil.ConfigureBinding(UserName, ProviderPassword, URL);
        //Localization localization = WSIUtil.GetLocalization();

        _cewsServices = new CEWSServices(this);

        IsConnected = _cewsServices.TestConnection();

        if (!IsConnected)
        {
          SetState(ProviderConnectionState.Unavailable);
          ApplicationLogging.WriteLogEntry($"Login to '{Name}' failed.", MethodBase.GetCurrentMethod(), TraceEventType.Error, 200);
          throw new RepositoryNotAvailableException(Name);
        }

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public override ISearch Search
    {
      get
      {
        try
        {
          if (_search  == null) { _search = new CEWSSearch(this); }
          return _search;
        }
        catch (Exception ex)
        {
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
          //  Re - throw the exception to the caller
          throw;
        }
      }
    }

    public override string FolderDelimiter => throw new NotImplementedException();

    public override bool LeadingFolderDelimiter => throw new NotImplementedException();

    public override ISearch CreateSearch()
    {
      throw new NotImplementedException();
    }

    public override IFolder GetFolder(string folderPath, long maxContentCount)
    {
      try
      {
        string errorMessage = string.Empty;
        ObjectValue ceFolder = _cewsServices.GetFolder(folderPath, (int)maxContentCount, ref errorMessage);

        if (!string.IsNullOrEmpty(errorMessage)) { throw new InvalidPathException($"Folder '{folderPath}' not available.  {errorMessage}", folderPath); }

        errorMessage = string.Empty;
        return new CEWSFolder(ref ceFolder, this);

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }





    /// <summary>
    /// Iterates through all of the properties of the specified object to find the named property and returns the property value.
    /// </summary>
    /// <param name="propertyName">The symbolic name of the property to find.</param>
    /// <param name="item">An ObjectValue object from the CEWSI API.</param>
    /// <returns></returns>
    /// <remarks>If the property is not found using the supplied property name an InvalidOperationException is raised.</remarks>
    protected internal static object GetPropertyValueByName(string propertyName, object item)
    {
      try
      {
        CPEServiceReference.PropertyType property = GetPropertyByName(propertyName, item);
        if (property == null)
          return null;
        else
        {
          PropertyInfo propertyInfo = property.GetType().GetProperty("Value");
          if (propertyInfo == null) { throw new PropertyDoesNotExistException(propertyName); }
          return propertyInfo.GetValue(property, null);
        }
      }

      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        // Re-throw the exception to the caller
        throw;
      }
    }

    /// <summary>
    /// Iterates through all of the properties of the specified object to find the named property and returns the property object.
    /// </summary>
    /// <param name="propertyName">The symbolic name of the property to find.</param>
    /// <param name="object">An ObjectValue object from the CEWSI API.</param>
    /// <returns>A PropertyType based object</returns>
    /// <remarks>If the property is not found using the supplied property name an InvalidOperationException is raised.</remarks>
    protected internal static CPEServiceReference.PropertyType GetPropertyByName(string propertyName, object @object)
    {
      try
      {
        return CEWSServices.GetPropertyByName(propertyName, (ObjectValue)@object);
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        // Re-throw the exception to the caller
        throw;
      }
    }


    internal new void SetState(ProviderConnectionState state)
    {
      try
      {
        base.SetState(state);
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

    #region Provider Identification

    private void AddProperties()
    {
      try
      {

        // Add the 'ServerName' property
        ProviderProperties.Add(new ProviderProperty(SERVER_NAME, typeof(string), true, string.Empty , 3, string.Empty, false, true));

        // Add the 'Port' property
        ProviderProperties.Add(new ProviderProperty(PORT_NUMBER, typeof(string), true, null, 8, string.Empty, false, true));

        // Add the 'Protocol' property
        ProviderProperties.Add(new ProviderProperty(PROTOCOL, typeof(string), true, "https", 8, string.Empty, false, true));

        //  Add the 'TrustedConnection' property
        ProviderProperties.Add(new ProviderProperty(TRUSTED_CONNECTION, typeof(bool), false, "false", 5, string.Empty, false, true));

        //  Add the 'UserName' property
        ProviderProperties.Add(new ProviderProperty(USER_NAME, typeof(string), false, string.Empty, 6, string.Empty, false, true));

        //  Add the 'Password' property
        ProviderProperties.Add(new ProviderProperty(PASSWORD, typeof(string), false, string.Empty, 7, string.Empty, false, true));

        //  Add the 'ObjectStore' property
        ProviderProperties.Add(new ProviderProperty(OBJECT_STORE, typeof(string), true, string.Empty, 8, string.Empty, true, true));

        // Sort the provider properties by the sequence number
        ProviderProperties.Sort(new ProviderProperties.ProviderPropertySequenceComparer());

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    #endregion

    private static string[] CleanFolderPath(string[] originalFolderPaths, string replacementValue = "-")
    {
      try
      {
        string[] cleanFolderPaths = (string[])originalFolderPaths.Clone();

        for (int pathCounter = 0; pathCounter < cleanFolderPaths.Length; pathCounter++)
        {
          if (!string.IsNullOrEmpty(cleanFolderPaths[pathCounter]))
          {
            cleanFolderPaths[pathCounter] = CleanFolderPath(cleanFolderPaths[pathCounter], replacementValue);
          }            
        }

        return cleanFolderPaths;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    /// <summary>
    /// Substitutes a replacement value for illegal characters in a folder name.
    /// </summary>
    /// <param name="originalFolderPath"></param>
    /// <param name="replacementValue"></param>
    /// <returns></returns>
    /// <remarks>Looks for \ / : * ? " < > | and replaces with value specified in lpReplacementValue.</remarks>
    private static string CleanFolderPath(string originalFolderPath, string replacementValue = "-")
    {

      if (string.IsNullOrEmpty(originalFolderPath)) { throw new ArgumentNullException(nameof(originalFolderPath)); }

      //  The name cannot be an empty string and cannot contain any of the following characters: \ / : * ? " < > |

      // Note: we are no longer subtituting  on '\' as there is no way to determine valid vs. invalid instances of this character in the complete folder path.
      // These will need to be cleaned in the transformation before sending to the import provider.
      try
      {
        Regex regex = new Regex("\\b\\x7C|\\x5C|\\x3A|\\x2A|\\x3F|\\x22|\\x3C|\\x3E|\\x7C\\b", RegexOptions.Compiled);
        return regex.Replace(originalFolderPath, replacementValue);
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    private string CreateURL()
    {
      // Create a URL that looks like the following
      // "https://localhost:9443/wsi/FNCEWS40MTOM/";
      try
      {
        if ((PortNumber == null) || (PortNumber == 80))
        {
          return $"{Protocol}://{ServerName}/wsi/FNCEWS40MTOM/";
        }
        else
        {
          return $"{Protocol}://{ServerName}:{PortNumber}/wsi/FNCEWS40MTOM/";
        }
        
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    protected bool DocumentExists(string id, string classId = "Document")
    {
      try
      {
        ObjectValue documentObject = _cewsServices.GetObject(id, classId);
        if ((documentObject != null) && (documentObject.objectId != null) && (documentObject.objectId.Length > 0))
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

    internal List<string> GetAllContentExportPropertyExclusions()
    {
      try
      {
        if (_cewsSettings == null) { GetSettings(); }

        List<string> exclusions = new List<string>();

        exclusions.AddRange(_cewsSettings.AppContentExportPropertyExclusions);
        exclusions.AddRange(_cewsSettings.UserContentExportPropertyExclusions);

        return exclusions;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private void GetSettings()
    {
      try
      {

        Stream resourceStream = null;
        Assembly assembly = Assembly.GetExecutingAssembly();
        string[] resourceNames = assembly.GetManifestResourceNames();
        foreach (string resourceName in resourceNames)
        {
          if (resourceName.Contains("cewssettings.json"))
            resourceStream = assembly.GetManifestResourceStream(resourceName);
        }

        IConfigurationRoot appConfig = new ConfigurationBuilder()
        .SetBasePath($"{Directory.GetCurrentDirectory()}\\Configuration")
        .AddJsonStream(resourceStream)        
        .AddEnvironmentVariables()
        .Build();

        // Get values from the config given their key and their target type.
        _cewsSettings = appConfig.GetRequiredSection("Settings").Get<Settings>();

        if (_cewsSettings == null) { throw new InvalidOperationException("Unable to read settings file."); }

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private void InitializeProperties()
    {
      try
      {
        foreach (ProviderProperty property in ProviderProperties)
        {
          switch (property.PropertyName)
          {
            case SERVER_NAME: 
              { 
                ServerName = property.PropertyValue.ToString(); 
                break; 
              }

            case PORT_NUMBER:
              {
                if (property.PropertyValue != null)
                {
                  int portnumber;
                  if (int.TryParse(property.PropertyValue.ToString(), out portnumber))
                  {
                    PortNumber = portnumber;
                  }                  
                }                
                break;
              }

            case PROTOCOL:
              {
                Protocol = property.PropertyValue.ToString();
                break;
              }

            case TRUSTED_CONNECTION:
              {
                TrustedConnection = bool.Parse(property.PropertyValue.ToString());
                break;
              }

            case USER_NAME:
              {
                UserName = property.PropertyValue.ToString();
                break;
              }

            case PASSWORD:
              {
                ProviderPassword = property.PropertyValue.ToString();
                break;
              }

            case OBJECT_STORE:
              {
                ObjectStoreName = property.PropertyValue.ToString();
                break;
              }
            default:
              break;
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

    //private void ParseConnectionString()
    //{
    //  try
    //  {

    //  }
    //  catch (Exception ex)
    //  {
    //    ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
    //    //  Re-throw the exception to the caller
    //    throw;
    //  }
    //}

    private bool SoapServerAvailable(ref string errorMessage)
    {
      try
      {
        bool canPing = false;
        Ping ping = new Ping();
        PingReply reply = ping.Send(ServerName, 5000);
        if (reply != null) { canPing = true; }

        if (!canPing)
        {
          errorMessage = $"Unable to ping Content Engine server at '{ServerName}', verify that server name and network connection is correct.";
          ApplicationLogging.LogWarning(errorMessage, MethodBase.GetCurrentMethod());
          return false;
        }

        return canPing;

      }
      catch (PingException pingEx)
      {
        if ((pingEx.InnerException != null) && (pingEx.InnerException.Message == "No such host is known"))
        {
          errorMessage = $"Host '{ServerName}' Unknown, make sure that the server name supplied is correct. If so ensure that a network connection is available and DNS is functioning.";
          ApplicationLogging.LogWarning(errorMessage, MethodBase.GetCurrentMethod());
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

    #endregion

  }
}
