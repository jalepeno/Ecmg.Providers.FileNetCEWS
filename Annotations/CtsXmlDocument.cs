using Documents.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Documents.Providers.FileNetCEWS.Annotations
{
  internal class CtsXmlDocument : XmlDocument
  {

    #region XML Helpers

    /// <summary>
    /// Queries the XML response for the existence of a particular node.
    /// </summary>
    /// <param name="xpath">The XML path.</param>
    /// <returns>True if exists, otherwise false.</returns>
    public bool QueryExists(string xpath)
    {
      try
      {
        if (string.IsNullOrEmpty(xpath)) throw new ArgumentNullException(nameof(xpath));

        XmlNodeList nodes = this.SelectNodes(xpath);

        if ((nodes == null) || (nodes.Item(0) == null))
        {
          return false;
        }
        else
        {
          return true;
        }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    /// <summary>
    /// Queries the XML for a single string.
    /// </summary>
    /// <param name="xpath">The XML path.</param>
    /// <returns>The text of a single node.</returns>
    public string QuerySingleString(string xpath)
    {
      string result = null;
      try
      {
        if (string.IsNullOrEmpty(xpath)) throw new ArgumentNullException(nameof(xpath));
        XmlNode node = this.SelectSingleNode(xpath);
        result = node.InnerText;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
      }
      return result;
    }

    public bool QuerySingleBoolean(string xpath)
    {
      bool result = false;
      try
      {
        if (string.IsNullOrEmpty(xpath)) throw new ArgumentNullException(nameof(xpath));
        string value = this.QuerySingleString(xpath);
        if (string.IsNullOrEmpty(value)) return false;
        if (value.ToLowerInvariant().CompareTo("true") == 0) { result = true; }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
      }
      return result;
    }

    public DateTime QuerySingleDate(string xpath)
    {
      DateTime result = DateTime.MinValue;
      try
      {
        if (string.IsNullOrEmpty(xpath)) throw new ArgumentNullException(nameof(xpath));
        string value = this.QuerySingleString(xpath);
        if (string.IsNullOrEmpty(value)) return result;
        if (!DateTime.TryParse(value, out result))
        {
          ApplicationLogging.WriteLogEntry($"Could not parse date {value}");
        }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
      }
      return result;
    }

    public int QuerySingleInteger(string xpath)
    {
      int result = -100;
      try
      {
        if (string.IsNullOrEmpty(xpath)) throw new ArgumentNullException(nameof(xpath));
        string value = this.QuerySingleString(xpath);
        if (string.IsNullOrEmpty(value)) return result;
        if (!int.TryParse(value, out result))
        {
          ApplicationLogging.WriteLogEntry($"Could not parse integer {value}");
        }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
      }
      return result;
    }

    /// <summary>
    /// Queries the XML for the text of an attribute.
    /// </summary>
    /// <param name="xpath">The XML path.</param>
    /// <param name="attribute">The XML attribute.</param>
    /// <returns>The text of the specified attribute.</returns>
    public string QuerySingleAttribute(string xpath, string attribute)
    {
      string result = null;
      try
      {
        if (string.IsNullOrEmpty(xpath)) throw new ArgumentNullException(nameof(xpath));
        if (string.IsNullOrEmpty(attribute)) throw new ArgumentNullException(nameof(attribute));

        XmlNode node = this.SelectSingleNode(xpath);
        if (node == null) return result;

        XmlAttribute item = node.Attributes[attribute];
        if (item == null) return result;

        result = item.Value;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
      }
      return result;
    }

    public DateTime QuerySingleAttributeAsDate(string xpath, string attribute)
    {
      DateTime result = new DateTime();
      try
      {
        if (string.IsNullOrEmpty(xpath)) throw new ArgumentNullException(nameof(xpath));
        if (string.IsNullOrEmpty(attribute)) throw new ArgumentNullException(nameof(attribute));

        string value = this.QuerySingleAttribute(xpath, attribute);
        if (!DateTime.TryParse(value, out result))
        {
          ApplicationLogging.WriteLogEntry($"Could not parse date {value}");
        }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
      }
      return result;
    }

    public int QuerySingleAttributeAsInteger(string xpath, string attribute)
    {
      int result = new int();
      try
      {
        if (string.IsNullOrEmpty(xpath)) throw new ArgumentNullException(nameof(xpath));
        if (string.IsNullOrEmpty(attribute)) throw new ArgumentNullException(nameof(attribute));

        string value = this.QuerySingleAttribute(xpath, attribute);
        if (!int.TryParse(value, out result))
        {
          ApplicationLogging.WriteLogEntry($"Could not parse integer {value}");
        }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
      }
      return result;
    }

    public Single QuerySingleAttributeAsSingle(string xpath, string attribute)
    {
      Single result = new Single();
      try
      {
        if (string.IsNullOrEmpty(xpath)) throw new ArgumentNullException(nameof(xpath));
        if (string.IsNullOrEmpty(attribute)) throw new ArgumentNullException(nameof(attribute));

        string value = this.QuerySingleAttribute(xpath, attribute);
        if (!Single.TryParse(value, out result))
        {
          ApplicationLogging.WriteLogEntry($"Could not parse integer {value}");
        }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
      }
      return result;
    }

    public bool QuerySingleAttributeAsBoolean(string xpath, string attribute)
    {
      bool result = false;
      try
      {
        if (string.IsNullOrEmpty(xpath)) throw new ArgumentNullException(nameof(xpath));
        if (string.IsNullOrEmpty(attribute)) throw new ArgumentNullException(nameof(attribute));

        string value = this.QuerySingleAttribute(xpath, attribute);
        if (value.ToLowerInvariant().CompareTo("true") == 0) { result = true; }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
      }
      return result;
    }

    #endregion

  }
}
