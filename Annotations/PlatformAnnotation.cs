using Documents.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Documents.Providers.FileNetCEWS.Annotations
{
  internal class PlatformAnnotation
  {

    #region Class Variables

    private readonly string _className;
    private readonly string _classId;
    private readonly string _subClassName;
    private readonly Type _annotationType;

    #endregion

    #region Internal Properties

    public string ClassName { get { return _className; } }
    public string SubClassName { get { return _subClassName; } }
    public string ClassId { get { return _classId; } }
    public Type  AnnotationType { get { return _annotationType; } }

    #endregion

    #region Constructors

    public PlatformAnnotation(string classId, string className, Type annotationType)
    {
      try
      {
        if (string.IsNullOrEmpty(classId)) throw new ArgumentNullException(nameof(classId));
        if (string.IsNullOrEmpty(className)) throw new ArgumentNullException(nameof(className));
        if (annotationType == null) throw new ArgumentNullException(nameof(annotationType));

        _classId = classId;
        _className = className;
        _annotationType = annotationType;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public PlatformAnnotation(string classId, string className, string subClassName, Type annotationType)
    {
      try
      {
        if (string.IsNullOrEmpty(classId)) throw new ArgumentNullException(nameof(classId));
        if (string.IsNullOrEmpty(className)) throw new ArgumentNullException(nameof(className));
        if (string.IsNullOrEmpty(subClassName)) throw new ArgumentNullException(nameof(subClassName));
        if (annotationType == null) throw new ArgumentNullException(nameof(annotationType));

        _classId = classId;
        _className = className;
        _subClassName = subClassName;
        _annotationType = annotationType;

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
