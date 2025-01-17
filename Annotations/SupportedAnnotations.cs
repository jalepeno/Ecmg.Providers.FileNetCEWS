using Documents.Annotations.Common;
using Documents.Annotations.Highlight;
using Documents.Annotations.Shape;
using Documents.Annotations.Special;
using Documents.Annotations.Text;
using Documents.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Documents.Providers.FileNetCEWS.Annotations
{
  internal class SupportedAnnotations
  {

    #region Class Variables

    private readonly Collection<PlatformAnnotation> _items;
    private static readonly SupportedAnnotations _instance = new SupportedAnnotations();

    #endregion

    #region Internal Properties

    internal ReadOnlyCollection<PlatformAnnotation> Items { get { return new ReadOnlyCollection<PlatformAnnotation>(_items); } }

    internal static SupportedAnnotations Instance { get { return _instance; } }

    #endregion

    #region Constructors

    private SupportedAnnotations()
    {
      try
      {
        _items = new Collection<PlatformAnnotation>();
        _items.Add(new PlatformAnnotation("{5CF11946-018F-11D0-A87A-00A0246922A5}", "Arrow", typeof(ArrowAnnotation)));
        _items.Add(new PlatformAnnotation("{5CF1194C-018F-11D0-A87A-00A0246922A5}", "Stamp", typeof(StampAnnotation)));
        _items.Add(new PlatformAnnotation("{5CF11945-018F-11D0-A87A-00A0246922A5}", "StickyNote", typeof(StickyNoteAnnotation)));
        _items.Add(new PlatformAnnotation("{5CF11941-018F-11D0-A87A-00A0246922A5}", "Text", typeof(TextAnnotation)));
        _items.Add(new PlatformAnnotation("{5CF11942-018F-11D0-A87A-00A0246922A5}", "Highlight", typeof(HighlightRectangle)));
        _items.Add(new PlatformAnnotation("{A91E5DF2-6B7B-11D1-B6D7-00609705F027}", "Proprietary", "v1-Rectangle", typeof(RectangleAnnotation)));
        _items.Add(new PlatformAnnotation("{A91E5DF2-6B7B-11D1-B6D7-00609705F027}", "Proprietary", "v1-Line", typeof(PointCollectionAnnotation)));
        _items.Add(new PlatformAnnotation("{A91E5DF2-6B7B-11D1-B6D7-00609705F027}", "Proprietary", "v1-Oval", typeof(EllipseAnnotation)));
        _items.Add(new PlatformAnnotation("{5CF11949-018F-11D0-A87A-00A0246922A5}", "Pen", typeof(PointCollectionAnnotation)));
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
