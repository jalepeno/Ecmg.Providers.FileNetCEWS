using Documents.Annotations;
using Documents.Annotations.Common;
using Documents.Annotations.Auditing;
using Documents.Annotations.Security;
using Documents.Utilities;
using Microsoft.IdentityModel.Tokens;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Documents.Annotations.Decoration;
using Documents.Annotations.Text;
using Documents.Annotations.Shape;
using System.Drawing.Drawing2D;
using Documents.Annotations.Highlight;
using Documents.Annotations.Special;
using System.Xml;
using Documents.Transformations;
using Documents.Annotations.Exception;

namespace Documents.Providers.FileNetCEWS.Annotations
{
  public class AnnotationExporter
  {

    #region Class Constants

    private const string PROP_DESC_PATH = "/FnAnno/PropDesc";
    private const string TEXT_PATH = PROP_DESC_PATH + "/F_TEXT";

    #endregion

    #region Public Properties

    public Double ScaleX { get; set; } = 96.0;

    public Double ScaleY { get; set; } = 96.0;

    #endregion

    #region Public Methods

    public Annotation ExportAnnotationObject(Stream annotation, string mimeType)
    {
      Annotation result = null;
      try
      {
        if (annotation == null) throw new ArgumentNullException(nameof(annotation));
        if (string.IsNullOrEmpty(mimeType)) throw new ArgumentNullException(nameof(mimeType));

        CtsXmlDocument annotationXml = new CtsXmlDocument();
        annotationXml.Load(annotation); 

        Type annotationType = GetAnnotationType(annotationXml);
        if (annotationType == null)
        {
          ApplicationLogging.WriteLogEntry("Could not determine CTS annotation class", MethodBase.GetCurrentMethod(), System.Diagnostics.TraceEventType.Warning, 11903);
          return result;
        }

        bool processed = false;
        bool isMultiPageTiff = false;
        string normalizedMimeType = mimeType.ToLowerInvariant();
        if ((normalizedMimeType == "image/tiff") || (normalizedMimeType.Equals("image/x-tiff"))) { isMultiPageTiff = true; }

        if (annotationType == typeof(StampAnnotation))
        {
          result = new StampAnnotation();
          ExportCommonMetadata(result, annotationXml, isMultiPageTiff);
          ExportStamp((StampAnnotation)result, annotationXml);
          processed = true;
          return result;
        }

        if (annotationType == typeof(StickyNoteAnnotation))
        {
          result = new StickyNoteAnnotation();
          ExportCommonMetadata(result, annotationXml, isMultiPageTiff);
          ExportStickyNote((StickyNoteAnnotation)result, annotationXml);
          processed = true;
          return result;
        }

        if (annotationType == typeof(TextAnnotation))
        {
          result = new TextAnnotation();
          ExportCommonMetadata(result, annotationXml, isMultiPageTiff);
          ExportText((TextAnnotation)result, annotationXml);
          processed = true;
          return result;
        }

        if (annotationType == typeof(PointCollectionAnnotation))
        {
          result = new PointCollectionAnnotation();
          ExportCommonMetadata(result, annotationXml, isMultiPageTiff);
          ExportPointCollection((PointCollectionAnnotation)result, annotationXml);
          processed = true;
          return result;
        }

        if (!processed) throw new UnsupportedAnnotationException("The annotation was not recognized during export.");

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
      return result;
    }

    #endregion

    #region Private Methods

    #region Annotation Exports

    private static Type GetAnnotationType(CtsXmlDocument xmlAnnotation)
    {
      Type result = null;
      try
      {
        if (xmlAnnotation == null) throw new ArgumentNullException(nameof(xmlAnnotation));

        string classId = xmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_CLASSID");
        string className = xmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_CLASSNAME");
        string subClassName = xmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_SUBCLASS");

        foreach (PlatformAnnotation item in SupportedAnnotations.Instance.Items)
        {
          if (item.ClassId.CompareTo(classId) != 0) { continue; }
          if (item.ClassName.CompareTo(className) != 0) { continue; }
          if (string.IsNullOrEmpty(item.SubClassName))
          {
            result = item.AnnotationType; break;
          }

          if (item.SubClassName.CompareTo(subClassName) != 0) { continue; }
          result = item.AnnotationType;
        }

        if (result == null)
        {
          if (string.IsNullOrEmpty(subClassName)) { subClassName = string.Empty; }
          string errorMessage = $"Could not map annotation type.  ClassId='{classId}', ClassName='{className}', SubClass='{subClassName}'";
          ApplicationLogging.WriteLogEntry(errorMessage, MethodBase.GetCurrentMethod(), System.Diagnostics.TraceEventType.Warning, 11404);
        }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
      return result;
    }

    private void ExportCommonMetadata(Annotation annotation, CtsXmlDocument xmlAnnotation, bool isMultiPageTiff)
    {
      try
      {
        if (annotation == null) throw new ArgumentNullException(nameof(annotation));
        if (xmlAnnotation == null) throw new ArgumentNullException(nameof(xmlAnnotation));

        annotation.ID = xmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_ANNOTATEDID");

        annotation.AuditEvents.Created = new CreateEvent() { EventTime = xmlAnnotation.QuerySingleAttributeAsDate(PROP_DESC_PATH, "F_ENTRYDATE") };
        annotation.AuditEvents.Modified = new ModifyEvent() { EventTime = xmlAnnotation.QuerySingleAttributeAsDate(PROP_DESC_PATH, "F_MODIFYDATE") };

        if (isMultiPageTiff) 
        {
          annotation.Layout.PageNumber = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_MULTIPAGETIFFPAGENUMBER");
        }
        else
        {
          annotation.Layout.PageNumber = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_PAGENUMBER");
        }

        annotation.Layout.UpperLeftExtent = new Point((float)(ScaleX * xmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LEFT")), 
          (float)(ScaleY * xmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_TOP")));

        annotation.Layout.LowerRightExtent = new Point((float)(ScaleX * xmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_WIDTH") + annotation.Layout.UpperLeftExtent.First),
          (float)(ScaleY * xmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_HEIGHT") + annotation.Layout.UpperLeftExtent.Second));

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private void ExportArrow(ArrowAnnotation annotation, CtsXmlDocument xmlAnnotation)
    {
      try
      {
        if (annotation == null) throw new ArgumentNullException(nameof(annotation));
        if (xmlAnnotation == null) throw new ArgumentNullException(nameof(xmlAnnotation));

        ExportLineMetadata(annotation, xmlAnnotation);
        annotation.Size = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_ARROWHEAD_SIZE");
        annotation.StartPoint = new Point((float)(xmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_START_X") * ScaleX), 
          (float)(xmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_START_Y") * ScaleY));

        annotation.EndPoint = new Point((float)(xmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_END_X") * ScaleX),
          (float)(xmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_END_Y") * ScaleY));

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private void ExportEllipse(EllipseAnnotation annotation, CtsXmlDocument xmlAnnotation)
    {
      try
      {
        if (annotation == null) throw new ArgumentNullException(nameof(annotation));
        if (xmlAnnotation == null) throw new ArgumentNullException(nameof(xmlAnnotation));

        annotation.Display.Foreground = ParseColor(xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_COLOR"));
        if (xmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_LINE_BACK_MODE") == null)
        {
          annotation.Display.Foreground.Opacity = 100;
        }
        else
        {
          int lineBackMode = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_BACK_MODE");
          annotation.Display.Foreground.Opacity = 50 * lineBackMode;
        }

        int patternCode = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_STYLE");
        var linePattern = GetLinePattern(patternCode);
        annotation.LineStyle.Pattern = linePattern;

        int lineWeight = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_WIDTH");
        annotation.LineStyle.LineWeight = lineWeight;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private void ExportHighlightRectangle(HighlightRectangle annotation, CtsXmlDocument xmlAnnotation)
    {
      try
      {
        if (annotation == null) throw new ArgumentNullException(nameof(annotation));
        if (xmlAnnotation == null) throw new ArgumentNullException(nameof(xmlAnnotation));

        ExportTextBackgroundMode(annotation, xmlAnnotation);

        annotation.HighlightColor = ParseColor(xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_BRUSHCOLOR"));

        if (string.IsNullOrEmpty(xmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_LINE_WIDTH"))) { return; }

        //  optional
        int borderWidth = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_WIDTH");
        if (borderWidth != 0)
        {
          int borderColor = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_COLOR");
          annotation.Display.Border = new BorderInfo() { Color = ParseColor(borderColor), LineStyle = new LineStyleInfo() { LineWeight = borderWidth, Pattern = LineStyleInfo.LinePattern.Solid } };
        }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private void ExportRectangle(RectangleAnnotation annotation, CtsXmlDocument xmlAnnotation)
    {
      try
      {
        if (annotation == null) throw new ArgumentNullException(nameof(annotation));
        if (xmlAnnotation == null) throw new ArgumentNullException(nameof(xmlAnnotation));

        if (xmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_BRUSHCOLOR") != null)
        {
          annotation.Display.Foreground = ParseColor(xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_BRUSHCOLOR"));
          annotation.Display.Foreground.Opacity = 50 * xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_TEXT_BACKMODE");
        }
        else
        {
          annotation.Display.Foreground = null;
        }

        annotation.Display.Border = new BorderInfo() { Color = ParseColor(xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_COLOR")) };
        annotation.LineStyle.LineWeight = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_WIDTH");
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private void ExportStamp(StampAnnotation annotation, CtsXmlDocument xmlAnnotation)
    {
      try
      {
        if (annotation == null) throw new ArgumentNullException(nameof(annotation));
        if (xmlAnnotation == null) throw new ArgumentNullException(nameof(xmlAnnotation));

        ExportBorderInfo(annotation, xmlAnnotation);
        annotation.TextElement = new TextMarkup();
        ExportFontMetadata(annotation.TextElement, xmlAnnotation);
        annotation.TextElement.TextRotation = (float)(360.0 - xmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_ROTATION"));
        ExportTextBackgroundMode(annotation, xmlAnnotation);
        ExportTextMetadata(annotation.TextElement, xmlAnnotation);
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private void ExportStickyNote(StickyNoteAnnotation annotation, CtsXmlDocument xmlAnnotation)
    {
      try
      {
        if (annotation == null) throw new ArgumentNullException(nameof(annotation));
        if (xmlAnnotation == null) throw new ArgumentNullException(nameof(xmlAnnotation));

        //  Daeja on P8 does not support setting the color of the stickynote, but it does specify the color.
        //  F_FORECOLOR = "10092543"
        annotation.Display.Background = ParseColor(xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_FORECOLOR"));

        //  F_ORDINAL is not used in P8.  CS and IS used this.
        //  annotation.NoteOrder = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_ORDINAL")
        ExportTextMetadata(annotation.TextNote, xmlAnnotation);

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private void ExportText(TextAnnotation annotation, CtsXmlDocument xmlAnnotation)
    {
      try
      {
        if (annotation == null) throw new ArgumentNullException(nameof(annotation));
        if (xmlAnnotation == null) throw new ArgumentNullException(nameof(xmlAnnotation));

        TextMarkup textMarkup = new TextMarkup();
        ExportBorderInfo(annotation, xmlAnnotation);
        ExportFontMetadata(textMarkup, xmlAnnotation);
        ExportTextBackgroundMode(annotation, xmlAnnotation);
        ExportTextMetadata(textMarkup, xmlAnnotation);
        annotation.TextMarkups.Add(textMarkup);

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private void ExportPointCollection(PointCollectionAnnotation annotation, CtsXmlDocument xmlAnnotation)
    {
      try
      {
        if (annotation == null) throw new ArgumentNullException(nameof(annotation));
        if (xmlAnnotation == null) throw new ArgumentNullException(nameof(xmlAnnotation));

        //  PointCollection maps to more than one annotation type.  We need to determine the type for the target
        string subClassName = xmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_SUBCLASS");

        if (string.Compare(subClassName, "v1-Line", true, System.Globalization.CultureInfo.InvariantCulture) == 0)
        {
          ExportLineAsPointCollection(annotation, xmlAnnotation);
          return;
        }

        if (string.Compare(subClassName, "Pen", true, System.Globalization.CultureInfo.InvariantCulture) == 0)
        {
          ExportPenAsPointCollection(annotation, xmlAnnotation);
          return;
        }

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private static void ExportLineAsPointCollection(PointCollectionAnnotation annotation, CtsXmlDocument xmlAnnotation)
    {
      try
      {

        int lineWeight = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_WIDTH");
        Single x1 = xmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_START_X");
        Single y1 = xmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_START_Y");
        Single x2 = xmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_END_X");
        Single y2 = xmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_END_Y");
        PointStyle pointStyle = new PointStyle() { Endpoint = PointStyle.EndpointStyle.None, Filled = false, Thickness = lineWeight };

        annotation.SetStartPoint(x1, y1, pointStyle);
        annotation.AddSegment(x2, y2, pointStyle);
        annotation.Display.Border.LineStyle.LineWeight = lineWeight;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private static void ExportPenAsPointCollection(PointCollectionAnnotation annotation, CtsXmlDocument xmlAnnotation)
    {
      try
      {

        //  To be developed
        int lineWeight = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_WIDTH");
        Single x1 = xmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_START_X");
        Single y1 = xmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_START_Y");
        Single x2 = xmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_END_X");
        Single y2 = xmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_END_Y");
        PointStyle pointStyle = new PointStyle() { Endpoint = PointStyle.EndpointStyle.None, Filled = false, Thickness = lineWeight };

        annotation.SetStartPoint(x1, y1, pointStyle);
        annotation.AddSegment(x2, y2, pointStyle);
        annotation.Display.Border.LineStyle.LineWeight = lineWeight;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    #endregion

    #region Annotation Metadata Helpers

    private void ExportBorderInfo(Annotation annotation, CtsXmlDocument xmlAnnotation)
    {
      try
      {
        if (annotation == null) throw new ArgumentNullException(nameof(annotation));
        if (xmlAnnotation == null) throw new ArgumentNullException(nameof(xmlAnnotation));

        bool hasBorder = xmlAnnotation.QuerySingleAttributeAsBoolean(PROP_DESC_PATH, "F_HASBORDER");
        if (hasBorder)
        {
          annotation.Display.Border = null;
          return;
        }

        annotation.Display.Border = new BorderInfo();
        annotation.Display.Background = ParseColor(xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_BACKCOLOR"));
        annotation.Display.Border.Color = ParseColor(xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_BORDER_COLOR"));
        annotation.Display.Border.Color.Opacity = 50 * xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_BORDER_BACKMODE");

        annotation.Display.Border.LineStyle = new LineStyleInfo() { LineWeight = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_BORDER_WIDTH") };

        //int lineBorderStyle = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_BORDER_STYLE");
        //LineStyleInfo.LinePattern pattern = LineStyleInfo.LinePattern.None;

        //switch (lineBorderStyle)
        //{
        //  case 0: { pattern = LineStyleInfo.LinePattern.Solid; break; }
        //  case 1: { pattern = LineStyleInfo.LinePattern.Dash; break; }
        //  case 2: { pattern = LineStyleInfo.LinePattern.Dot; break; }
        //  case 3: { pattern = LineStyleInfo.LinePattern.DashDot; break; }
        //  case 4: { pattern = LineStyleInfo.LinePattern.DashDotDot; break; }
        //  default: { ApplicationLogging.WriteLogEntry($"Unknown border linestyle value {lineBorderStyle}");  break; }
        //}

        //annotation.Display.Border.LineStyle.Pattern = pattern;
        annotation.Display.Border.LineStyle.Pattern = GetLinePattern(xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_BORDER_STYLE"));

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private LineStyleInfo.LinePattern GetLinePattern(int patternCode)
    {
      try
      {
        LineStyleInfo.LinePattern pattern = LineStyleInfo.LinePattern.None;

        switch (patternCode)
        {
          case 0: { pattern = LineStyleInfo.LinePattern.Solid; break; }
          case 1: { pattern = LineStyleInfo.LinePattern.Dash; break; }
          case 2: { pattern = LineStyleInfo.LinePattern.Dot; break; }
          case 3: { pattern = LineStyleInfo.LinePattern.DashDot; break; }
          case 4: { pattern = LineStyleInfo.LinePattern.DashDotDot; break; }
          default: { ApplicationLogging.WriteLogEntry($"Unknown border linestyle value {patternCode}"); break; }
        }
        return pattern;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private void ExportFontMetadata(TextMarkup textMarkup, CtsXmlDocument xmlAnnotation)
    {
      try
      {
        if (textMarkup == null) throw new ArgumentNullException(nameof(textMarkup));
        if (xmlAnnotation == null) throw new ArgumentNullException(nameof(xmlAnnotation));

        textMarkup.Font = new FontInfo();
        textMarkup.Font.IsBold = xmlAnnotation.QuerySingleAttributeAsBoolean(PROP_DESC_PATH, "F_FONT_BOLD");
        textMarkup.Font.IsItalic = xmlAnnotation.QuerySingleAttributeAsBoolean(PROP_DESC_PATH, "F_FONT_ITALIC");
        textMarkup.Font.IsStrikethrough = xmlAnnotation.QuerySingleAttributeAsBoolean(PROP_DESC_PATH, "F_FONT_STRIKETHROUGH");
        textMarkup.Font.IsUnderline = xmlAnnotation.QuerySingleAttributeAsBoolean(PROP_DESC_PATH, "F_FONT_UNDERLINE");
        textMarkup.Font.FontName = xmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_FONT_NAME");
        textMarkup.Font.FontSize = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_FONT_SIZE");
        textMarkup.Font.Color = ParseColor(xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_FORECOLOR"));

        //  Preserve the font family, in case the target platform does not know the font by the exact name specified.
        //  If the font size is 0.0, then we can't initialize the platform's font info for font family.
        //  Check for fonts less than 0.5, because 0.0 may not test true.
        Single safeFontSize;
        if (textMarkup.Font.FontSize < 0.5)
        {
          safeFontSize = 12.0f;
        }
        else
        {
          safeFontSize = textMarkup.Font.FontSize;
        }

        System.Drawing.Font platformFont = new System.Drawing.Font(textMarkup.Font.FontName, safeFontSize);
        textMarkup.Font.FontFamily = platformFont.FontFamily.Name;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private void ExportLineMetadata(ArrowAnnotation annotation, CtsXmlDocument xmlAnnotation)
    {
      try
      {
        if (annotation == null) throw new ArgumentNullException(nameof(annotation));
        if (xmlAnnotation == null) throw new ArgumentNullException(nameof(xmlAnnotation));

        annotation.Display.Foreground = ParseColor(xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_COLOR"));

        if (xmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_LINE_BACK_MODE") == null)
        {
          annotation.Display.Foreground.Opacity = 100;
        }
        else
        {
          int lineBackMode = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_BACK_MODE");
          annotation.Display.Foreground.Opacity = lineBackMode;
        }

        annotation.LineStyle = new LineStyleInfo();
        annotation.LineStyle.Pattern = GetLinePattern(xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_STYLE"));
        annotation.LineStyle.LineWeight = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_WIDTH");

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private static void ExportTextBackgroundMode(Annotation annotation, CtsXmlDocument xmlAnnotation)
    {
      try
      {
        if (annotation == null) throw new ArgumentNullException(nameof(annotation));
        if (xmlAnnotation == null) throw new ArgumentNullException(nameof(xmlAnnotation));

        if (xmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_BACKCOLOR") != null)
        {
          int backgroundColor = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_BACKCOLOR");
          annotation.Display.Background = ParseColor(backgroundColor);
          annotation.Display.Background.Opacity = 100;

          if (xmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_TEXT_BACKMODE") != null)
          {
            int backgroundMode = xmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_TEXT_BACKMODE");
            annotation.Display.Background.Opacity = 50 * backgroundMode;
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

    private static void ExportTextMetadata(TextMarkup textMarkup, CtsXmlDocument xmlAnnotation)
    {
      try
      {
        if (textMarkup == null) throw new ArgumentNullException(nameof(textMarkup));
        if (xmlAnnotation == null) throw new ArgumentNullException(nameof(xmlAnnotation));

        string hexString = xmlAnnotation.QuerySingleString(TEXT_PATH);
        if (string.IsNullOrEmpty(hexString)) { return; }

        textMarkup.Text = DecodeUnicodeHexString(hexString);

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    private static ColorInfo ParseColor(int rgbColor)
    {
      ColorInfo result;
      try
      {
        //  This isn't accurate
        result = new ColorInfo() { Red = rgbColor & 255, Green = (rgbColor >> 8) & 255, Blue = (rgbColor >> 16) & 255, };
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
      return result;
    }

    private static string DecodeUnicodeHexString(string source)
    {
      string result = string.Empty;
      try
      {
        if (string.IsNullOrEmpty(source)) throw new ArgumentNullException(nameof(source));

        if (source.Length % 4 != 0) throw new ArgumentException("The Unicode hexadecimal string length must be evenly divisible by 4", nameof(source));

        StringBuilder builder = new StringBuilder();

        for (int position = 0; position < source.Length; position += 4)
        {
          int charValue = int.Parse(source.Substring(position, 4), System.Globalization.NumberStyles.HexNumber);
          builder.Append(charValue);
        }
        result = builder.ToString();
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
      return result;
    }

    #endregion

    #endregion

  }
}
