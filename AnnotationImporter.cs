using Documents.Annotations;
using Documents.Annotations.Common;
using Documents.Annotations.Decoration;
using Documents.Annotations.Exception;
using Documents.Annotations.Highlight;
using Documents.Annotations.Shape;
using Documents.Annotations.Special;
using Documents.Annotations.Text;
using Documents.Utilities;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace Documents.Providers.FileNetCEWS
{

  /// <summary>
  /// Builds an XML document / string suitable for Daeja ViewOne, based on the CTS Annotation model.
  /// </summary>
  public class AnnotationImporter
  {
    private const double NINETYSIX = 96.0;

    #region Class Constants

    public Single ScaleX = (float)NINETYSIX;
    public Single ScaleY = (float)NINETYSIX;

    #endregion

    #region IAnnotationImporter

    public void WriteAnnotationContent(StreamWriter stream, Annotation annotation, int contentElementIndex = 1)
    {
      try
      {
        //  Import annotation

        //  <FnAnno>
        XmlDocument xmlDoc = new XmlDocument();
        XmlElement fnAnno = xmlDoc.CreateElement("FnAnno");
        xmlDoc.AppendChild(fnAnno);

        //  <PropDesc>
        XmlElement propDesc = xmlDoc.CreateElement("PropDesc");
        ImportCommonMetadata(xmlDoc, propDesc, annotation, contentElementIndex);

        bool processed = false;

        if (annotation.GetType() == typeof(TextAnnotation))
        {
          Import(xmlDoc, propDesc, (TextAnnotation)annotation);
          processed = true;
        }

        if (annotation.GetType() == typeof(HighlightRectangle))
        {
          Import(xmlDoc, propDesc, (HighlightRectangle)annotation);
          processed = true;
        }

        if (annotation.GetType() == typeof(ArrowAnnotation))
        {
          Import(xmlDoc, propDesc, (ArrowAnnotation)annotation);
          processed = true;
        }

        if (annotation.GetType() == typeof(StickyNoteAnnotation))
        {
          Import(xmlDoc, propDesc, (StickyNoteAnnotation)annotation);
          processed = true;
        }

        if (annotation.GetType() == typeof(RectangleAnnotation))
        {
          Import(xmlDoc, propDesc, (RectangleAnnotation)annotation);
          processed = true;
        }

        if (annotation.GetType() == typeof(StampAnnotation))
        {
          Import(xmlDoc, propDesc, (StampAnnotation)annotation);
          processed = true;
        }

        if (!processed) { throw new UnsupportedAnnotationException("The annotation was not recognized during import."); }

        //  Load into annotation content element
        string annotationXml = xmlDoc.InnerXml;
        stream.Write(annotationXml);

        //  Never close the stream that we're given, just flush the data through.
        stream.Flush();

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    #endregion

    #region Annotation Import Methods

    private void ImportCommonMetadata(XmlDocument doc, XmlElement element, Annotation annotation, int contentElementIndex)
    {
      try
      {

        //  F_ANNOTATEDID = "{28EADBFC-CACE-4882-B9B4-9068B2D894EB}" 
        //  Daeja stores the id of the annotation here, not the id of the document or of the content element.
        element.Attributes.Append(CreateAttribute(doc, "F_ANNOTATEDID", annotation.ID));

        //  F_ENTRYDATE = "2010-07-01T21:20:36.0000000-05:00"
        element.Attributes.Append(CreateAttribute(doc, "F_ENTRYDATE", annotation.AuditEvents.Created.EventTime));

        //  F_HEIGHT = "0.3333333333333333"
        element.Attributes.Append(CreateAttribute(doc, "F_HEIGHT", (annotation.Layout.LowerRightExtent.Second / ScaleY) - (annotation.Layout.UpperLeftExtent.Second / ScaleY)));

        //  F_ID = "{28EADBFC-CACE-4882-B9B4-9068B2D894EB}"  (Non-portable, must be set to the id in the target system.)
        element.Attributes.Append(CreateAttribute(doc, "F_ID", annotation.ID));

        //  F_LEFT = "0.26666666666666666"
        element.Attributes.Append(CreateAttribute(doc, "F_LEFT", annotation.Layout.UpperLeftExtent.First / ScaleX));

        //  F_MODIFYDATE = "2010-07-01T21:20:50.0000000-05:00"
        element.Attributes.Append(CreateAttribute(doc, "F_MODIFYDATE", annotation.AuditEvents.Modified.EventTime));

        //  TODO: Test content MIME type for 	image/tiff or image/x-tiff
        if (true)
        {
          element.Attributes.Append(CreateAttribute(doc, "F_PAGENUMBER", annotation.Layout.PageNumber));
          element.Attributes.Append(CreateAttribute(doc, "F_MULTIPAGETIFFPAGENUMBER", 0));
        }
        else
        {
          //element.Attributes.Append(CreateAttribute(doc, "F_MULTIPAGETIFFPAGENUMBER", annotation.Layout.PageNumber));
        }

        //  F_NAME = "-1-1"
        element.Attributes.Append(CreateAttribute(doc, "F_NAME", $"-{contentElementIndex}-{annotation.ID}"));

        //  F_TOP = "0.3"
        element.Attributes.Append(CreateAttribute(doc, "F_TOP", annotation.Layout.UpperLeftExtent.Second / ScaleY));

        //  F_WIDTH = "0.5833333333333334" >
        element.Attributes.Append(CreateAttribute(doc, "F_WIDTH", (annotation.Layout.LowerRightExtent.First / ScaleX) - (annotation.Layout.UpperLeftExtent.First / ScaleX)));

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    /// <summary>
    /// Imports the highlight rectangle.
    /// </summary>
    /// <param name="doc"></param>
    /// <param name="element"></param>
    /// <param name="annotation"></param>
    private static void Import(XmlDocument doc, XmlElement element, HighlightRectangle annotation)
    {
      try
      {
        //  <PropDesc
        //  F_CLASSNAME = "Highlight"
        element.Attributes.Append(CreateAttribute(doc, "F_CLASSNAME", "Highlight"));

        //  F_CLASSID = "{5CF11942-018F-11D0-A87A-00A0246922A5}"
        element.Attributes.Append(CreateAttribute(doc, "F_CLASSID", "{5CF11942-018F-11D0-A87A-00A0246922A5}"));

        //  F_TEXT_BACKMODE = "1"
        //  F_TEXT_BACKMODE - P8 really sets this, although it doesn't make sense why.
        ImportTextBackgroundMode(doc, element, annotation);

        if (annotation.Display.Border != null)
        { 
          //  F_LINE_COLOR = "65535"
          element.Attributes.Append(CreateAttribute(doc, "F_LINE_COLOR", annotation.Display.Border.Color));

          //  F_LINE_WIDTH = "8"
          element.Attributes.Append(CreateAttribute(doc, "F_LINE_WIDTH", annotation.Display.Border.LineStyle.LineWeight));
        }

        //  F_BRUSHCOLOR = "39423"
        element.Attributes.Append(CreateAttribute(doc, "F_BRUSHCOLOR", annotation.HighlightColor));

        //  <F_CUSTOM_BYTES/>
        doc.LastChild.AppendChild(element);
        ImportCustomBytes(doc, annotation);

        //  <F_POINTS/>
        ImportPoints(doc, annotation);

        //  <F_TEXT/>
        ImportText(doc, annotation);

        //  </PropDesc> implied

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    private static void Import(XmlDocument doc, XmlElement element, TextAnnotation annotation)
    {
      try
      {
        //XmlAttribute attrib;

        //  F_CLASSNAME = "Text"
        element.Attributes.Append(CreateAttribute(doc, "F_CLASSNAME", "Text"));

        //  F_CLASSID = "{5CF11941-018F-11D0-A87A-00A0246922A5}"
        element.Attributes.Append(CreateAttribute(doc, "F_CLASSID", "{5CF11941-018F-11D0-A87A-00A0246922A5}"));

        // F_HASBORDER = "true" 
        // F_BACKCOLOR = "34026" 
        // F_BORDER_BACKMODE = "2" 
        // F_BORDER_COLOR = "65280" 
        // F_BORDER_STYLE = "0" 
        // F_BORDER_WIDTH = "1"
        ImportBorderInfo(doc, element, annotation);

        // F_FONT_BOLD = "true"
        // F_FONT_ITALIC = "false"
        // F_FONT_NAME = "arial"
        // F_FONT_SIZE = "12"
        // F_FONT_STRIKETHROUGH = "false"
        // F_FONT_UNDERLINE = "false"
        // F_FORECOLOR = "65280"
        ImportFontMetadata(doc, element, annotation.TextMarkups[0]);

        ImportTextBackgroundMode(doc, element, annotation);

        //  Commit <PropDesc> element
        doc.LastChild.AppendChild(element);

        //  <F_CUSTOM_BYTES/>
        ImportCustomBytes(doc, annotation);

        //  <F_POINTS/>
        ImportPoints(doc, annotation);

        //  <F_TEXT Encoding="unicode">0054006500730074000A0054006500780074</F_TEXT>
        ImportText(doc, annotation);

        //  </PropDesc> implied

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    //  Arrow
    private void Import(XmlDocument doc, XmlElement element, ArrowAnnotation annotation)
    {
      try
      {
        //  <PropDesc
        //  F_LINE_BACKMODE="2"
        //  F_LINE_COLOR="16711680"
        //  F_LINE_STYLE="0"
        //  F_LINE_WIDTH="12"
        ImportLineMetadata(doc, element, annotation);

        //  F_CLASSNAME="Arrow"
        element.Attributes.Append(CreateAttribute(doc, "F_CLASSNAME", "Arrow"));

        //  F_CLASSID="{5CF11946-018F-11D0-A87A-00A0246922A5}"
        element.Attributes.Append(CreateAttribute(doc, "F_CLASSID", "{5CF11946-018F-11D0-A87A-00A0246922A5}"));

        //  F_ARROWHEAD_SIZE="1" 
        //  Size is a non-portable enumeration (1,2,3) This is fine for FileNet/Daeja systems but may require a formal enumeration (small, medium, large) later on.
        element.Attributes.Append(CreateAttribute(doc, "F_ARROWHEAD_SIZE", annotation.Size));

        //  F_LINE_START_X="1.7966666666666666"
        element.Attributes.Append(CreateAttribute(doc, "F_LINE_START_X", annotation.StartPoint.First / ScaleX));

        //  F_LINE_START_Y="0.8333333333333334"
        element.Attributes.Append(CreateAttribute(doc, "F_LINE_START_Y", annotation.StartPoint.Second / ScaleY));

        //  F_LINE_END_X="1.28"
        element.Attributes.Append(CreateAttribute(doc, "F_LINE_END_X", annotation.EndPoint.First / ScaleX));

        //  F_LINE_END_Y="0.7366666666666667"
        element.Attributes.Append(CreateAttribute(doc, "F_LINE_END_Y", annotation.EndPoint.Second / ScaleY));

        //  <F_CUSTOM_BYTES/>
        doc.LastChild.AppendChild(element);
        ImportCustomBytes(doc, annotation);

        //  <F_POINTS/>
        ImportPoints(doc, annotation);

        //  <F_TEXT/>
        ImportText(doc, annotation);

        //  </PropDesc> implied

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    //  Stickynote
    private static void Import(XmlDocument doc, XmlElement element, StickyNoteAnnotation annotation)
    {
      try
      {
        //  <PropDesc
        //  F_CLASSNAME = "StickyNote"
        element.Attributes.Append(CreateAttribute(doc, "F_CLASSNAME", "StickyNote"));

        //  F_CLASSID = "{5CF11945-018F-11D0-A87A-00A0246922A5}"
        element.Attributes.Append(CreateAttribute(doc, "F_CLASSID", "{5CF11945-018F-11D0-A87A-00A0246922A5}"));

        //  F_FORECOLOR = "10092543"
        //  This value is fixed for now, as CS does not export it.
        element.Attributes.Append(CreateAttribute(doc, "F_FORECOLOR", FormatColor(new ColorInfo() { Red = 153, Green = 255, Blue = 255, Opacity = 100 })));

        //  F_ORDINAL
        element.Attributes.Append(CreateAttribute(doc, "F_ORDINAL", annotation.NoteOrder));

        //	Commit <PropDesc> element
        doc.LastChild.AppendChild(element);

        //  <F_CUSTOM_BYTES/>
        ImportCustomBytes(doc, annotation);

        //  <F_POINTS/>
        ImportPoints(doc, annotation);

        //  <F_TEXT Encoding="unicode">0053007400690063006B00790020006E006F00740065</F_TEXT>
        ImportText(doc, annotation);

        //  </PropDesc> implied

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    //  Pen
    private static void Import(XmlDocument doc, XmlElement element, PointCollectionAnnotation annotation)
    {
      try
      {
        //  <PropDesc
        //  F_ANNOTATEDID="{F284249D-D541-43D6-A1AD-7D5A4286C876}" F_CLASSID="{5CF11949-018F-11D0-A87A-00A0246922A5}" F_CLASSNAME="Pen" F_ENTRYDATE="2010-07-01T21:34:56.0000000-05:00" F_HEIGHT="0.6033333333333334" F_ID="{F284249D-D541-43D6-A1AD-7D5A4286C876}" F_LEFT="0.16333333333333333" F_LINE_BACKMODE="2" F_LINE_COLOR="65280" F_LINE_STYLE="0" F_LINE_WIDTH="12" F_MODIFYDATE="2010-07-01T21:35:14.0000000-05:00" F_MULTIPAGETIFFPAGENUMBER="0" F_NAME="-1-1" F_PAGENUMBER="1" F_TOP="0.16666666666666666" F_WIDTH="0.23">

        //  Commit <PropDesc> element
        doc.LastChild.AppendChild(element);

        //  <F_CUSTOM_BYTES/>
        ImportCustomBytes(doc, annotation);

        //  <F_POINTS>0 0 0 0 7 4 11 8 11 13 11 20 15 25 19 31 26 37 30 42 33 47 37 51 45 55 48 59 52 64 63 68 67 72 85 79 93 83 108 88 119 91 134 95 145 98 160 100 171 103 182 105 193 106 200 110 204 115 204 120 208 126 208 130 211 136 211 140 215 144 219 149 219 153 219 157 219 161 219 165 219 170 215 174 215 178 215 182 215 187 215 191 211 195 211 199 211 205 211 209 215 214 215 218 219 222 223 229 226 235 230 239 241 246 245 250 255 255</F_POINTS>
        ImportPoints(doc, annotation);

        //  <F_TEXT/>
        ImportText(doc, annotation);

        //  </PropDesc> implied

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    //  Stamp
    private static void Import(XmlDocument doc, XmlElement element, StampAnnotation annotation)
    {
      try
      {
        //  <!--  Stamp annotation
        //  Yellow box, red border, reading "Urgent!" with dots missing from text
        //  text is at a diagonal from lower-left to upper right -->
        //  <PropDesc F_ANNOTATEDID="{9E8F50B6-CE56-4E9C-A9F4-6E637D8BFA3E}" F_BACKCOLOR="65535" F_BORDER_BACKMODE="2" F_BORDER_COLOR="255" F_BORDER_STYLE="0" F_BORDER_WIDTH="1" F_CLASSID="{5CF1194C-018F-11D0-A87A-00A0246922A5}" F_CLASSNAME="Stamp" F_ENTRYDATE="2010-07-01T21:38:34.0000000-05:00" F_FONT_BOLD="true" F_FONT_ITALIC="false" F_FONT_NAME="arial" F_FONT_SIZE="22" F_FONT_STRIKETHROUGH="false" F_FONT_UNDERLINE="false" F_FORECOLOR="255" F_HASBORDER="true" F_HEIGHT="1.37" F_ID="{9E8F50B6-CE56-4E9C-A9F4-6E637D8BFA3E}" F_LEFT="0.3433333333333333" F_MODIFYDATE="2010-07-01T21:39:03.0000000-05:00" F_MULTIPAGETIFFPAGENUMBER="0" F_NAME="-1-1" F_PAGENUMBER="1" F_ROTATION="45" F_TEXT_BACKMODE="2" F_TOP="1.19" F_WIDTH="1.5033333333333334">

        //  <PropDesc
        //  F_CLASSNAME = "Stamp"
        element.Attributes.Append(CreateAttribute(doc, "F_CLASSNAME", "Stamp"));

        //  F_CLASSID = "{5CF11942-018F-11D0-A87A-00A0246922A5}"
        element.Attributes.Append(CreateAttribute(doc, "F_CLASSID", "{5CF1194C-018F-11D0-A87A-00A0246922A5}"));

        ImportBorderInfo(doc, element, annotation);
        ImportFontMetadata(doc, element, annotation.TextElement);

        //  F_BACKCOLOR - should already be set
        //  node.Attributes.Append(Me.CreateAttribute(doc, "F_BACKCOLOR", ctsAnnotation.Display.Background))

        //  F_FORECOLOR = "255" - should already be set

        //  F_ROTATION = "45"
        element.Attributes.Append(CreateAttribute(doc, "F_ROTATION", (float)(360.0 - annotation.TextElement.TextRotation)));

        //  F_TEXT_BACKMODE = "2" - should already be set
        ImportTextBackgroundMode(doc, element, annotation);

        //  Commit<PropDesc> element
        doc.LastChild.AppendChild(element);

        //  < F_CUSTOM_BYTES />
        ImportCustomBytes(doc, annotation);

        //  < F_POINTS />
        ImportPoints(doc, annotation);

        //  <F_TEXT Encoding="unicode">0055007200670065006E00740021</F_TEXT>
        ImportText(doc, annotation);

        //  </PropDesc> implied

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }


    //  Rectangle
    private static void Import(XmlDocument doc, XmlElement element, RectangleAnnotation annotation)
    {
      try
      {
        int fillMode = 0;

        element.Attributes.Append(CreateAttribute(doc, "F_CLASSNAME", "Proprietary"));
        element.Attributes.Append(CreateAttribute(doc, "F_SUBCLASS", "v1-Rectangle"));
        element.Attributes.Append(CreateAttribute(doc, "F_CLASSID", "{A91E5DF2-6B7B-11D1-B6D7-00609705F027}"));
        element.Attributes.Append(CreateAttribute(doc, "F_BRUSHCOLOR", annotation.Display.Foreground));
        element.Attributes.Append(CreateAttribute(doc, "F_LINE_COLOR", annotation.Display.Border.Color));
        element.Attributes.Append(CreateAttribute(doc, "F_LINE_WIDTH", annotation.LineStyle.LineWeight));

        if (annotation.Display.Foreground.Opacity > 25) { fillMode = 1; }

        if ((annotation.IsFilled) || (annotation.Display.Foreground.Opacity > 75)) { fillMode = 2; }
        element.Attributes.Append(CreateAttribute(doc, "F_TEXT_BACKMODE", fillMode));

        doc.LastChild.AppendChild(element);
        ImportCustomBytes(doc, annotation);
        ImportPoints(doc, annotation);
        ImportText(doc, annotation);

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    #region Annotation Metadata Helpers

    /// <summary>
    /// Imports the font metadata.
    /// </summary>
    /// <param name="doc">The doc.</param>
    /// <param name="node">The node.</param>
    /// <param name="textMarkupInfo">The text markup info.</param>
    private static void ImportFontMetadata(XmlDocument doc, XmlElement node, TextMarkup textMarkupInfo)
    {
      try
      {
        //  P8 provides support for only one text markup within an annotation.

        //  F_FONT_BOLD = "true"
        node.Attributes.Append(CreateAttribute(doc, "F_FONT_BOLD", textMarkupInfo.Font.IsBold));

        //  F_FONT_ITALIC = "false"
        node.Attributes.Append(CreateAttribute(doc, "F_FONT_ITALIC", textMarkupInfo.Font.IsItalic));

        //  F_FONT_NAME = "arial"
        //  Ideally, we should do a check on the local platform to see if the font is installed and if not, select another based on the FontFamily property.
        //  The operation is expensive, so the font lookup wrapper should be memorized.
        node.Attributes.Append(CreateAttribute(doc, "F_FONT_NAME", textMarkupInfo.Font.FontName.ToLowerInvariant()));

        //  F_FONT_SIZE = "12"
        node.Attributes.Append(CreateAttribute(doc, "F_FONT_SIZE", textMarkupInfo.Font.FontSize));

        //  F_FONT_STRIKETHROUGH = "false"
        node.Attributes.Append(CreateAttribute(doc, "F_FONT_STRIKETHROUGH", textMarkupInfo.Font.IsStrikethrough));

        //  F_FONT_UNDERLINE = "false"
        node.Attributes.Append(CreateAttribute(doc, "F_FONT_UNDERLINE", textMarkupInfo.Font.IsUnderline));

        //  F_FORECOLOR = "65280"
        node.Attributes.Append(CreateAttribute(doc, "F_FORECOLOR", FormatColor(textMarkupInfo.Font.Color)));

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    /// <summary>
    /// Imports the border info.
    /// Only valid for Stamp and Text annotations on this platform.
    /// </summary>
    /// <param name="doc">The doc.</param>
    /// <param name="node">The node.</param>
    /// <param name="annotation">The CTS annotation.</param>
    private static void ImportBorderInfo(XmlDocument doc, XmlElement node, Annotation annotation)
    {
      try
      {
        if (annotation == null) { throw new ArgumentNullException(nameof(annotation)); }

        //  F_HASBORDER 
        if (annotation.Display.Border == null) 
        { 
          node.Attributes.Append(CreateAttribute(doc, "F_HASBORDER", false));
          return;
        }
        else
        {
          node.Attributes.Append(CreateAttribute(doc, "F_HASBORDER", true));
        }

        //  F_BACKCOLOR
        node.Attributes.Append(CreateAttribute(doc, "F_BACKCOLOR", FormatColor(annotation.Display.Background)));

        //  F_BORDER_BACKMODE
        int borderBackMode = 0;
        if (annotation.Display.Border.Color.Opacity > 25) { borderBackMode = 1; }
        if (annotation.Display.Border.Color.Opacity > 75) { borderBackMode = 2; }

        node.Attributes.Append(CreateAttribute(doc, "F_BORDER_BACKMODE", borderBackMode));

        //  F_BORDER_COLOR
        node.Attributes.Append(CreateAttribute(doc, "F_BORDER_COLOR", FormatColor(annotation.Display.Border.Color)));

        //  F_BORDER_WIDTH
        node.Attributes.Append(CreateAttribute(doc, "F_BORDER_WIDTH", annotation.Display.Border.LineStyle.LineWeight));

        //  F_BORDER_STYLE
        int borderStyle = 0;
        switch (annotation.Display.Border.LineStyle.Pattern)
        {
          case LineStyleInfo.LinePattern.Solid:
            {
              borderStyle = 1;
              break;
            }

          case LineStyleInfo.LinePattern.Dash:
            {
              borderStyle = 2;
              break;
            }

          case LineStyleInfo.LinePattern.Dot:
            {
              borderStyle = 3;
              break;
            }

          case LineStyleInfo.LinePattern.DashDot:
            {
              borderStyle = 4;
              break;
            }

          case LineStyleInfo.LinePattern.DashDotDot:
            {
              borderStyle = 5;
              break;
            }

          default:
            //  TODO: Throw exception or favor robustness?
            break;
        }

        node.Attributes.Append(CreateAttribute(doc, "F_BORDER_STYLE", borderStyle));

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    /// <summary>
    /// Imports the text background mode.
    /// </summary>
    /// <param name="doc">The doc.</param>
    /// <param name="node">The node.</param>
    /// <param name="annotation">The CTS annotation.</param>
    private static void ImportTextBackgroundMode(XmlDocument doc, XmlElement node, Annotation annotation)
    {
      try
      {
        if (annotation == null) { throw new ArgumentNullException(nameof(annotation)); }

        if (annotation.Display.Background == null) { return; }

        //  F_BACKCOLOR
        node.Attributes.Append(CreateAttribute(doc, "F_BACKCOLOR", FormatColor(annotation.Display.Background)));

        //  F_TEXT_BACKMODE = "2"
        if (annotation.Display.Background.Opacity <= 50)
        {
          node.Attributes.Append(CreateAttribute(doc, "F_TEXT_BACKMODE", 1));
        }
        else
        {
          node.Attributes.Append(CreateAttribute(doc, "F_TEXT_BACKMODE", 2));
        }

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    /// <summary>
    /// Imports the line metadata.
    /// </summary>
    /// <param name="doc">The doc.</param>
    /// <param name="node">The node.</param>
    /// <param name="annotation">The CTS annotation.</param>
    private static void ImportLineMetadata(XmlDocument doc, XmlElement node, LineBase annotation)
    {
      try
      {

        if (annotation == null) { throw new ArgumentNullException(nameof(annotation)); }

        //  F_LINE_BACKMODE="2"
        if (annotation.Display.Foreground.Opacity <= 50)
        {
          node.Attributes.Append(CreateAttribute(doc, "F_LINE_BACKMODE", 1));
        }
        else
        {
          node.Attributes.Append(CreateAttribute(doc, "F_LINE_BACKMODE", 2));
        }

        //  F_LINE_COLOR="16711680"
        node.Attributes.Append(CreateAttribute(doc, "F_LINE_COLOR", annotation.Display.Foreground));

        //  F_LINE_STYLE="0"
        int lineStyle = 0;
        switch (annotation.LineStyle.Pattern)
        {
          case LineStyleInfo.LinePattern.Solid:
            {
              lineStyle = 0;
              break;
            }

          case LineStyleInfo.LinePattern.Dash:
            {
              lineStyle = 1;
              break;
            }

          case LineStyleInfo.LinePattern.Dot:
            {
              lineStyle = 2;
              break;
            }

          case LineStyleInfo.LinePattern.DashDot:
            {
              lineStyle = 3;
              break;
            }

          case LineStyleInfo.LinePattern.DashDotDot:
            {
              lineStyle = 4;
              break;
            }

          default:
            break;
        }
        node.Attributes.Append(CreateAttribute(doc, "F_LINE_STYLE", lineStyle));

        //  F_LINE_WIDTH="12"
        node.Attributes.Append(CreateAttribute(doc, "F_LINE_WIDTH", annotation.LineStyle.LineWeight));

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    #endregion

    #region XML builder helper methods

    /// <summary>
    /// Processes the "custom bytes" portion of the annotation.
    /// Every annotation in this system has at least an empty XML tag for this.
    /// </summary>
    /// <param name="doc">The doc.</param>
    /// <param name="annotation">The CTS annotation.</param>
    private static void ImportCustomBytes(XmlDocument doc, Annotation annotation)
    {
      try
      {
        doc.DocumentElement.LastChild.AppendChild(doc.CreateElement("F_CUSTOM_BYTES"));
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    /// <summary>
    /// Processes the "data points" portion of the annotation.
    /// Every annotation in this system has at least an empty XML tag for this.
    /// </summary>
    /// <param name="doc">The doc.</param>
    /// <param name="annotation">The CTS annotation.</param>
    private static void ImportPoints(XmlDocument doc, Annotation annotation)
    {
      try
      {
        doc.DocumentElement.LastChild.AppendChild(doc.CreateElement("F_POINTS"));
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }

    }

    /// <summary>
    /// Processes the "text string" portion of the annotation.
    /// Every annotation in this system has at least an empty XML tag for this.
    /// </summary>
    /// <param name="doc">The doc.</param>
    /// <param name="annotation">The CTS annotation.</param>
    private static void ImportText(XmlDocument doc, Annotation ctsAnnotation)
    {
      try
      {
        XmlElement element = doc.CreateElement("F_TEXT");
        string encodedOutput = null;

        if (ctsAnnotation.GetType() == typeof(TextAnnotation))
        {
          TextAnnotation annotation = (TextAnnotation)ctsAnnotation;

          //  For P8, only one text markup per annotation is possible at this time in Daeja
          encodedOutput = FormatUnicodeHexString(annotation.TextMarkups[0].Text);
        }

        if (ctsAnnotation.GetType() == typeof(StampAnnotation))
        {
          StampAnnotation annotation = (StampAnnotation)ctsAnnotation;
          encodedOutput = FormatUnicodeHexString(annotation.TextElement.Text);
        }

        if (ctsAnnotation.GetType() == typeof(StickyNoteAnnotation))
        {
          StickyNoteAnnotation annotation = (StickyNoteAnnotation)ctsAnnotation;
          encodedOutput = FormatUnicodeHexString(annotation.TextNote.Text);
        }

        XmlNode previousParent = doc.DocumentElement.LastChild;

        if (encodedOutput != null)
        {
          element.Attributes.Append(CreateAttribute(doc, "Encoding", "uinicode"));
          element.InnerText = encodedOutput;
        }

        previousParent.AppendChild(element);

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }

    }

    #endregion


    #endregion

    #region Data Exchange Helper Methods

    #region Overloads for CreateAttribute

    /// <summary>
    /// Creates an XML attribute having the name and value provided.
    /// </summary>
    /// <param name="doc">The doc</param>
    /// <param name="attributeName">Name of the attribute</param>
    /// <param name="attributeValue">The attribute value</param>
    /// <returns></returns>
    private static XmlAttribute CreateAttribute(XmlDocument doc, string attributeName, string attributeValue)
    {
      try
      {
        XmlAttribute attribute = doc.CreateAttribute(attributeName);
        attribute.Value = attributeValue;
        return attribute;
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }      
    }

    /// <summary>
    /// Creates an XML attribute having the name and value provided.
    /// </summary>
    /// <param name="doc">The doc</param>
    /// <param name="attributeName">Name of the attribute</param>
    /// <param name="attributeValue">The attribute value</param>
    /// <returns></returns>
    private static XmlAttribute CreateAttribute(XmlDocument doc, string attributeName, Single attributeValue)
    {
      try
      {
        return CreateAttribute(doc, attributeName, attributeValue.ToString());
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    private static XmlAttribute CreateAttribute(XmlDocument doc, string attributeName, DateTimeOffset attributeValue)
    {
      try
      {
        return CreateAttribute(doc, attributeName, FormatDate(attributeValue));
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    private static XmlAttribute CreateAttribute(XmlDocument doc, string attributeName, ColorInfo attributeValue)
    {
      try
      {
        return CreateAttribute(doc, attributeName, FormatColor(attributeValue));
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    private static XmlAttribute CreateAttribute(XmlDocument doc, string attributeName, bool attributeValue)
    {
      //  The only way Daeja P8 will recognize Boolean attributes is if they are in *lowercase*
      try
      {
        if (attributeValue)
        {
          return CreateAttribute(doc, attributeName, "true");
        }
        else
        {
          return CreateAttribute(doc, attributeName, "false");
        }
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    #endregion

    /// <summary>
    /// Formats the date.
    /// </summary>
    /// <param name="date">The CTS Date</param>
    /// <returns></returns>
    private static string FormatDate(DateTimeOffset date)
    {
      //  F_ENTRYDATE =  "2010-07-01T21:20:36.0000000-05:00"
      //  F_MODIFYDATE = "2010-07-01T21:20:50.0000000-05:00"
      try
      {
        return String.Format("{0:yyyy-MM-dd}T{0:HH:mm:ss.fffffffK}", date);
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    private static string FormatColor(ColorInfo color)
    {
      try
      {
        //  left as multiple statements for debugging.
        int result = color.Blue;
        result *= 256;
        result += color.Green;
        result *= 256;
        result += color.Red;
        return result.ToString();
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    private static string FormatUnicodeHexString(string source)
    {
      try
      {
        if (source == null) {  return string.Empty; }

        StringBuilder builder = new StringBuilder();
        for (int i = 0; i < source.Length; i++)
        {
          builder.AppendFormat("{0:X4}", Asc(source[i]));
        }

        return builder.ToString();

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    //
    // Summary:
    //     Returns an integer value representing the character code corresponding to a character.
    //
    //
    // Parameters:
    //   String:
    //     Required. Any valid Char or String expression. If String is a String expression,
    //     only the first character of the string is used for input. If String is Nothing
    //     or contains no characters, an System.ArgumentException error occurs.
    //
    // Returns:
    //     The character code corresponding to a character.
    public static int Asc(char String)
    {
      int num = Convert.ToInt32(String);
      if (num < 128)
      {
        return num;
      }

      try
      {
        Encoding fileIOEncoding = Encoding.Default;
        char[] chars = new char[1] { String };
        byte[] array;
        if (fileIOEncoding.IsSingleByte)
        {
          array = new byte[1];
          fileIOEncoding.GetBytes(chars, 0, 1, array, 0);
          return array[0];
        }

        array = new byte[2];
        if (fileIOEncoding.GetBytes(chars, 0, 1, array, 0) == 1)
        {
          return array[0];
        }

        if (BitConverter.IsLittleEndian)
        {
          byte b = array[0];
          array[0] = array[1];
          array[1] = b;
        }

        return BitConverter.ToInt16(array, 0);
      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re-throw the exception to the caller
        throw;
      }
    }

    #endregion


  }
}
