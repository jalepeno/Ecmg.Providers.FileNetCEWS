using CPEServiceReference;
using Documents.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Documents.Providers.FileNetCEWS
{
  public class Factory
  {

    #region Enumerations

    private enum SimpleVariableType
    {
      SingletonBinary,
      SingletonBoolean,
      SingletonDateTime,
      SingletonFloat64,
      SingletonID,
      SingletonInteger32,
      SingletonObject,
      SingletonString,
      ListOfBinary=10,
      ListOfBoolean,
      ListOfDateTime,
      ListOfFloat64,
      ListOfID,
      ListOfInteger32,
      ListOfObject,
      ListOfString
    }

    #endregion

    public static ObjectStoreScope ObjectStoreScope(string objectStoreName) { return new ObjectStoreScope() { objectStore = objectStoreName }; }

    public static CreateAction CreateAction(string classId) { return new CreateAction() { classId = classId }; }

    public static CheckoutAction CheckoutAction(ReservationType reservationType = ReservationType.Exclusive, bool reservationTypeSpecified = true)
      { return new CheckoutAction() { reservationType = reservationType, reservationTypeSpecified = reservationTypeSpecified }; } 

    public static CheckinAction CheckinAction(bool minorVersion = true, bool minorVersionSpecified = true) 
      { return new CheckinAction() { checkinMinorVersion = minorVersion, checkinMinorVersionSpecified = minorVersionSpecified }; }

    public static UpdateAction UpdateAction() { return new UpdateAction(); }

    public static ContentData ContentData(InlineContent content) { return new ContentData() { Value = content, propertyId = "Content" }; }

    public static DependentObjectType DependentObjectType(string classId, DependentObjectTypeDependentAction dependentAction = DependentObjectTypeDependentAction.Insert, bool dependentActionSpecified = true)
    { return new DependentObjectType() { classId = classId, dependentAction = dependentAction, dependentActionSpecified = dependentActionSpecified }; }

    public static DependentObjectType ContentTransfer() 
    {
      DependentObjectType contentTransfer = DependentObjectType("ContentTransfer");
      contentTransfer.Property = new PropertyType[2];
      return contentTransfer;
    }

    public static ExecuteChangesRequest ExecuteChangesRequest(ChangeRequestType changeRequestType, bool refresh = true, bool refreshSpecified = true)
    {
      ExecuteChangesRequest request = new ExecuteChangesRequest() { ChangeRequest = new ChangeRequestType[] { changeRequestType }, refreshSpecified = refreshSpecified };
      if (refreshSpecified) { request.refreshSpecified = true; }
      return request;
    }

    #region Singleton Variables

    public static SingletonBinary SingletonBinary(string propertyId, bool settable = true, bool settableSpecified = true, byte[] value = null) { return (SingletonBinary)SimpleVariable(SimpleVariableType.SingletonBinary, propertyId, settable, settableSpecified); }

    public static SingletonId SingletonId(string propertyId, string value = "", bool valueSpecified = false, bool settable = true, bool settableSpecified = true) 
    {
      SingletonId returnValue = (SingletonId)SimpleVariable(SimpleVariableType.SingletonID, propertyId, settable, settableSpecified); 
      if (valueSpecified ) { returnValue.Value = value; }
      return returnValue;
    }

    public static SingletonBoolean SingletonBoolean(string propertyId, bool value = false, bool valueSpecified = false, bool settable = true, bool settableSpecified = true)
    {
      SingletonBoolean returnValue = (SingletonBoolean)SimpleVariable(SimpleVariableType.SingletonBoolean, propertyId, settable, settableSpecified);
      if (valueSpecified) { returnValue.Value = value; returnValue.ValueSpecified = valueSpecified; }
      return returnValue;
    }

    public static SingletonDateTime SingletonDateTime(string propertyId, DateTime value, bool valueSpecified = false, bool settable = true, bool settableSpecified = true)
    {
      SingletonDateTime returnValue = (SingletonDateTime)SimpleVariable(SimpleVariableType.SingletonDateTime, propertyId, settable, settableSpecified);
      if (valueSpecified) 
      { 
        returnValue.Value = value; 
        returnValue.ValueSpecified = valueSpecified;
      }
      return returnValue;
    }

    public static SingletonFloat64 SingletonFloat64(string propertyId, double value = 0, bool valueSpecified = false, bool settable = true, bool settableSpecified = true)
    {
      SingletonFloat64 returnValue = (SingletonFloat64)SimpleVariable(SimpleVariableType.SingletonFloat64, propertyId, settable, settableSpecified);
      if (valueSpecified) { returnValue.Value = value; returnValue.ValueSpecified = valueSpecified; }
      return returnValue;
    }

    public static SingletonInteger32 SingletonInteger32(string propertyId, long value = 0, bool valueSpecified = false, bool settable = true, bool settableSpecified = true)
    {
      SingletonInteger32 returnValue = (SingletonInteger32)SimpleVariable(SimpleVariableType.SingletonInteger32, propertyId, settable, settableSpecified);
      if (valueSpecified) 
      { 
        returnValue.Value = (int)value; 
        returnValue.ValueSpecified = true;
      }
      return returnValue;
    }

    public static SingletonObject SingletonObject(string propertyId, Object value = null, bool valueSpecified = false, bool settable = true, bool settableSpecified = true)
    {
      SingletonObject returnValue = (SingletonObject)SimpleVariable(SimpleVariableType.SingletonObject, propertyId, settable, settableSpecified);
      if (valueSpecified) { returnValue.Value = (ObjectEntryType)value; }

      return returnValue;
    }

    public static SingletonObject SingletonObject(string propertyId, string objectStoreName, string value = "", bool valueSpecified = false, bool settable = true, bool settableSpecified = true)
    {
      SingletonObject returnValue = (SingletonObject)SimpleVariable(SimpleVariableType.SingletonObject, propertyId, settable, settableSpecified);
      if (valueSpecified)
      {
        ObjectValue objectValue = new ObjectValue { objectId = value, objectStore = objectStoreName };
        returnValue.Value = objectValue;
      }
      return returnValue;
    }

    public static SingletonString SingletonString(string propertyId, string value = "", bool valueSpecified = false, bool settable = true, bool settableSpecified = true)
    {
      SingletonString returnValue = (SingletonString)SimpleVariable(SimpleVariableType.SingletonString, propertyId, settable, settableSpecified);
      if (valueSpecified) { returnValue.Value = value; }
      return returnValue;
    }

    #endregion

    #region ListOf Variables

    public static ListOfBinary ListOfBinary(string propertyId, bool settable, bool settableSpecified) { return (ListOfBinary)SimpleVariable(SimpleVariableType.ListOfBinary, propertyId, settable, settableSpecified); }

    public static ListOfBoolean ListOfBoolean(string propertyId, bool settable, bool settableSpecified) { return (ListOfBoolean)SimpleVariable(SimpleVariableType.ListOfBoolean, propertyId, settable, settableSpecified); }

    public static ListOfDateTime ListOfDateTime(string propertyId, bool settable, bool settableSpecified) { return (ListOfDateTime)SimpleVariable(SimpleVariableType.ListOfDateTime, propertyId, settable, settableSpecified); }

    public static ListOfFloat64 ListOfFloat64(string propertyId, bool settable, bool settableSpecified) { return (ListOfFloat64)SimpleVariable(SimpleVariableType.ListOfFloat64, propertyId, settable, settableSpecified); }

    public static ListOfId ListOfId(string propertyId, bool settable, bool settableSpecified) { return (ListOfId)SimpleVariable(SimpleVariableType.ListOfID, propertyId, settable, settableSpecified); }

    public static ListOfInteger32 ListOfInteger32(string propertyId, bool settable, bool settableSpecified) { return (ListOfInteger32)SimpleVariable(SimpleVariableType.ListOfInteger32, propertyId, settable, settableSpecified); }

    public static ListOfObject ListOfObject(string propertyId, bool settable, bool settableSpecified) { return (ListOfObject)SimpleVariable(SimpleVariableType.ListOfObject, propertyId, settable, settableSpecified); }

    public static ListOfString ListOfString(string propertyId, bool settable, bool settableSpecified) { return (ListOfString)SimpleVariable(SimpleVariableType.ListOfString, propertyId, settable, settableSpecified); }

    #endregion

    public static ObjectReference ObjectReference(string classId, string objectId, string objectStore)
    {
      ObjectReference objectReference = new ObjectReference() { classId = classId };
      if (objectId.Length > 0) {  objectReference.objectId = objectId; }
      if (objectStore.Length > 0) { objectReference.objectStore = objectStore; }
      return objectReference;
    }

    public static FilterElementType FilterElementType(string name, int? maxRecursion = null)
    {
      try
      {
        FilterElementType filter = new FilterElementType();

        if (maxRecursion != null)
        {
          filter.maxRecursion = maxRecursion.Value;
          filter.maxRecursionSpecified = true;
        }
        filter.Value = name;

        return filter;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public static PropertyFilterType PropertyFilterType(int maxRecursion = 1, bool maxRecursionSpecified = true, FilterElementType[] includeProperties = null, string[] excludeProperties = null, int maxElements = 0, bool maxElementsSpecified = false)
    {
      try
      {
        //  Construct Property Filter
        PropertyFilterType propertyFilterType = new PropertyFilterType();

        if (maxRecursionSpecified)
        {
          propertyFilterType.maxRecursion = maxRecursion;
          propertyFilterType.maxRecursionSpecified = maxRecursionSpecified;
        }

        if (maxElementsSpecified)
        {
          propertyFilterType.maxElements = maxElements;
          propertyFilterType.maxElementsSpecified = maxElementsSpecified;
        }

        if (includeProperties != null)
        {
          propertyFilterType.IncludeProperties = includeProperties;
        }

        if (excludeProperties != null)
        {
          propertyFilterType.ExcludeProperties = excludeProperties;
        }

        return propertyFilterType;

      }
      catch (Exception ex)
      {
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod());
        //  Re - throw the exception to the caller
        throw;
      }
    }

    public static RepositorySearch RepositorySearch(string objectStoreName, string searchSql, RepositorySearchModeType repositorySearchMode, bool repositorySearchModeSpecified = true, int maxElements = 100, bool maxElementsSpecified = true)
    {
      return new RepositorySearch() 
        { 
        repositorySearchMode = repositorySearchMode, 
        repositorySearchModeSpecified = repositorySearchModeSpecified, 
        SearchScope = ObjectStoreScope(objectStoreName), 
        SearchSQL = searchSql, 
        maxElements = maxElements, 
        maxElementsSpecified = maxElementsSpecified 
      };
    }

    #region Private Methods

    private static Object SimpleVariable(SimpleVariableType variableType, string propertyId, bool settable = true, bool settableSpecified = true)
    {
      ModifiablePropertyType variable;

      switch (variableType)
      {
        case SimpleVariableType.SingletonBinary: { variable = new SingletonBinary(); break; }
        case SimpleVariableType.SingletonBoolean: { variable = new SingletonBoolean(); break; }
        case SimpleVariableType.SingletonDateTime: { variable = new SingletonDateTime(); break; }
        case SimpleVariableType.SingletonFloat64: { variable = new SingletonFloat64(); break; }
        case SimpleVariableType.SingletonID: { variable = new SingletonId(); break; }
        case SimpleVariableType.SingletonInteger32: { variable = new SingletonInteger32(); break; }
        case SimpleVariableType.SingletonObject: { variable = new SingletonObject(); break; }
        case SimpleVariableType.SingletonString: { variable = new SingletonString(); break; }
        case SimpleVariableType.ListOfBinary: { variable = new ListOfBinary(); break; }
        case SimpleVariableType.ListOfBoolean: { variable = new ListOfBoolean(); break; }
        case SimpleVariableType.ListOfDateTime: { variable = new ListOfDateTime(); break; }
        case SimpleVariableType.ListOfFloat64: { variable = new ListOfFloat64(); break; }
        case SimpleVariableType.ListOfID: { variable = new ListOfId(); break; }
        case SimpleVariableType.ListOfInteger32: { variable = new ListOfInteger32(); break; }
        case SimpleVariableType.ListOfObject: { variable = new ListOfObject(); break; }
        case SimpleVariableType.ListOfString: { variable = new ListOfString(); break; }

        default: { variable = new SingletonString(); break; }
      }

      variable.propertyId = propertyId;
      variable.settable = settable;
      variable.settableSpecified = settableSpecified;

      return variable;

    }

    #endregion

  }

}
