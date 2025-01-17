// ---------------------------------------------------------------------------------
// <copyright company="ECMG">
// Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
// Copying or reuse without permission is strictly forbidden.
// </copyright>
// ---------------------------------------------------------------------------------

using System;
using Documents.Utilities;

static class StringExtensions
{

  /// <summary>
  /// Compares the current source string instance to each of the specified items
  /// </summary>
  /// <param name="source">The source string</param>
  /// <param name="items">The items to compare the source string against</param>
  /// <returns>True if the source string is like one of the specified items</returns>
  /// <remarks>Assumes the comparison is case sensitive</remarks>
  public static bool IsLike(this string source, params string[] items)
  {
    return IsLike(source, false, items);
  }

  /// <summary>
  /// Compares the current source string instance to each of the specified items
  /// </summary>
  /// <param name="source">The source string</param>
  /// <param name="caseInsensitive">Detemines whether or not the comparison is case sensitive</param>
  /// <param name="items">The items to compare the source string against</param>
  /// <returns>True if the source string is like one of the specified items</returns>
  /// <remarks></remarks>
  public static bool IsLike(this string source, bool caseInsensitive, params string[] items)
  {
    try
    {
      for (int lintItemCounter = 0; lintItemCounter <= items.Length - 1; lintItemCounter++)
      {
        if (string.Compare(source, items[lintItemCounter], caseInsensitive) == 0)
          return true;
      }
      return false;
    }
    catch (Exception ex)
    {
      ApplicationLogging.LogException(ex, System.Reflection.MethodBase.GetCurrentMethod());
      // Re-throw the exception to the caller
      throw;
    }
  }
}
