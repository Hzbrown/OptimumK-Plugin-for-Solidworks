using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace sw_drawer
{
    internal static class SolidWorksHelperUtils
    {
        internal static bool TryEnableCoincidentMateAxisAlignment(
            ModelDoc2 swModel,
            Feature mateFeature,
            Action<string> logWarning = null)
        {
            if (swModel == null || mateFeature == null)
            {
                return false;
            }

            // Macro-style first: select mate and invoke AddMate5 with macro-like arguments.
            if (TryEnableCoincidentMateAxisAlignmentByMacro(swModel, mateFeature, logWarning))
            {
                return true;
            }

            // Fallback: edit mate definition directly.
            return TryEnableCoincidentMateAxisAlignmentByDefinition(swModel, mateFeature);
        }

        internal static string NormalizeComponentNameCore(string name, bool stripBracketWrapper)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return string.Empty;
            }

            string normalized = name.Trim();

            if (stripBracketWrapper)
            {
                int leftBracket = normalized.IndexOf('[');
                int rightBracket = normalized.IndexOf(']');
                if (leftBracket >= 0 && rightBracket > leftBracket)
                {
                    normalized = normalized.Substring(leftBracket + 1, rightBracket - leftBracket - 1).Trim();
                }
            }

            int caretIdx = normalized.IndexOf('^');
            if (caretIdx > 0)
            {
                normalized = normalized.Substring(0, caretIdx);
            }

            int angleIdx = normalized.IndexOf('<');
            if (angleIdx > 0)
            {
                normalized = normalized.Substring(0, angleIdx);
            }

            int dashIdx = normalized.LastIndexOf('-');
            if (dashIdx > 0 && dashIdx < normalized.Length - 1)
            {
                bool allDigits = true;
                for (int i = dashIdx + 1; i < normalized.Length; i++)
                {
                    if (!char.IsDigit(normalized[i]))
                    {
                        allDigits = false;
                        break;
                    }
                }

                if (allDigits)
                {
                    normalized = normalized.Substring(0, dashIdx);
                }
            }

            return normalized.Trim();
        }

        private static bool TryEnableCoincidentMateAxisAlignmentByDefinition(ModelDoc2 swModel, Feature mateFeature)
        {
            try
            {
                object mateDefinition = mateFeature.GetDefinition();
                if (mateDefinition == null)
                {
                    return false;
                }

                Type definitionType = mateDefinition.GetType();
                var accessSelections = definitionType.GetMethod("AccessSelections");
                bool selectionAccessOpen = false;

                if (accessSelections != null)
                {
                    object accessResult = accessSelections.Invoke(mateDefinition, new object[] { swModel, null });
                    if (accessResult is bool && !(bool)accessResult)
                    {
                        return false;
                    }
                    selectionAccessOpen = true;
                }

                bool changed = false;
                changed |= TrySetIntProperty(mateDefinition, "MateAlignment", (int)swMateAlign_e.swMateAlignALIGNED);
                changed |= TrySetBooleanProperty(mateDefinition, "AlignAxes", true);
                changed |= TrySetBooleanProperty(mateDefinition, "AlignAxis", true);

                if (!changed)
                {
                    if (selectionAccessOpen)
                    {
                        var releaseSelections = definitionType.GetMethod("ReleaseSelectionAccess");
                        releaseSelections?.Invoke(mateDefinition, null);
                    }
                    return false;
                }

                bool modified = mateFeature.ModifyDefinition(mateDefinition, swModel, null);

                if (selectionAccessOpen)
                {
                    var releaseSelections = definitionType.GetMethod("ReleaseSelectionAccess");
                    releaseSelections?.Invoke(mateDefinition, null);
                }

                return modified;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryEnableCoincidentMateAxisAlignmentByMacro(
            ModelDoc2 swModel,
            Feature mateFeature,
            Action<string> logWarning)
        {
            try
            {
                string mateName = mateFeature.Name;
                if (string.IsNullOrWhiteSpace(mateName))
                {
                    return false;
                }

                swModel.ClearSelection2(true);
                bool selected = swModel.Extension.SelectByID2(
                    mateName,
                    "MATE",
                    0, 0, 0,
                    false,
                    0,
                    null,
                    (int)swSelectOption_e.swSelectOptionDefault);

                if (!selected)
                {
                    swModel.ClearSelection2(true);
                    return false;
                }

                object modelObject = swModel;
                Type modelType = modelObject.GetType();
                System.Reflection.MethodInfo[] methods = modelType.GetMethods();

                for (int i = 0; i < methods.Length; i++)
                {
                    System.Reflection.MethodInfo method = methods[i];
                    if (!string.Equals(method.Name, "AddMate5", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    System.Reflection.ParameterInfo[] parameters = method.GetParameters();
                    if (!TryBuildMacroAddMate5Arguments(parameters, out object[] args, out int errorIndex))
                    {
                        continue;
                    }

                    object invokeResult;
                    try
                    {
                        invokeResult = method.Invoke(modelObject, args);
                    }
                    catch
                    {
                        continue;
                    }

                    int errorStatus = 0;
                    if (errorIndex >= 0 && errorIndex < args.Length && args[errorIndex] is int)
                    {
                        errorStatus = (int)args[errorIndex];
                    }

                    if (invokeResult != null && errorStatus == 0)
                    {
                        swModel.ClearSelection2(true);
                        return true;
                    }
                }

                swModel.ClearSelection2(true);
                return false;
            }
            catch (Exception ex)
            {
                logWarning?.Invoke($"Macro-style mate alignment fallback failed: {ex.Message}");
                try { swModel.ClearSelection2(true); } catch { }
                return false;
            }
        }

        private static bool TryBuildMacroAddMate5Arguments(
            System.Reflection.ParameterInfo[] parameters,
            out object[] args,
            out int errorIndex)
        {
            args = null;
            errorIndex = -1;

            if (parameters == null || parameters.Length == 0)
            {
                return false;
            }

            int[] intValues = new[]
            {
                (int)swMateType_e.swMateCOORDINATE,
                (int)swMateAlign_e.swMateAlignALIGNED,
                0
            };

            bool[] boolValues = new[]
            {
                false,
                false,
                false,
                false
            };

            double[] doubleValues = new[]
            {
                0.0,
                0.001, 0.001, 0.001, 0.001,
                0.5235987755983, 0.5235987755983, 0.5235987755983
            };

            int intCursor = 0;
            int boolCursor = 0;
            int doubleCursor = 0;

            args = new object[parameters.Length];

            for (int i = 0; i < parameters.Length; i++)
            {
                Type paramType = parameters[i].ParameterType;
                Type elementType = paramType.IsByRef ? paramType.GetElementType() : paramType;

                if (parameters[i].IsOut || (paramType.IsByRef && elementType == typeof(int)))
                {
                    args[i] = 0;
                    errorIndex = i;
                    continue;
                }

                if (elementType == typeof(int))
                {
                    args[i] = intCursor < intValues.Length ? intValues[intCursor++] : 0;
                    continue;
                }

                if (elementType == typeof(bool))
                {
                    args[i] = boolCursor < boolValues.Length ? boolValues[boolCursor++] : false;
                    continue;
                }

                if (elementType == typeof(double))
                {
                    args[i] = doubleCursor < doubleValues.Length ? doubleValues[doubleCursor++] : 0.0;
                    continue;
                }

                if (elementType == typeof(object))
                {
                    args[i] = null;
                    continue;
                }

                return false;
            }

            return true;
        }

        private static bool TrySetBooleanProperty(object target, string propertyName, bool value)
        {
            if (target == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return false;
            }

            try
            {
                var propertyInfo = target.GetType().GetProperty(
                    propertyName,
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.IgnoreCase);

                if (propertyInfo == null || !propertyInfo.CanWrite || propertyInfo.PropertyType != typeof(bool))
                {
                    return false;
                }

                propertyInfo.SetValue(target, value, null);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TrySetIntProperty(object target, string propertyName, int value)
        {
            if (target == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return false;
            }

            try
            {
                var propertyInfo = target.GetType().GetProperty(
                    propertyName,
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.IgnoreCase);

                if (propertyInfo == null || !propertyInfo.CanWrite)
                {
                    return false;
                }

                Type propertyType = propertyInfo.PropertyType;
                if (propertyType == typeof(int))
                {
                    propertyInfo.SetValue(target, value, null);
                    return true;
                }

                if (propertyType.IsEnum)
                {
                    object enumValue = Enum.ToObject(propertyType, value);
                    propertyInfo.SetValue(target, enumValue, null);
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }
    }
}