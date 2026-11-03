using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection.Metadata;
using System.Reflection.PortableExecutable;

namespace FormulaNavigator.Core.Tests
{
    internal static class ManagedAddInVerifier
    {
        public static int Verify(string path)
        {
            try
            {
                using (var stream = File.OpenRead(path))
                using (var pe = new PEReader(stream))
                {
                    MetadataReader metadata = pe.GetMetadataReader();
                    foreach (AssemblyReferenceHandle handle in metadata.AssemblyReferences)
                    {
                        string name = metadata.GetString(metadata.GetAssemblyReference(handle).Name);
                        if (name == "Microsoft.Office.Interop.Excel" || name == "office" || name == "stdole")
                            throw new InvalidDataException("Unpacked Office interop runtime dependency: " + name);
                    }
                    VerifyOnDemandArchitecture(metadata);
                    var gatewayMethods = new HashSet<string>(StringComparer.Ordinal);
                    var nativeHostMethods = new HashSet<string>(StringComparer.Ordinal);
                    foreach (TypeDefinitionHandle handle in metadata.TypeDefinitions)
                    {
                        TypeDefinition type = metadata.GetTypeDefinition(handle);
                        if (metadata.GetString(type.Namespace) == "Microsoft.Office.Interop.Excel")
                            throw new InvalidDataException("Unused Excel event interop types must not be embedded.");
                        string typeName = metadata.GetString(type.Name);
                        if (typeName == "WorkbookIndex" || typeName == "FormulaSnapshotReader"
                            || typeName == "DynamicSnapshotReader")
                            throw new InvalidDataException("Removed background index type returned: " + typeName);
                        foreach (MethodDefinitionHandle methodHandle in type.GetMethods())
                        {
                            string methodName = metadata.GetString(metadata.GetMethodDefinition(methodHandle).Name);
                            if (typeName == "ExcelGateway") gatewayMethods.Add(methodName);
                            if (typeName == "ExcelNativeTraceHost") nativeHostMethods.Add(methodName);
                        }
                    }
                    RequireMethods(gatewayMethods, "TraceDependents", "ReadDependentCells", "CaptureActiveCell",
                        "Inspect", "Navigate", "Dispose");
                    RequireMethods(nativeHostMethods, "SaveState", "CheckUserArrows", "CreateArrows",
                        "Navigate", "RemoveCreatedArrows", "RestoreState");
                    foreach (string removed in new[] { "StartIndexing", "PumpIndexes", "ReconcileWorkbooks",
                        "FindIndexDependentsAsync", "BuildSnapshotAsync", "PumpDynamicIndex",
                        "OnSheetCalculated", "OnSheetChanged" })
                        if (gatewayMethods.Contains(removed))
                            throw new InvalidDataException("Unused background index method returned: " + removed);
                    var commands = new HashSet<string>(StringComparer.Ordinal);
                    foreach (TypeDefinitionHandle handle in metadata.TypeDefinitions)
                    {
                        TypeDefinition type = metadata.GetTypeDefinition(handle);
                        foreach (MethodDefinitionHandle methodHandle in type.GetMethods())
                        {
                            MethodDefinition method = metadata.GetMethodDefinition(methodHandle);
                            foreach (CustomAttributeHandle attributeHandle in method.GetCustomAttributes())
                            {
                                CustomAttribute attribute = metadata.GetCustomAttribute(attributeHandle);
                                EntityHandle constructor = attribute.Constructor;
                                TypeDefinition declaringType;
                                if (constructor.Kind == HandleKind.MemberReference)
                                {
                                    MemberReference member = metadata.GetMemberReference((MemberReferenceHandle)constructor);
                                    if (member.Parent.Kind != HandleKind.TypeReference) continue;
                                    TypeReference reference = metadata.GetTypeReference((TypeReferenceHandle)member.Parent);
                                    if (metadata.GetString(reference.Name) != "ExcelCommandAttribute") continue;
                                }
                                else if (constructor.Kind == HandleKind.MethodDefinition)
                                {
                                    MethodDefinition constructorMethod = metadata.GetMethodDefinition((MethodDefinitionHandle)constructor);
                                    declaringType = metadata.GetTypeDefinition(constructorMethod.GetDeclaringType());
                                    if (metadata.GetString(declaringType.Name) != "ExcelCommandAttribute") continue;
                                }
                                else continue;
                                commands.Add(metadata.GetString(method.Name));
                            }
                        }
                    }
                    if (commands.Count != 2 || !commands.Contains("Explore") || !commands.Contains("FindDependents")
                        || commands.Contains("CompareNativeDependents") || commands.Contains("ShowStatus")
                        || commands.Contains("RebindKeys"))
                        throw new InvalidDataException("Exactly two Excel commands (Explore, FindDependents) must be exported; found: "
                            + string.Join(", ", commands));
                }
                Console.WriteLine("Managed add-in verification passed: on-demand-only architecture; no background timers, "
                    + "workers or Excel event sinks; native trace and navigation preserved; no Office PIA required.");
                return 0;
            }
            catch (Exception error)
            {
                Console.Error.WriteLine("Managed add-in verification failed: " + error.Message);
                return 1;
            }
        }

        private static void RequireMethods(HashSet<string> actual, params string[] required)
        {
            foreach (string name in required)
                if (!actual.Contains(name))
                    throw new InvalidDataException("Required on-demand navigation/trace method missing: " + name);
        }

        private static void VerifyOnDemandArchitecture(MetadataReader metadata)
        {
            // Inspect the compiled assembly, not just the startup source. UI one-shot
            // Dispatcher.BeginInvoke and native host Task.FromResult are allowed.
            // Periodic schedulers and task producers are not needed by either command.
            foreach (TypeReferenceHandle handle in metadata.TypeReferences)
            {
                TypeReference type = metadata.GetTypeReference(handle);
                string ns = metadata.GetString(type.Namespace);
                string name = metadata.GetString(type.Name);
                if ((ns == "System.Windows.Threading" && name == "DispatcherTimer")
                    || (ns == "System.Threading" && (name == "Timer" || name == "ThreadPool"))
                    || (ns == "System.Timers" && name == "Timer")
                    || (ns == "System.Windows.Forms" && name == "Timer")
                    || (ns == "Microsoft.Office.Interop.Excel" && name.StartsWith("AppEvents", StringComparison.Ordinal)))
                    throw new InvalidDataException("Unexpected background scheduler/event sink: " + ns + "." + name);
            }
            foreach (MemberReferenceHandle handle in metadata.MemberReferences)
            {
                MemberReference member = metadata.GetMemberReference(handle);
                string name = metadata.GetString(member.Name);
                if (name.StartsWith("add_Workbook", StringComparison.Ordinal)
                    || name.StartsWith("add_Sheet", StringComparison.Ordinal)
                    || name == "add_NewWorkbook")
                    throw new InvalidDataException("Unexpected automatic Excel event subscription: " + name);
                if (member.Parent.Kind != HandleKind.TypeReference) continue;
                TypeReference parent = metadata.GetTypeReference((TypeReferenceHandle)member.Parent);
                if (metadata.GetString(parent.Namespace) == "System.Threading.Tasks"
                    && (name == "Run" || name == "Delay" || name == "StartNew"
                        || name == "ContinueWith" || name == "get_Factory"))
                    throw new InvalidDataException("Unexpected background task producer: " + name);
            }
        }
    }
}
