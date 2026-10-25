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
                    var embeddedTypes = new HashSet<string>(StringComparer.Ordinal);
                    foreach (TypeDefinitionHandle handle in metadata.TypeDefinitions)
                    {
                        TypeDefinition type = metadata.GetTypeDefinition(handle);
                        if (metadata.GetString(type.Namespace) == "Microsoft.Office.Interop.Excel")
                            embeddedTypes.Add(metadata.GetString(type.Name));
                    }
                    if (!embeddedTypes.Contains("Workbook") || !embeddedTypes.Contains("Range"))
                        throw new InvalidDataException("Expected embedded Excel event argument types are missing.");
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
                Console.WriteLine("Managed add-in verification passed: Excel interop types are embedded; no separate Office PIA is required.");
                return 0;
            }
            catch (Exception error)
            {
                Console.Error.WriteLine("Managed add-in verification failed: " + error.Message);
                return 1;
            }
        }
    }
}
