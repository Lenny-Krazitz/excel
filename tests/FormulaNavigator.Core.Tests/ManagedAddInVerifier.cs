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
