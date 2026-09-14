using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace BBI.JD
{
    /// <summary>
    /// The CG_* shared parameters written onto every placed marker so the centre
    /// of gravity can be scheduled and tagged. Binding is instance-level on the
    /// marker family's category. GUIDs are fixed forever - schedules rely on them.
    /// </summary>
    public static class CgSharedParameters
    {
        public const string GroupName = "CenterGravity";

        public sealed class ParamDef
        {
            public string Name;
            public Guid Guid;
            public ForgeTypeId Spec;
        }

        public static readonly ParamDef Marker      = new() { Name = "CG_Marker",      Guid = new Guid("7d1a0e00-0000-4000-a000-000000000001"), Spec = SpecTypeId.Boolean.YesNo };
        public static readonly ParamDef LiftName    = new() { Name = "CG_LiftName",    Guid = new Guid("7d1a0e00-0000-4000-a000-000000000002"), Spec = SpecTypeId.String.Text };
        public static readonly ParamDef Weighting   = new() { Name = "CG_Weighting",   Guid = new Guid("7d1a0e00-0000-4000-a000-000000000003"), Spec = SpecTypeId.String.Text };
        public static readonly ParamDef Reference   = new() { Name = "CG_Reference",   Guid = new Guid("7d1a0e00-0000-4000-a000-000000000004"), Spec = SpecTypeId.String.Text };
        public static readonly ParamDef Coordinates = new() { Name = "CG_Coordinates", Guid = new Guid("7d1a0e00-0000-4000-a000-000000000005"), Spec = SpecTypeId.String.Text };
        public static readonly ParamDef X           = new() { Name = "CG_X",           Guid = new Guid("7d1a0e00-0000-4000-a000-000000000006"), Spec = SpecTypeId.Length };
        public static readonly ParamDef Y           = new() { Name = "CG_Y",           Guid = new Guid("7d1a0e00-0000-4000-a000-000000000007"), Spec = SpecTypeId.Length };
        public static readonly ParamDef Z           = new() { Name = "CG_Z",           Guid = new Guid("7d1a0e00-0000-4000-a000-000000000008"), Spec = SpecTypeId.Length };
        public static readonly ParamDef Volume      = new() { Name = "CG_Volume",      Guid = new Guid("7d1a0e00-0000-4000-a000-000000000009"), Spec = SpecTypeId.Volume };
        public static readonly ParamDef Weight      = new() { Name = "CG_Weight",      Guid = new Guid("7d1a0e00-0000-4000-a000-00000000000a"), Spec = SpecTypeId.Mass };
        public static readonly ParamDef Date        = new() { Name = "CG_Date",        Guid = new Guid("7d1a0e00-0000-4000-a000-00000000000b"), Spec = SpecTypeId.String.Text };

        public static readonly ParamDef[] All = { Marker, LiftName, Weighting, Reference, Coordinates, X, Y, Z, Volume, Weight, Date };

        /// <summary>
        /// Ensure every CG_* parameter is bound (instance) to <paramref name="category"/>.
        /// Opens its own transaction only when a binding is missing.
        /// </summary>
        public static void EnsureBound(Document doc, Category category)
        {
            if (doc == null || category == null)
            {
                return;
            }

            BindingMap map = doc.ParameterBindings;

            List<ParamDef> missing = new();
            foreach (ParamDef d in All)
            {
                if (!IsBoundToCategory(map, d.Name, category))
                {
                    missing.Add(d);
                }
            }

            if (missing.Count == 0)
            {
                return;
            }

            Application app = doc.Application;
            string previousFile = app.SharedParametersFilename;
            string tempFile = Path.Combine(Path.GetTempPath(), "CenterGravity_SharedParameters.txt");

            try
            {
                if (!File.Exists(tempFile))
                {
                    WriteSharedParameterFileHeader(tempFile);
                }

                app.SharedParametersFilename = tempFile;

                DefinitionFile file = app.OpenSharedParameterFile();
                if (file == null)
                {
                    return;
                }

                DefinitionGroup group = file.Groups.get_Item(GroupName) ?? file.Groups.Create(GroupName);

                using Transaction t = new(doc, "Bind Center Gravity parameters");
                t.Start();

                foreach (ParamDef d in missing)
                {
                    Definition definition = group.Definitions.get_Item(d.Name);

                    if (definition == null)
                    {
                        ExternalDefinitionCreationOptions options = new(d.Name, d.Spec) { GUID = d.Guid };
                        definition = group.Definitions.Create(options);
                    }

                    CategorySet categories = app.Create.NewCategorySet();
                    categories.Insert(category);

                    if (map.Contains(definition) && map.get_Item(definition) is ElementBinding current)
                    {
                        foreach (Category c in current.Categories)
                        {
                            categories.Insert(c);
                        }

                        map.ReInsert(definition, app.Create.NewInstanceBinding(categories), GroupTypeId.Data);
                    }
                    else
                    {
                        map.Insert(definition, app.Create.NewInstanceBinding(categories), GroupTypeId.Data);
                    }
                }

                t.Commit();
            }
            finally
            {
                try { app.SharedParametersFilename = previousFile; }
                catch (Exception) { /* leave whatever we could restore */ }
            }
        }

        private static bool IsBoundToCategory(BindingMap map, string name, Category category)
        {
            DefinitionBindingMapIterator it = map.ForwardIterator();
            it.Reset();

            while (it.MoveNext())
            {
                if (it.Key is Definition def && def.Name == name && it.Current is ElementBinding binding)
                {
                    foreach (Category c in binding.Categories)
                    {
                        if (c.Id == category.Id)
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        private static void WriteSharedParameterFileHeader(string path)
        {
            StringBuilder sb = new();
            sb.AppendLine("# This is a Revit shared parameter file.");
            sb.AppendLine("# Generated by CenterGravity - do not edit.");
            sb.AppendLine("*META\tVERSION\tMINVERSION");
            sb.AppendLine("META\t2\t1");
            sb.AppendLine("*GROUP\tID\tNAME");
            sb.AppendLine("GROUP\t1\t" + GroupName);
            sb.AppendLine("*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\tHIDEWHENNOVALUE");

            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        }
    }
}
