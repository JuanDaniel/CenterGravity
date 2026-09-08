using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace BBI.JD
{
    /// <summary>Supplies mass density (Revit internal units) for materials / elements.</summary>
    public interface IDensityProvider
    {
        /// <summary>Density for a material, or the configured fallback, or null when nothing is known.</summary>
        double? GetMaterialDensity(ElementId materialId);

        /// <summary>Density for an element (structural material, then any material), or the fallback, or null.</summary>
        double? GetElementDensity(Element element);
    }

    /// <summary>
    /// Reads density from each material's structural asset, with an optional user fallback
    /// used whenever a material has no usable density. Results are cached per material.
    /// </summary>
    public class RevitDensityProvider : IDensityProvider
    {
        private readonly Document document;
        private readonly double? fallbackInternal;
        private readonly Dictionary<long, double?> cache = new();

        public RevitDensityProvider(Document document, double? fallbackInternal)
        {
            this.document = document;
            this.fallbackInternal = fallbackInternal;
        }

        public double? GetMaterialDensity(ElementId materialId)
        {
            double? found = ResolveMaterialDensity(materialId);
            return found ?? fallbackInternal;
        }

        public double? GetElementDensity(Element element)
        {
            ElementId materialId = ElementId.InvalidElementId;

            Parameter p = element?.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM);
            if (p != null && p.StorageType == StorageType.ElementId)
            {
                materialId = p.AsElementId();
            }

            if ((materialId == null || materialId == ElementId.InvalidElementId) && element != null)
            {
                try
                {
                    foreach (ElementId id in element.GetMaterialIds(false))
                    {
                        materialId = id;
                        break;
                    }
                }
                catch (Exception)
                {
                    // element does not support material queries
                }
            }

            return GetMaterialDensity(materialId);
        }

        private double? ResolveMaterialDensity(ElementId materialId)
        {
            if (materialId == null || materialId == ElementId.InvalidElementId)
            {
                return null;
            }

            long key = materialId.Value;

            if (cache.TryGetValue(key, out double? cached))
            {
                return cached;
            }

            double? density = null;

            try
            {
                if (document.GetElement(materialId) is Material material
                    && material.StructuralAssetId != null
                    && material.StructuralAssetId != ElementId.InvalidElementId
                    && document.GetElement(material.StructuralAssetId) is PropertySetElement pse)
                {
                    StructuralAsset asset = pse.GetStructuralAsset();

                    if (asset != null && asset.Density > 0)
                    {
                        density = asset.Density;
                    }
                }
            }
            catch (Exception)
            {
                density = null;
            }

            cache[key] = density;
            return density;
        }
    }
}
