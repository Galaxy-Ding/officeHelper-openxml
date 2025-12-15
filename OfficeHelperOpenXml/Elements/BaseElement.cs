using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using OfficeHelperOpenXml.Interfaces;

namespace OfficeHelperOpenXml.Elements
{
    public abstract class BaseElement : IElement
    {
        public abstract string ElementType { get; }
        public string Name { get; set; }
        public string SpecialType { get; set; }
        private readonly Dictionary<string, IElementComponent> _components;

        protected BaseElement()
        {
            Name = "";
            SpecialType = "";
            _components = new Dictionary<string, IElementComponent>();
        }

        protected abstract void InitializeComponents(Shape shape, SlidePart slidePart);

        public void InitializeFromShape(Shape shape, SlidePart slidePart)
        {
            var nvSpPr = shape.NonVisualShapeProperties;
            if (nvSpPr?.NonVisualDrawingProperties != null)
            {
                Name = nvSpPr.NonVisualDrawingProperties.Name ?? "";
            }
            InitializeComponents(shape, slidePart);
        }

        public T GetComponent<T>() where T : class, IElementComponent
        {
            var componentType = typeof(T).Name.Replace("Component", "");
            return _components.TryGetValue(componentType, out var component) ? component as T : null;
        }

        public void AddComponent(IElementComponent component)
        {
            if (component == null) return;
            _components[component.ComponentType] = component;
        }

        public void RemoveComponent(string componentType) => _components.Remove(componentType);
        public bool HasComponent(string componentType) => _components.ContainsKey(componentType);
        public IEnumerable<IElementComponent> GetAllComponents() => _components.Values;

        public virtual void ApplyToShape(Shape shape, SlidePart slidePart)
        {
            if (shape == null) return;
            foreach (var component in _components.Values.Where(c => c.IsEnabled))
            {
                try { component.ApplyToShape(shape, slidePart); }
                catch { }
            }
        }

        public virtual string ToJson()
        {
            var parts = new List<string> { $"\"type\":\"{ElementType.ToLower()}\"", $"\"name\":\"{Name}\"" };
            
            // Add special_type if present
            if (!string.IsNullOrEmpty(SpecialType))
            {
                parts.Add($"\"special_type\":\"{SpecialType}\"");
            }
            
            var nested = new Dictionary<string, string> { {"Shadow","shadow"}, {"Fill","fill"}, {"Line","line"}, {"Picture","picture"} };

            foreach (var c in _components.Values.Where(x => x.IsEnabled))
            {
                try
                {
                    var json = c.ToJson();
                    if (json != "null")
                    {
                        var content = json.Trim();
                        if (nested.TryGetValue(c.ComponentType, out var field))
                        {
                            // Components now return unwrapped content
                            parts.Add($"\"{field}\":{{{content}}}");
                        }
                        else
                        {
                            // Position and Text components return unwrapped content
                            parts.Add(content);
                        }
                    }
                }
                catch { }
            }
            return "{" + string.Join(",", parts) + "}";
        }
    }
}
