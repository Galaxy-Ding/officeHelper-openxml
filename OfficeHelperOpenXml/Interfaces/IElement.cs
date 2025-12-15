using System;
using System.Collections.Generic;

namespace OfficeHelperOpenXml.Interfaces
{
    public interface IElement
    {
        string ElementType { get; }
        string Name { get; set; }
        T GetComponent<T>() where T : class, IElementComponent;
        void AddComponent(IElementComponent component);
        void RemoveComponent(string componentType);
        bool HasComponent(string componentType);
        IEnumerable<IElementComponent> GetAllComponents();
        string ToJson();
    }
}
