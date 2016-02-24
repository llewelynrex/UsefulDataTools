using System;
using System.Collections;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;

namespace UsefulDataTools
{
    public static class XmlOutputExtensions
    {
        public static XDocument ToXml<T>(this T input, string path = null, string encoding = "UTF-8", string rootElementName = "Root")
        {
            if (input == null)
                throw new ArgumentNullException(nameof(input));

            var xDocument = new XDocument(new XDeclaration("1.0", encoding, null));
            var rootXElement = new XElement(XName.Get(rootElementName));

            xDocument.Add(rootXElement);

            var hashCode = input.GetHashCode();
            var hashCodeElement = new HashCodeElement(hashCode);

            var type = input.GetType();
            if (type.IsSimpleType())
            {
                input.Process(xDocument, rootXElement, type.Name, string.Empty, type, _ => input, hashCodeElement);
                return xDocument;
            }


            var enumerable = input as IEnumerable;
            if (enumerable != null)
                ToXmlInternal(enumerable, xDocument, rootXElement, hashCodeElement);
            else
                ToXmlInternal(input, xDocument, rootXElement, hashCodeElement);

            if (!string.IsNullOrEmpty(path))
            {
                using (var fileStream = new FileStream(path, FileMode.Create))
                {
                    xDocument.Save(fileStream);
                }
            }

            return xDocument;
        }

        private static void ToXmlInternal(this IEnumerable enumerable, XDocument xDocument, XContainer parentXContainer, HashCodeElement parentHashCodeElement)
        {
            var enumerableType = enumerable.GetType();
            var genericTypeNames = string.Join(",", enumerableType.GenericTypeArguments.Select(x => x.FullName));
            var childXElement = new XElement(XName.Get($"IEnumerable_{genericTypeNames}_"));
            var childTypeXAttribute = new XAttribute(XName.Get("Type"), $"IEnumerable_{genericTypeNames}_");
            childXElement.Add(childTypeXAttribute);
            parentXContainer.Add(childXElement);

            var hashCode = enumerableType.GetHashCode();
            var hashCodeElement = new HashCodeElement(hashCode) {ParentHashCodeElement = parentHashCodeElement};

            if (hashCodeElement.HashKeyExistsInStructure(hashCode))
            {
                var childRecursiveXAttribute = new XAttribute(XName.Get("Recursive"), hashCode.ToString());
                childXElement.Add(childRecursiveXAttribute);
                return;
            }

            foreach (var item in enumerable)
                item.ToXmlInternal(xDocument, childXElement, hashCodeElement);
        }

        private static void ToXmlInternal<T>(this T element, XDocument xDocument, XContainer parentXContainer, HashCodeElement parentHashCodeElement)
        {
            var itemType = element.GetType();
            var properties = itemType.GetProperties();
            var fields = itemType.GetFields();
            var hashCode = element.GetHashCode();
            var hashCodeElement = new HashCodeElement(hashCode) {ParentHashCodeElement = parentHashCodeElement};

            var currentXElement = new XElement(XName.Get(itemType.Name));
            var currentXElementTypeXAttribute = new XAttribute(XName.Get("Type"),itemType.FullName);
            currentXElement.Add(currentXElementTypeXAttribute);
            parentXContainer.Add(currentXElement);

            if (hashCodeElement.HashKeyExistsInStructure(hashCode))
            {
                var currentXElementRecursiveXAttribute = new XAttribute(XName.Get("Recursive"),hashCode.ToString());
                currentXElement.Add(currentXElementRecursiveXAttribute);
                return;
            }

            foreach (var property in properties)
                ProcessProperty(element, xDocument, hashCodeElement, property, currentXElement);

            foreach (var field in fields)
                ProcessField(element, xDocument, hashCodeElement, field, currentXElement);
        }

        private static void ProcessProperty<T>(T element, XDocument xmlDocument, HashCodeElement parentHashCodeElement, PropertyInfo property, XContainer currentElement)
        {
            element.Process(xmlDocument, currentElement, property.Name, "Property", property.PropertyType, property.GetValue, parentHashCodeElement);
        }

        private static void ProcessField<T>(T element, XDocument xmlDocument, HashCodeElement parentHashCodeElement, FieldInfo field, XContainer currentXContainer)
        {
            element.Process(xmlDocument, currentXContainer, field.Name, "Field", field.FieldType, field.GetValue, parentHashCodeElement);
        }

        private static void Process<T>(this T element, XDocument xDocument, XContainer currentXContainer, string name, string memberType, Type type, Func<object, object> getValueFunction, HashCodeElement parentHashCodeElement)
        {
            var childXElement = new XElement(XName.Get($"{name}"));
            currentXContainer.Add(childXElement);
            if (!string.IsNullOrEmpty(memberType))
            {
                var memberTypeXAttribute = new XAttribute(XName.Get("MemberType"),memberType);
                childXElement.Add(memberTypeXAttribute);
            }

            if (type.IsSimpleType())
            {
                var underlyingType = Nullable.GetUnderlyingType(type);
                if (underlyingType != null)
                {
                    var childXElementTypeXAttrbute = new XAttribute(XName.Get("Type"), $"{underlyingType.FullName}?");
                    childXElement.Add(childXElementTypeXAttrbute);
                }
                else
                {
                    var childXElementTypeXAttrbute = new XAttribute(XName.Get("Type"), type.FullName);
                    childXElement.Add(childXElementTypeXAttrbute);
                }

                if (getValueFunction.Invoke(element) == null)
                    return;

                if (type == typeof (string))
                {
                    var xCData = new XCData(getValueFunction.Invoke(element).ToString());
                    childXElement.Add(xCData);
                }
                else
                    childXElement.Value = getValueFunction.Invoke(element).ToString();

                return;
            }

            var implementsIEnumerable = false;
            var interfaces = type.GetInterfaces();
            if (interfaces.Any())
                implementsIEnumerable = interfaces.Any(x => x.FullName == "System.Collections.IEnumerable");

            if (implementsIEnumerable)
            {
                var genericTypeNames = string.Join(",", type.GenericTypeArguments.Select(x => x.FullName));
                var childXElementTypeAttrbute = new XAttribute(XName.Get("Type"), $"IEnumerable_{genericTypeNames}_");
                childXElement.Add(childXElementTypeAttrbute);
            }
            else
            {
                var childXElementTypeAttrbute = new XAttribute(XName.Get("Type"), type.FullName);
                childXElement.Add(childXElementTypeAttrbute);
            }

            if (getValueFunction.Invoke(element) == null)
                return;

            var item = getValueFunction.Invoke(element);
            var enumerable = item as IEnumerable;
            if (enumerable != null)
                enumerable.ToXmlInternal(xDocument, childXElement, parentHashCodeElement);
            else
                item.ToXmlInternal(xDocument, childXElement, parentHashCodeElement);
        }
    }
}