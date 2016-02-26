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
        /// <summary>
        /// Generates a <see cref="XDocument"/> with a default encoding of UTF-8 and a root element named Root from the 
        /// object <see cref="input"/> of type <see cref="T"/> by recursively evaluating each field and property until 
        /// only simple types are reached.
        /// </summary>
        /// <typeparam name="T">Any type for which an <see cref="XDocument"/> is required.</typeparam>
        /// <param name="input">The input object which will be recursively evaluated.</param>
        /// <returns><see cref="XDocument"/></returns>
        public static XDocument ToXml<T>(this T input)
        {
            return input.ToXml(null, "UTF-8", "Root");
        }

        /// <summary>
        /// Generates a <see cref="XDocument"/> with a default encoding of UTF-8 and a root element named Root from the 
        /// object <see cref="input"/> of type <see cref="T"/> by recursively evaluating each field and property until 
        /// only simple types are reached. The result is saved to <see cref="path"/>.
        /// </summary>
        /// <typeparam name="T">Any type for which an <see cref="XDocument"/> is required.</typeparam>
        /// <param name="input">The input object which will be recursively evaluated.</param>
        /// <param name="path">If the path is specified, the generated <see cref="XDocument"/> will be saved as an xml file.</param>
        /// <returns><see cref="XDocument"/></returns>
        public static XDocument ToXml<T>(this T input, string path)
        {
            return input.ToXml(path, "UTF-8", "Root");
        }

        /// <summary>
        /// Generates a <see cref="XDocument"/> from the object <see cref="input"/> of type <see cref="T"/> by recursively
        /// evaluating each field and property until only simple types are reached.
        /// </summary>
        /// <typeparam name="T">Any type for which an <see cref="XDocument"/> is required.</typeparam>
        /// <param name="input">The input object which will be recursively evaluated.</param>
        /// <param name="path">If the path is specified, the generated <see cref="XDocument"/> will be saved as an xml file.</param>
        /// <param name="encoding">The encoding which will be used for the <see cref="XDeclaration"/>.</param>
        /// <param name="rootElementName">The name of the root element of the generated <see cref="XDocument"/>.</param>
        /// <returns><see cref="XDocument"/></returns>
        public static XDocument ToXml<T>(this T input, string path, string encoding, string rootElementName)
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

            if (hashCodeElement.HashCodeExistsInStructure(hashCode))
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

            if (hashCodeElement.HashCodeExistsInStructure(hashCode))
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