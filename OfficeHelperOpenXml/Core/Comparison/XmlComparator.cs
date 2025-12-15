using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace OfficeHelperOpenXml.Core.Comparison
{
    /// <summary>
    /// Compares two XML files and identifies structural and content differences
    /// </summary>
    public class XmlComparator
    {
        /// <summary>
        /// Compare two XML files and identify all differences
        /// </summary>
        /// <param name="xmlPath1">Path to first XML file (original)</param>
        /// <param name="xmlPath2">Path to second XML file (generated)</param>
        /// <returns>XmlComparisonResult containing all identified differences</returns>
        public XmlComparisonResult Compare(string xmlPath1, string xmlPath2)
        {
            var result = new XmlComparisonResult
            {
                FilePath = xmlPath1
            };

            try
            {
                // Parse both XML files
                XDocument doc1 = XDocument.Load(xmlPath1);
                XDocument doc2 = XDocument.Load(xmlPath2);

                // Compare root elements
                if (doc1.Root != null && doc2.Root != null)
                {
                    var differences = CompareElements(doc1.Root, doc2.Root, GetElementPath(doc1.Root));
                    result.Differences.AddRange(differences);
                }
                else if (doc1.Root == null && doc2.Root != null)
                {
                    result.Differences.Add(new XmlDifference
                    {
                        XPath = "/",
                        Type = DifferenceType.MissingElement,
                        ExpectedValue = "Root element",
                        ActualValue = "(missing)",
                        Description = "First XML file has no root element"
                    });
                }
                else if (doc1.Root != null && doc2.Root == null)
                {
                    result.Differences.Add(new XmlDifference
                    {
                        XPath = "/",
                        Type = DifferenceType.ExtraElement,
                        ExpectedValue = "(missing)",
                        ActualValue = "Root element",
                        Description = "Second XML file has no root element"
                    });
                }
            }
            catch (XmlException ex)
            {
                throw new ComparisonException($"Failed to parse XML file: {ex.Message}", ex)
                {
                    Context = "XML Parsing",
                    FilePath = xmlPath1,
                    Category = ErrorCategory.Xml
                };
            }
            catch (Exception ex)
            {
                throw new ComparisonException($"Error comparing XML files: {ex.Message}", ex)
                {
                    Context = "XML Comparison",
                    FilePath = xmlPath1,
                    Category = ErrorCategory.System
                };
            }

            return result;
        }

        /// <summary>
        /// Recursively compare two XML elements and their children
        /// </summary>
        /// <param name="elem1">First element (expected)</param>
        /// <param name="elem2">Second element (actual)</param>
        /// <param name="xpath">Current XPath location</param>
        /// <returns>List of differences found</returns>
        private List<XmlDifference> CompareElements(XElement elem1, XElement elem2, string xpath)
        {
            var differences = new List<XmlDifference>();

            // Compare element names
            if (elem1.Name != elem2.Name)
            {
                differences.Add(new XmlDifference
                {
                    XPath = xpath,
                    Type = DifferenceType.DifferentAttributeValue,
                    ExpectedValue = elem1.Name.ToString(),
                    ActualValue = elem2.Name.ToString(),
                    Description = "Element names differ"
                });
                // If names differ, don't continue comparing children
                return differences;
            }

            // Compare attributes
            var attrDifferences = CompareAttributes(elem1, elem2, xpath);
            differences.AddRange(attrDifferences);

            // Compare text content (only if element has no child elements)
            var elem1Children = elem1.Elements().ToList();
            var elem2Children = elem2.Elements().ToList();

            if (elem1Children.Count == 0 && elem2Children.Count == 0)
            {
                // Leaf node - compare text content
                string text1 = NormalizeWhitespace(elem1.Value);
                string text2 = NormalizeWhitespace(elem2.Value);

                if (text1 != text2)
                {
                    differences.Add(new XmlDifference
                    {
                        XPath = xpath,
                        Type = DifferenceType.DifferentTextContent,
                        ExpectedValue = text1,
                        ActualValue = text2,
                        Description = "Text content differs"
                    });
                }
            }

            // Compare child elements
            var childDifferences = CompareChildElements(elem1Children, elem2Children, xpath);
            differences.AddRange(childDifferences);

            return differences;
        }

        /// <summary>
        /// Compare attributes of two elements
        /// </summary>
        /// <param name="elem1">First element (expected)</param>
        /// <param name="elem2">Second element (actual)</param>
        /// <param name="xpath">Current XPath location</param>
        /// <returns>List of attribute differences</returns>
        private List<XmlDifference> CompareAttributes(XElement elem1, XElement elem2, string xpath)
        {
            var differences = new List<XmlDifference>();

            var attrs1 = elem1.Attributes().ToDictionary(a => a.Name, a => a.Value);
            var attrs2 = elem2.Attributes().ToDictionary(a => a.Name, a => a.Value);

            // Find missing attributes (in elem1 but not in elem2)
            foreach (var attr in attrs1)
            {
                // Skip relationship ID attributes - they don't affect functionality
                if (ShouldIgnoreAttribute(attr.Key, elem1))
                {
                    continue;
                }

                if (!attrs2.ContainsKey(attr.Key))
                {
                    differences.Add(new XmlDifference
                    {
                        XPath = $"{xpath}/@{attr.Key}",
                        Type = DifferenceType.MissingAttribute,
                        ExpectedValue = attr.Value,
                        ActualValue = "(missing)",
                        Description = $"Attribute '{attr.Key}' is missing in second element"
                    });
                }
                else if (attrs2[attr.Key] != attr.Value)
                {
                    // Attribute exists but value differs
                    differences.Add(new XmlDifference
                    {
                        XPath = $"{xpath}/@{attr.Key}",
                        Type = DifferenceType.DifferentAttributeValue,
                        ExpectedValue = attr.Value,
                        ActualValue = attrs2[attr.Key],
                        Description = $"Attribute '{attr.Key}' has different value"
                    });
                }
            }

            // Find extra attributes (in elem2 but not in elem1)
            foreach (var attr in attrs2)
            {
                // Skip relationship ID attributes - they don't affect functionality
                if (ShouldIgnoreAttribute(attr.Key, elem2))
                {
                    continue;
                }

                if (!attrs1.ContainsKey(attr.Key))
                {
                    differences.Add(new XmlDifference
                    {
                        XPath = $"{xpath}/@{attr.Key}",
                        Type = DifferenceType.ExtraAttribute,
                        ExpectedValue = "(missing)",
                        ActualValue = attr.Value,
                        Description = $"Attribute '{attr.Key}' is extra in second element"
                    });
                }
            }

            return differences;
        }

        /// <summary>
        /// Determines if an attribute should be ignored during comparison
        /// </summary>
        /// <param name="attributeName">The attribute name to check</param>
        /// <returns>True if the attribute should be ignored, false otherwise</returns>
        private bool ShouldIgnoreAttribute(XName attributeName, XElement parentElement)
        {
            string localName = attributeName.LocalName;
            string attrNamespace = attributeName.NamespaceName;
            string parentNamespace = parentElement.Name.NamespaceName;

            // Ignore relationship ID attributes (Id attribute in .rels files)
            // These attributes have no namespace prefix, but their parent element is in the relationships namespace
            // Parent namespace: http://schemas.openxmlformats.org/package/2006/relationships
            if (localName == "Id" && string.IsNullOrEmpty(attrNamespace) && 
                parentNamespace.Contains("/relationships"))
            {
                return true;
            }

            // Ignore r:id, r:embed, r:link attributes (relationship references in content files)
            // These have the r: namespace prefix
            // The r: namespace is: http://schemas.openxmlformats.org/officeDocument/2006/relationships
            if (attrNamespace.Contains("/relationships") && 
                (localName == "id" || localName == "embed" || localName == "link"))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Compare child elements of two parent elements
        /// </summary>
        /// <param name="children1">Children of first element</param>
        /// <param name="children2">Children of second element</param>
        /// <param name="parentXPath">XPath of parent element</param>
        /// <returns>List of differences in child elements</returns>
        private List<XmlDifference> CompareChildElements(List<XElement> children1, List<XElement> children2, string parentXPath)
        {
            var differences = new List<XmlDifference>();

            // Group children by element name
            var groups1 = children1.GroupBy(e => e.Name).ToDictionary(g => g.Key, g => g.ToList());
            var groups2 = children2.GroupBy(e => e.Name).ToDictionary(g => g.Key, g => g.ToList());

            var allNames = new HashSet<XName>(groups1.Keys);
            allNames.UnionWith(groups2.Keys);

            foreach (var name in allNames)
            {
                var list1 = groups1.ContainsKey(name) ? groups1[name] : new List<XElement>();
                var list2 = groups2.ContainsKey(name) ? groups2[name] : new List<XElement>();

                // Compare counts
                if (list1.Count != list2.Count)
                {
                    if (list1.Count > list2.Count)
                    {
                        // Missing elements in second
                        for (int i = list2.Count; i < list1.Count; i++)
                        {
                            string xpath = $"{parentXPath}/{GetLocalName(name)}[{i + 1}]";
                            differences.Add(new XmlDifference
                            {
                                XPath = xpath,
                                Type = DifferenceType.MissingElement,
                                ExpectedValue = $"<{GetLocalName(name)}>",
                                ActualValue = "(missing)",
                                Description = $"Element '{GetLocalName(name)}' is missing in second XML"
                            });
                        }
                    }
                    else
                    {
                        // Extra elements in second
                        for (int i = list1.Count; i < list2.Count; i++)
                        {
                            string xpath = $"{parentXPath}/{GetLocalName(name)}[{i + 1}]";
                            differences.Add(new XmlDifference
                            {
                                XPath = xpath,
                                Type = DifferenceType.ExtraElement,
                                ExpectedValue = "(missing)",
                                ActualValue = $"<{GetLocalName(name)}>",
                                Description = $"Element '{GetLocalName(name)}' is extra in second XML"
                            });
                        }
                    }
                }

                // Compare matching elements
                int minCount = Math.Min(list1.Count, list2.Count);
                for (int i = 0; i < minCount; i++)
                {
                    string xpath = $"{parentXPath}/{GetLocalName(name)}[{i + 1}]";
                    var childDiffs = CompareElements(list1[i], list2[i], xpath);
                    differences.AddRange(childDiffs);
                }
            }

            return differences;
        }

        /// <summary>
        /// Normalize whitespace in text content to avoid false positives
        /// </summary>
        /// <param name="text">Text to normalize</param>
        /// <returns>Normalized text</returns>
        private string NormalizeWhitespace(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            // Replace all whitespace sequences with a single space
            text = Regex.Replace(text, @"\s+", " ");

            // Trim leading and trailing whitespace
            text = text.Trim();

            return text;
        }

        /// <summary>
        /// Get XPath for an element
        /// </summary>
        /// <param name="element">Element to get path for</param>
        /// <returns>XPath string</returns>
        private string GetElementPath(XElement element)
        {
            if (element == null)
            {
                return "/";
            }

            var path = new List<string>();
            var current = element;

            while (current != null)
            {
                string name = GetLocalName(current.Name);
                
                // Get position among siblings with same name
                var siblings = current.Parent?.Elements(current.Name).ToList();
                int position = 1;
                
                if (siblings != null && siblings.Count > 1)
                {
                    position = siblings.IndexOf(current) + 1;
                    path.Insert(0, $"{name}[{position}]");
                }
                else
                {
                    path.Insert(0, name);
                }

                current = current.Parent;
            }

            return "/" + string.Join("/", path);
        }

        /// <summary>
        /// Get local name from XName (without namespace)
        /// </summary>
        /// <param name="name">XName to extract local name from</param>
        /// <returns>Local name string</returns>
        private string GetLocalName(XName name)
        {
            return name.LocalName;
        }
    }
}
