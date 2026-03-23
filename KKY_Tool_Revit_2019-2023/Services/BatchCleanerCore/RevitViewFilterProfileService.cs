using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using XmlSaveOptions = System.Xml.Linq.SaveOptions;
using Autodesk.Revit.DB;
using KKY_Tool_Revit.Models;

namespace KKY_Tool_Revit.Services
{
    public static class RevitViewFilterProfileService
    {
        public static void SaveToXml(ViewFilterProfile profile, string filePath)
        {
            if (profile == null) throw new ArgumentNullException(nameof(profile));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));

            string definitionXml = profile.FilterDefinitionXml;
            if (string.IsNullOrWhiteSpace(definitionXml)
                && !string.IsNullOrWhiteSpace(profile.ParameterToken)
                && profile.RuleValue != null)
            {
                definitionXml = BuildSimpleDefinitionXml(profile);
            }

            var root = new XElement("ViewFilterProfile",
                new XAttribute("Version", "2"),
                new XElement("FilterName", profile.FilterName ?? string.Empty),
                new XElement("CategoriesCsv", profile.CategoriesCsv ?? string.Empty),
                new XElement("ParameterToken", profile.ParameterToken ?? string.Empty),
                new XElement("Operator", profile.Operator.ToString()),
                new XElement("RuleValue", profile.RuleValue ?? string.Empty),
                new XElement("StructureSummary", profile.StructureSummary ?? string.Empty));

            if (!string.IsNullOrWhiteSpace(definitionXml))
            {
                XElement definitionNode = TryParseXml(definitionXml);
                if (definitionNode != null)
                {
                    root.Add(new XElement("FilterDefinition", definitionNode));
                }
                else
                {
                    root.Add(new XElement("FilterDefinitionText", new XCData(definitionXml)));
                }
            }

            new XDocument(root).Save(filePath);
        }

        public static ViewFilterProfile LoadFromXml(string filePath)
        {
            if (!File.Exists(filePath)) throw new FileNotFoundException("필터 XML 파일을 찾을 수 없습니다.", filePath);

            XDocument doc = XDocument.Load(filePath);
            XElement root = doc.Root ?? throw new InvalidDataException("필터 XML 구조가 올바르지 않습니다.");

            string definitionXml = null;
            XElement filterDefinition = root.Element("FilterDefinition");
            if (filterDefinition != null)
            {
                XElement inner = filterDefinition.Elements().FirstOrDefault();
                if (inner != null)
                {
                    definitionXml = inner.ToString(XmlSaveOptions.DisableFormatting);
                }
            }

            if (string.IsNullOrWhiteSpace(definitionXml))
            {
                definitionXml = (string)root.Element("FilterDefinitionText");
            }

            var profile = new ViewFilterProfile
            {
                FilterName = (string)root.Element("FilterName") ?? string.Empty,
                CategoriesCsv = (string)root.Element("CategoriesCsv") ?? string.Empty,
                ParameterToken = (string)root.Element("ParameterToken") ?? string.Empty,
                RuleValue = (string)root.Element("RuleValue") ?? string.Empty,
                Operator = ParseOperator((string)root.Element("Operator")),
                StructureSummary = (string)root.Element("StructureSummary") ?? string.Empty,
                FilterDefinitionXml = definitionXml
            };

            ApplyPreviewFromDefinition(profile);
            return profile;
        }

        public static ViewFilterProfile ExtractProfileFromFilter(Document doc, ElementId filterId)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (filterId == null || filterId == ElementId.InvalidElementId) throw new ArgumentNullException(nameof(filterId));

            ParameterFilterElement filter = doc.GetElement(filterId) as ParameterFilterElement;
            if (filter == null)
            {
                throw new InvalidOperationException("현재 버전은 Parameter Filter XML 추출을 지원합니다.");
            }

            List<string> categoryTokens = filter.GetCategories()
                .Select(x => ToCategoryToken(doc, x))
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            ElementFilter elementFilter = filter.GetElementFilter();
            if (elementFilter == null)
            {
                throw new InvalidOperationException("필터 규칙이 없는 필터는 추출할 수 없습니다.");
            }

            XElement serialized = SerializeElementFilter(doc, elementFilter);
            RulePreview preview = FindFirstRulePreview(doc, elementFilter);

            return new ViewFilterProfile
            {
                FilterName = filter.Name,
                CategoriesCsv = string.Join(", ", categoryTokens),
                ParameterToken = preview?.ParameterToken ?? string.Empty,
                Operator = preview?.Operator ?? FilterRuleOperator.Equals,
                RuleValue = preview?.RuleValue ?? string.Empty,
                StructureSummary = BuildStructureSummary(doc, elementFilter),
                FilterDefinitionXml = serialized.ToString(XmlSaveOptions.DisableFormatting)
            };
        }

        public static ElementId EnsureFilter(Document doc, ViewFilterProfile profile, Action<string> log)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (profile == null) throw new ArgumentNullException(nameof(profile));
            if (!profile.IsConfigured()) throw new InvalidOperationException("필터 프로필이 완전하지 않습니다.");

            ICollection<ElementId> categoryIds = ResolveCategories(doc, profile.GetCategoryTokens(), log);
            if (categoryIds.Count == 0)
            {
                throw new InvalidOperationException("필터 카테고리를 하나도 찾지 못했습니다.");
            }

            ElementFilter elementFilter;
            if (!string.IsNullOrWhiteSpace(profile.FilterDefinitionXml))
            {
                XElement definition = TryParseXml(profile.FilterDefinitionXml);
                if (definition == null)
                {
                    throw new InvalidDataException("필터 정의 XML을 해석할 수 없습니다.");
                }

                var cache = new Dictionary<string, ParameterResolution>(StringComparer.OrdinalIgnoreCase);
                elementFilter = DeserializeElementFilter(doc, categoryIds, definition, cache);
            }
            else
            {
                ElementId parameterId;
                StorageType storageType;
                ResolveParameterIdentity(doc, categoryIds, profile.ParameterToken, out parameterId, out storageType);
                FilterRule rule = CreateRule(parameterId, storageType, profile.Operator, profile.RuleValue);
                elementFilter = new ElementParameterFilter(new List<FilterRule> { rule }, false);
            }

            ParameterFilterElement existing = new FilteredElementCollector(doc)
                .OfClass(typeof(ParameterFilterElement))
                .Cast<ParameterFilterElement>()
                .FirstOrDefault(x => string.Equals(x.Name, profile.FilterName, StringComparison.OrdinalIgnoreCase));

            if (existing != null)
            {
                try
                {
                    existing.SetCategories(categoryIds);
                    existing.SetElementFilter(elementFilter);
                    log?.Invoke("기존 필터 업데이트: " + existing.Name);
                    return existing.Id;
                }
                catch
                {
                    doc.Delete(existing.Id);
                }
            }

            ParameterFilterElement created = ParameterFilterElement.Create(doc, profile.FilterName, categoryIds, elementFilter);
            log?.Invoke("필터 생성: " + created.Name);
            return created.Id;
        }

        public static ElementFilter CreateElementFilter(Document doc, ViewFilterProfile profile, Action<string> log)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (profile == null) throw new ArgumentNullException(nameof(profile));
            if (!profile.IsConfigured()) throw new InvalidOperationException("?꾪꽣 ?꾨줈?꾩씠 ?꾩쟾?섏? ?딆뒿?덈떎.");

            ICollection<ElementId> categoryIds = ResolveCategories(doc, profile.GetCategoryTokens(), log);
            if (categoryIds.Count == 0)
            {
                throw new InvalidOperationException("?꾪꽣 移댄뀒怨좊━瑜??섎굹??李얠? 紐삵뻽?듬땲??");
            }

            if (!string.IsNullOrWhiteSpace(profile.FilterDefinitionXml))
            {
                XElement definition = TryParseXml(profile.FilterDefinitionXml);
                if (definition == null)
                {
                    throw new InvalidDataException("?꾪꽣 ?뺤쓽 XML???댁꽍?????놁뒿?덈떎.");
                }

                var cache = new Dictionary<string, ParameterResolution>(StringComparer.OrdinalIgnoreCase);
                return DeserializeElementFilter(doc, categoryIds, definition, cache);
            }

            ElementId parameterId;
            StorageType storageType;
            ResolveParameterIdentity(doc, categoryIds, profile.ParameterToken, out parameterId, out storageType);
            FilterRule rule = CreateRule(parameterId, storageType, profile.Operator, profile.RuleValue);
            return new ElementParameterFilter(new List<FilterRule> { rule }, false);
        }

        public static IList<ElementId> GetMatchingElementIds(Document doc, ViewFilterProfile profile, IEnumerable<ElementId> candidateIds, Action<string> log)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (profile == null) throw new ArgumentNullException(nameof(profile));

            ElementFilter filter = CreateElementFilter(doc, profile, log);
            HashSet<int> candidateIdSet = candidateIds != null
                ? new HashSet<int>(candidateIds.Where(x => x != null && x != ElementId.InvalidElementId).Select(x => x.IntegerValue))
                : null;

            ICollection<ElementId> categoryIds = ResolveCategories(doc, profile.GetCategoryTokens(), log);
            var results = new List<ElementId>();

            foreach (ElementId categoryId in categoryIds)
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .WherePasses(new ElementCategoryFilter(categoryId))
                    .WherePasses(filter);

                foreach (Element element in collector)
                {
                    if (element == null) continue;
                    if (candidateIdSet != null && !candidateIdSet.Contains(element.Id.IntegerValue)) continue;
                    results.Add(element.Id);
                }
            }

            return results
                .Distinct(new ElementIdComparer())
                .ToList();
        }

        public static void ApplyFilterToView(View view, ElementId filterId, bool enabled)
        {
            if (view == null) throw new ArgumentNullException(nameof(view));
            if (filterId == null || filterId == ElementId.InvalidElementId) return;

            ICollection<ElementId> filters = view.GetFilters();
            bool alreadyApplied = filters.Any(x => x.IntegerValue == filterId.IntegerValue);

            if (!alreadyApplied)
            {
                view.AddFilter(filterId);
                alreadyApplied = true;
            }

            if (!alreadyApplied)
            {
                return;
            }

            MethodInfo setIsFilterEnabled = typeof(View).GetMethod("SetIsFilterEnabled", new[] { typeof(ElementId), typeof(bool) });
            if (setIsFilterEnabled != null)
            {
                setIsFilterEnabled.Invoke(view, new object[] { filterId, enabled });
                return;
            }

            try
            {
                view.SetFilterVisibility(filterId, enabled);
            }
            catch
            {
                // older versions may not require extra work here
            }
        }

        private static void ApplyPreviewFromDefinition(ViewFilterProfile profile)
        {
            if (profile == null) return;
            if (string.IsNullOrWhiteSpace(profile.FilterDefinitionXml)) return;

            XElement definition = TryParseXml(profile.FilterDefinitionXml);
            if (definition == null) return;

            XElement firstRule = definition.DescendantsAndSelf().FirstOrDefault(x => string.Equals(x.Name.LocalName, "Rule", StringComparison.OrdinalIgnoreCase));
            if (firstRule == null) return;

            string parameterName = ((string)firstRule.Attribute("ParameterName") ?? string.Empty).Trim();
            string parameterToken = ((string)firstRule.Attribute("ParameterToken") ?? string.Empty).Trim();
            string value = ((string)firstRule.Attribute("Value") ?? string.Empty).Trim();
            string opText = ((string)firstRule.Attribute("Operator") ?? string.Empty).Trim();

            if (LooksLikeGuid(profile.ParameterToken) && !string.IsNullOrWhiteSpace(parameterName))
            {
                profile.ParameterToken = parameterName;
            }
            else if (string.IsNullOrWhiteSpace(profile.ParameterToken))
            {
                profile.ParameterToken = !string.IsNullOrWhiteSpace(parameterName) ? parameterName : parameterToken;
            }

            if (string.IsNullOrWhiteSpace(profile.RuleValue))
            {
                profile.RuleValue = value;
            }

            if (profile.Operator == default(FilterRuleOperator) && !string.IsNullOrWhiteSpace(opText))
            {
                profile.Operator = ParseOperator(opText);
            }
        }

        private static bool LooksLikeGuid(string text)
        {
            Guid value;
            return Guid.TryParse((text ?? string.Empty).Trim(), out value);
        }

        public static bool ViewHasPotentiallyVisibleModelElements(Document doc, View view)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_IOSModelGroups, true));

            foreach (Element element in collector)
            {
                if (element?.Category == null) continue;
                if (element.ViewSpecific) continue;
                if (element.Category.CategoryType != CategoryType.Model) continue;
                return true;
            }

            return false;
        }

        public static bool WouldViewBeEmptyWhenFilterHidden(Document doc, View view, ElementId filterId)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (view == null) throw new ArgumentNullException(nameof(view));
            if (filterId == null || filterId == ElementId.InvalidElementId) return false;

            ParameterFilterElement parameterFilter = doc.GetElement(filterId) as ParameterFilterElement;
            if (parameterFilter == null)
            {
                return !ViewHasPotentiallyVisibleModelElements(doc, view);
            }

            ElementFilter elementFilter = parameterFilter.GetElementFilter();
            if (elementFilter == null)
            {
                return !ViewHasPotentiallyVisibleModelElements(doc, view);
            }

            IList<ElementId> visibleModelElementIds = GetPotentiallyVisibleModelElementIds(doc, view, null);
            if (visibleModelElementIds.Count == 0)
            {
                return true;
            }

            IList<ElementId> matchedElementIds = GetPotentiallyVisibleModelElementIds(doc, view, elementFilter);
            if (matchedElementIds.Count == 0)
            {
                return false;
            }

            HashSet<int> matched = new HashSet<int>(matchedElementIds.Select(x => x.IntegerValue));
            foreach (ElementId id in visibleModelElementIds)
            {
                if (!matched.Contains(id.IntegerValue))
                {
                    return false;
                }
            }

            return true;
        }

        private static IList<ElementId> GetPotentiallyVisibleModelElementIds(Document doc, View view, ElementFilter extraFilter)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .WherePasses(new ElementCategoryFilter(BuiltInCategory.OST_IOSModelGroups, true));

            if (extraFilter != null)
            {
                collector = collector.WherePasses(extraFilter);
            }

            return collector
                .ToElements()
                .Where(x => x != null)
                .Where(x => x.Category != null)
                .Where(x => !x.ViewSpecific)
                .Where(x => x.Category.CategoryType == CategoryType.Model)
                .Select(x => x.Id)
                .Distinct(new ElementIdComparer())
                .ToList();
        }

        private static FilterRuleOperator ParseOperator(string value)
        {
            FilterRuleOperator parsed;
            if (Enum.TryParse(value, true, out parsed)) return parsed;
            return FilterRuleOperator.Equals;
        }

        private static string BuildSimpleDefinitionXml(ViewFilterProfile profile)
        {
            var rule = new XElement("Rule",
                new XAttribute("ParameterToken", profile.ParameterToken ?? string.Empty),
                new XAttribute("ParameterName", profile.ParameterToken ?? string.Empty),
                new XAttribute("Operator", profile.Operator.ToString()),
                new XAttribute("Value", profile.RuleValue ?? string.Empty),
                new XAttribute("StorageType", StorageType.String.ToString()),
                new XAttribute("CaseSensitive", false));

            var group = new XElement("ParameterGroup",
                new XAttribute("Inverted", false),
                rule);

            return group.ToString(XmlSaveOptions.DisableFormatting);
        }

        private static XElement SerializeElementFilter(Document doc, ElementFilter elementFilter)
        {
            LogicalAndFilter andFilter = elementFilter as LogicalAndFilter;
            if (andFilter != null)
            {
                return new XElement("Logical",
                    new XAttribute("Type", "And"),
                    andFilter.GetFilters().Select(x => SerializeElementFilter(doc, x)));
            }

            LogicalOrFilter orFilter = elementFilter as LogicalOrFilter;
            if (orFilter != null)
            {
                return new XElement("Logical",
                    new XAttribute("Type", "Or"),
                    orFilter.GetFilters().Select(x => SerializeElementFilter(doc, x)));
            }

            ElementParameterFilter parameterFilter = elementFilter as ElementParameterFilter;
            if (parameterFilter != null)
            {
                var node = new XElement("ParameterGroup",
                    new XAttribute("Inverted", parameterFilter.Inverted));

                foreach (FilterRule rule in parameterFilter.GetRules())
                {
                    node.Add(SerializeRule(doc, rule));
                }

                return node;
            }

            throw new InvalidOperationException("지원되지 않는 필터 구조입니다: " + elementFilter.GetType().Name);
        }

        private static XElement SerializeRule(Document doc, FilterRule rule)
        {
            FilterRule effectiveRule = UnwrapInverseRule(rule, out bool invertedRule);
            ParameterDescriptor descriptor = DescribeParameter(doc, rule.GetRuleParameter());
            string operatorName = GetRuleOperatorName(rule);
            string evaluatorName = TryGetEvaluatorName(rule);
            string value = ToRuleValue(rule);
            string ruleClassName = effectiveRule.GetType().Name;
            string storageType = GuessRuleStorageTypeName(rule, value);
            bool caseSensitive = TryGetBooleanProperty(effectiveRule, "CaseSensitive") ?? TryGetBooleanProperty(effectiveRule, "IsCaseSensitive") ?? false;
            double epsilon = TryGetDoubleProperty(effectiveRule, "Epsilon") ?? 0.0001;

            var node = new XElement("Rule",
                new XAttribute("ParameterToken", descriptor.ParameterToken ?? string.Empty),
                new XAttribute("Operator", operatorName),
                new XAttribute("Evaluator", evaluatorName ?? string.Empty),
                new XAttribute("RuleClass", ruleClassName ?? string.Empty),
                new XAttribute("InvertedRule", invertedRule),
                new XAttribute("Value", value ?? string.Empty),
                new XAttribute("StorageType", storageType),
                new XAttribute("CaseSensitive", caseSensitive),
                new XAttribute("Epsilon", epsilon.ToString("G17", CultureInfo.InvariantCulture)));

            if (!string.IsNullOrWhiteSpace(descriptor.ParameterName))
            {
                node.Add(new XAttribute("ParameterName", descriptor.ParameterName));
            }

            if (!string.IsNullOrWhiteSpace(descriptor.ParameterGuid))
            {
                node.Add(new XAttribute("ParameterGuid", descriptor.ParameterGuid));
            }

            if (!string.IsNullOrWhiteSpace(descriptor.ParameterKind))
            {
                node.Add(new XAttribute("ParameterKind", descriptor.ParameterKind));
            }

            return node;
        }

        private static ElementFilter DeserializeElementFilter(Document doc, ICollection<ElementId> categoryIds, XElement node, IDictionary<string, ParameterResolution> cache)
        {
            if (node == null) throw new ArgumentNullException(nameof(node));

            if (string.Equals(node.Name.LocalName, "Logical", StringComparison.OrdinalIgnoreCase))
            {
                string type = ((string)node.Attribute("Type") ?? "And").Trim();
                List<ElementFilter> filters = node.Elements().Select(x => DeserializeElementFilter(doc, categoryIds, x, cache)).ToList();
                if (filters.Count == 0)
                {
                    throw new InvalidDataException("Logical 필터 하위 조건이 없습니다.");
                }

                if (string.Equals(type, "Or", StringComparison.OrdinalIgnoreCase))
                {
                    return new LogicalOrFilter(filters);
                }

                return new LogicalAndFilter(filters);
            }

            if (string.Equals(node.Name.LocalName, "ParameterGroup", StringComparison.OrdinalIgnoreCase))
            {
                bool inverted = ParseBoolean(node.Attribute("Inverted"), false);
                List<FilterRule> rules = node.Elements("Rule").Select(x => DeserializeRule(doc, categoryIds, x, cache)).ToList();
                if (rules.Count == 0)
                {
                    throw new InvalidDataException("ParameterGroup 안에 Rule이 없습니다.");
                }

                return new ElementParameterFilter(rules, inverted);
            }

            if (string.Equals(node.Name.LocalName, "Rule", StringComparison.OrdinalIgnoreCase))
            {
                FilterRule singleRule = DeserializeRule(doc, categoryIds, node, cache);
                return new ElementParameterFilter(new List<FilterRule> { singleRule }, false);
            }

            throw new InvalidDataException("지원되지 않는 필터 노드입니다: " + node.Name.LocalName);
        }

        private static FilterRule DeserializeRule(Document doc, ICollection<ElementId> categoryIds, XElement node, IDictionary<string, ParameterResolution> cache)
        {
            string parameterToken = ((string)node.Attribute("ParameterToken") ?? string.Empty).Trim();
            string parameterName = ((string)node.Attribute("ParameterName") ?? string.Empty).Trim();
            string parameterGuid = ((string)node.Attribute("ParameterGuid") ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(parameterToken) && string.IsNullOrWhiteSpace(parameterName) && string.IsNullOrWhiteSpace(parameterGuid))
            {
                throw new InvalidDataException("Rule의 파라미터 식별 정보가 비어 있습니다.");
            }

            ParameterResolution resolution = ResolveParameterIdentityCached(doc, categoryIds, parameterToken, parameterName, parameterGuid, cache);
            FilterRuleOperator op = ParseOperator((string)node.Attribute("Operator"));
            string rawValue = (string)node.Attribute("Value") ?? string.Empty;

            return CreateRuleFromSerializedNode(resolution.ParameterId, resolution.StorageType, op, rawValue, node);
        }

        private static FilterRule CreateRuleFromSerializedNode(ElementId parameterId, StorageType storageType, FilterRuleOperator op, string rawValue, XElement ruleNode)
        {
            bool caseSensitive = ParseBoolean(ruleNode.Attribute("CaseSensitive"), false);
            double epsilon = ParseDouble(ruleNode.Attribute("Epsilon"), 0.0001);
            StorageType storageTypeHint = ParseStorageTypeName((string)ruleNode.Attribute("StorageType"), storageType);
            storageType = storageTypeHint;

            switch (op)
            {
                case FilterRuleOperator.HasValue:
                    return InvokeRuleFactoryRequired(new[] { "CreateHasValueParameterRule" }, parameterId, rawValue, caseSensitive, epsilon);
                case FilterRuleOperator.HasNoValue:
                    return CreateInverseRuleFallback(parameterId, rawValue, caseSensitive, epsilon,
                        new[] { "CreateHasNoValueParameterRule" },
                        () => InvokeRuleFactoryRequired(new[] { "CreateHasValueParameterRule" }, parameterId, rawValue, caseSensitive, epsilon));
                case FilterRuleOperator.Contains:
                    return InvokeRuleFactoryRequired(new[] { "CreateContainsRule" }, parameterId, rawValue, caseSensitive, epsilon);
                case FilterRuleOperator.NotContains:
                    return CreateInverseRuleFallback(parameterId, rawValue, caseSensitive, epsilon,
                        new[] { "CreateNotContainsRule" },
                        () => InvokeRuleFactoryRequired(new[] { "CreateContainsRule" }, parameterId, rawValue, caseSensitive, epsilon));
                case FilterRuleOperator.BeginsWith:
                    return InvokeRuleFactoryRequired(new[] { "CreateBeginsWithRule" }, parameterId, rawValue, caseSensitive, epsilon);
                case FilterRuleOperator.NotBeginsWith:
                    return CreateInverseRuleFallback(parameterId, rawValue, caseSensitive, epsilon,
                        new[] { "CreateNotBeginsWithRule" },
                        () => InvokeRuleFactoryRequired(new[] { "CreateBeginsWithRule" }, parameterId, rawValue, caseSensitive, epsilon));
                case FilterRuleOperator.EndsWith:
                    return InvokeRuleFactoryRequired(new[] { "CreateEndsWithRule" }, parameterId, rawValue, caseSensitive, epsilon);
                case FilterRuleOperator.NotEndsWith:
                    return CreateInverseRuleFallback(parameterId, rawValue, caseSensitive, epsilon,
                        new[] { "CreateNotEndsWithRule" },
                        () => InvokeRuleFactoryRequired(new[] { "CreateEndsWithRule" }, parameterId, rawValue, caseSensitive, epsilon));
                case FilterRuleOperator.NotEquals:
                    return CreateInverseRuleFallback(parameterId, rawValue, caseSensitive, epsilon,
                        new[] { "CreateNotEqualsRule" },
                        () => CreateRule(parameterId, storageType, FilterRuleOperator.Equals, rawValue));
                case FilterRuleOperator.Greater:
                    return InvokeRuleFactoryRequired(new[] { "CreateGreaterRule" }, parameterId, rawValue, caseSensitive, epsilon);
                case FilterRuleOperator.GreaterOrEqual:
                    return InvokeRuleFactoryRequired(new[] { "CreateGreaterOrEqualRule" }, parameterId, rawValue, caseSensitive, epsilon);
                case FilterRuleOperator.Less:
                    return InvokeRuleFactoryRequired(new[] { "CreateLessRule" }, parameterId, rawValue, caseSensitive, epsilon);
                case FilterRuleOperator.LessOrEqual:
                    return InvokeRuleFactoryRequired(new[] { "CreateLessOrEqualRule" }, parameterId, rawValue, caseSensitive, epsilon);
                case FilterRuleOperator.Equals:
                default:
                    return CreateRule(parameterId, storageType, FilterRuleOperator.Equals, rawValue);
            }
        }

        private static FilterRule InvokeRuleFactoryRequired(IEnumerable<string> methodNames, ElementId parameterId, string rawValue, bool caseSensitive, double epsilon)
        {
            foreach (string methodName in methodNames)
            {
                FilterRule rule = TryInvokeRuleFactory(methodName, parameterId, rawValue, caseSensitive, epsilon);
                if (rule != null) return rule;
            }

            throw new InvalidOperationException("지원되지 않는 필터 규칙 연산입니다: " + string.Join("/", methodNames));
        }
        private static FilterRule CreateInverseRuleFallback(ElementId parameterId, string rawValue, bool caseSensitive, double epsilon, IEnumerable<string> directMethodNames, Func<FilterRule> positiveFactory)
        {
            foreach (string methodName in directMethodNames)
            {
                FilterRule directRule = TryInvokeRuleFactory(methodName, parameterId, rawValue, caseSensitive, epsilon);
                if (directRule != null) return directRule;
            }

            FilterRule positiveRule = positiveFactory != null ? positiveFactory() : null;
            if (positiveRule == null)
            {
                throw new InvalidOperationException("역필터 규칙을 생성하지 못했습니다: " + string.Join("/", directMethodNames ?? Enumerable.Empty<string>()));
            }

            return new FilterInverseRule(positiveRule);
        }


        private static FilterRule TryInvokeRuleFactory(string methodName, ElementId parameterId, string rawValue, bool caseSensitive, double epsilon)
        {
            MethodInfo[] methods = typeof(ParameterFilterRuleFactory)
                .GetMethods(BindingFlags.Public | BindingFlags.Static)
                .Where(x => string.Equals(x.Name, methodName, StringComparison.Ordinal))
                .ToArray();

            foreach (MethodInfo method in methods)
            {
                ParameterInfo[] parameters = method.GetParameters();
                if (parameters.Length == 1 && parameters[0].ParameterType == typeof(ElementId))
                {
                    return method.Invoke(null, new object[] { parameterId }) as FilterRule;
                }

                if (parameters.Length == 2 && parameters[0].ParameterType == typeof(ElementId))
                {
                    if (parameters[1].ParameterType == typeof(string))
                    {
                        return method.Invoke(null, new object[] { parameterId, rawValue ?? string.Empty }) as FilterRule;
                    }

                    if (parameters[1].ParameterType == typeof(int))
                    {
                        int intValue;
                        if (!int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out intValue)) continue;
                        return method.Invoke(null, new object[] { parameterId, intValue }) as FilterRule;
                    }

                    if (parameters[1].ParameterType == typeof(ElementId))
                    {
                        int elementIdValue;
                        if (!int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out elementIdValue)) continue;
                        return method.Invoke(null, new object[] { parameterId, new ElementId(elementIdValue) }) as FilterRule;
                    }
                }

                if (parameters.Length == 3 && parameters[0].ParameterType == typeof(ElementId))
                {
                    if (parameters[1].ParameterType == typeof(string) && parameters[2].ParameterType == typeof(bool))
                    {
                        return method.Invoke(null, new object[] { parameterId, rawValue ?? string.Empty, caseSensitive }) as FilterRule;
                    }

                    if (parameters[1].ParameterType == typeof(double) && parameters[2].ParameterType == typeof(double))
                    {
                        double doubleValue;
                        if (!double.TryParse(rawValue, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out doubleValue)) continue;
                        return method.Invoke(null, new object[] { parameterId, doubleValue, epsilon }) as FilterRule;
                    }
                }
            }

            return null;
        }

        private static RulePreview FindFirstRulePreview(Document doc, ElementFilter elementFilter)
        {
            LogicalAndFilter andFilter = elementFilter as LogicalAndFilter;
            if (andFilter != null)
            {
                foreach (ElementFilter child in andFilter.GetFilters())
                {
                    RulePreview nested = FindFirstRulePreview(doc, child);
                    if (nested != null) return nested;
                }
                return null;
            }

            LogicalOrFilter orFilter = elementFilter as LogicalOrFilter;
            if (orFilter != null)
            {
                foreach (ElementFilter child in orFilter.GetFilters())
                {
                    RulePreview nested = FindFirstRulePreview(doc, child);
                    if (nested != null) return nested;
                }
                return null;
            }

            ElementParameterFilter parameterFilter = elementFilter as ElementParameterFilter;
            if (parameterFilter != null)
            {
                FilterRule firstRule = parameterFilter.GetRules().FirstOrDefault();
                if (firstRule == null) return null;

                return new RulePreview
                {
                    ParameterToken = ToParameterToken(doc, firstRule.GetRuleParameter()),
                    Operator = ParseOperator(GetRuleOperatorName(firstRule)),
                    RuleValue = ToRuleValue(firstRule)
                };
            }

            return null;
        }

        private static string BuildStructureSummary(Document doc, ElementFilter elementFilter)
        {
            LogicalAndFilter andFilter = elementFilter as LogicalAndFilter;
            if (andFilter != null)
            {
                return "AND(" + string.Join(", ", andFilter.GetFilters().Select(x => BuildStructureSummary(doc, x))) + ")";
            }

            LogicalOrFilter orFilter = elementFilter as LogicalOrFilter;
            if (orFilter != null)
            {
                return "OR(" + string.Join(", ", orFilter.GetFilters().Select(x => BuildStructureSummary(doc, x))) + ")";
            }

            ElementParameterFilter parameterFilter = elementFilter as ElementParameterFilter;
            if (parameterFilter != null)
            {
                List<string> parts = parameterFilter.GetRules().Select(x => BuildRuleSummary(doc, x)).ToList();
                string text = parts.Count == 1 ? parts[0] : "ALL(" + string.Join(", ", parts) + ")";
                if (parameterFilter.Inverted)
                {
                    return "NOT(" + text + ")";
                }
                return text;
            }

            return elementFilter.GetType().Name;
        }

        private static string BuildRuleSummary(Document doc, FilterRule rule)
        {
            string parameterToken = ToParameterToken(doc, rule.GetRuleParameter());
            string op = GetRuleOperatorName(rule);
            string value = ToRuleValue(rule);

            if (op == nameof(FilterRuleOperator.HasValue) || op == nameof(FilterRuleOperator.HasNoValue))
            {
                return parameterToken + " " + op;
            }

            return parameterToken + " " + op + " " + Quote(value);
        }

        private static string Quote(string value)
        {
            return "\"" + (value ?? string.Empty) + "\"";
        }

        private static string GetRuleOperatorName(FilterRule rule)
        {
            FilterRule effectiveRule = UnwrapInverseRule(rule, out bool inverted);
            string evaluatorName = TryGetEvaluatorName(effectiveRule);
            string ruleType = effectiveRule.GetType().Name;
            string source = (evaluatorName + "|" + ruleType).ToLowerInvariant();
            string baseOperator;

            if (source.Contains("hasnovalue")) baseOperator = nameof(FilterRuleOperator.HasNoValue);
            else if (source.Contains("hasvalue")) baseOperator = nameof(FilterRuleOperator.HasValue);
            else if (source.Contains("notcontains")) baseOperator = nameof(FilterRuleOperator.NotContains);
            else if (source.Contains("contains")) baseOperator = nameof(FilterRuleOperator.Contains);
            else if (source.Contains("notbegins") || source.Contains("notstarts")) baseOperator = nameof(FilterRuleOperator.NotBeginsWith);
            else if (source.Contains("begins") || source.Contains("starts")) baseOperator = nameof(FilterRuleOperator.BeginsWith);
            else if (source.Contains("notends")) baseOperator = nameof(FilterRuleOperator.NotEndsWith);
            else if (source.Contains("ends")) baseOperator = nameof(FilterRuleOperator.EndsWith);
            else if (source.Contains("greaterorequal") || source.Contains("greaterequal")) baseOperator = nameof(FilterRuleOperator.GreaterOrEqual);
            else if (source.Contains("lessorequal") || source.Contains("lessequal")) baseOperator = nameof(FilterRuleOperator.LessOrEqual);
            else if (source.Contains("greater")) baseOperator = nameof(FilterRuleOperator.Greater);
            else if (source.Contains("less")) baseOperator = nameof(FilterRuleOperator.Less);
            else if (source.Contains("notequal")) baseOperator = nameof(FilterRuleOperator.NotEquals);
            else baseOperator = nameof(FilterRuleOperator.Equals);

            return ApplyRuleInversion(baseOperator, inverted);
        }

        private static string TryGetEvaluatorName(FilterRule rule)
        {
            if (rule == null) return string.Empty;
            FilterRule effectiveRule = UnwrapInverseRule(rule, out _);

            MethodInfo getEvaluator = effectiveRule.GetType().GetMethod("GetEvaluator", BindingFlags.Public | BindingFlags.Instance);
            if (getEvaluator != null)
            {
                object evaluator = getEvaluator.Invoke(effectiveRule, null);
                if (evaluator != null) return evaluator.GetType().Name;
            }

            PropertyInfo evaluatorProperty = effectiveRule.GetType().GetProperty("Evaluator", BindingFlags.Public | BindingFlags.Instance);
            if (evaluatorProperty != null)
            {
                object evaluator = evaluatorProperty.GetValue(effectiveRule, null);
                if (evaluator != null) return evaluator.GetType().Name;
            }

            return string.Empty;
        }

        private static string GuessRuleStorageTypeName(FilterRule rule, string value)
        {
            FilterRule effectiveRule = UnwrapInverseRule(rule, out _);
            string ruleType = effectiveRule.GetType().Name.ToLowerInvariant();
            if (ruleType.Contains("double")) return StorageType.Double.ToString();
            if (ruleType.Contains("integer") || ruleType.Contains("int")) return StorageType.Integer.ToString();
            if (ruleType.Contains("elementid")) return StorageType.ElementId.ToString();
            if (ruleType.Contains("string")) return StorageType.String.ToString();

            int intValue;
            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out intValue)) return StorageType.Integer.ToString();

            double doubleValue;
            if (double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out doubleValue)) return StorageType.Double.ToString();

            return StorageType.String.ToString();
        }

        private static FilterRule UnwrapInverseRule(FilterRule rule, out bool inverted)
        {
            inverted = false;
            FilterRule current = rule;

            while (current is FilterInverseRule inverseRule)
            {
                MethodInfo getInnerRule = inverseRule.GetType().GetMethod("GetInnerRule", BindingFlags.Public | BindingFlags.Instance);
                if (getInnerRule == null)
                {
                    break;
                }

                FilterRule innerRule = getInnerRule.Invoke(inverseRule, null) as FilterRule;
                if (innerRule == null)
                {
                    break;
                }

                current = innerRule;
                inverted = !inverted;
            }

            return current ?? rule;
        }

        private static string ApplyRuleInversion(string operatorName, bool inverted)
        {
            if (!inverted) return operatorName ?? nameof(FilterRuleOperator.Equals);

            switch (operatorName ?? string.Empty)
            {
                case nameof(FilterRuleOperator.Equals):
                    return nameof(FilterRuleOperator.NotEquals);
                case nameof(FilterRuleOperator.NotEquals):
                    return nameof(FilterRuleOperator.Equals);
                case nameof(FilterRuleOperator.Contains):
                    return nameof(FilterRuleOperator.NotContains);
                case nameof(FilterRuleOperator.NotContains):
                    return nameof(FilterRuleOperator.Contains);
                case nameof(FilterRuleOperator.BeginsWith):
                    return nameof(FilterRuleOperator.NotBeginsWith);
                case nameof(FilterRuleOperator.NotBeginsWith):
                    return nameof(FilterRuleOperator.BeginsWith);
                case nameof(FilterRuleOperator.EndsWith):
                    return nameof(FilterRuleOperator.NotEndsWith);
                case nameof(FilterRuleOperator.NotEndsWith):
                    return nameof(FilterRuleOperator.EndsWith);
                case nameof(FilterRuleOperator.HasValue):
                    return nameof(FilterRuleOperator.HasNoValue);
                case nameof(FilterRuleOperator.HasNoValue):
                    return nameof(FilterRuleOperator.HasValue);
                default:
                    return operatorName ?? nameof(FilterRuleOperator.Equals);
            }
        }

        private static bool? TryGetBooleanProperty(object source, string propertyName)
        {
            if (source == null) return null;

            PropertyInfo property = source.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
            if (property == null) return null;
            if (property.PropertyType != typeof(bool)) return null;

            return (bool)property.GetValue(source, null);
        }

        private static double? TryGetDoubleProperty(object source, string propertyName)
        {
            if (source == null) return null;

            PropertyInfo property = source.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
            if (property == null) return null;
            if (property.PropertyType != typeof(double)) return null;

            return (double)property.GetValue(source, null);
        }

        private static StorageType ParseStorageTypeName(string value, StorageType defaultValue)
        {
            StorageType parsed;
            if (!string.IsNullOrWhiteSpace(value) && Enum.TryParse(value.Trim(), true, out parsed))
            {
                return parsed;
            }

            return defaultValue;
        }

        private static bool ParseBoolean(XAttribute attribute, bool defaultValue)
        {
            if (attribute == null) return defaultValue;
            bool parsed;
            return bool.TryParse(attribute.Value, out parsed) ? parsed : defaultValue;
        }

        private static double ParseDouble(XAttribute attribute, double defaultValue)
        {
            if (attribute == null) return defaultValue;
            double parsed;
            return double.TryParse(attribute.Value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out parsed)
                ? parsed
                : defaultValue;
        }

        private static XElement TryParseXml(string xml)
        {
            if (string.IsNullOrWhiteSpace(xml)) return null;
            try
            {
                return XElement.Parse(xml);
            }
            catch
            {
                return null;
            }
        }

        private static string ToCategoryToken(Document doc, ElementId categoryId)
        {
            if (categoryId == null || categoryId == ElementId.InvalidElementId) return string.Empty;

            if (categoryId.IntegerValue < 0 && Enum.IsDefined(typeof(BuiltInCategory), categoryId.IntegerValue))
            {
                return ((BuiltInCategory)categoryId.IntegerValue).ToString();
            }

            Category category = doc.Settings.Categories.Cast<Category>()
                .FirstOrDefault(x => x != null && x.Id.IntegerValue == categoryId.IntegerValue);

            return category?.Name ?? categoryId.IntegerValue.ToString(CultureInfo.InvariantCulture);
        }

        private static string ToParameterToken(Document doc, ElementId parameterId)
        {
            return DescribeParameter(doc, parameterId).ParameterToken;
        }

        private static ParameterDescriptor DescribeParameter(Document doc, ElementId parameterId)
        {
            if (parameterId == null || parameterId == ElementId.InvalidElementId)
            {
                throw new InvalidOperationException("필터 규칙의 파라미터를 확인할 수 없습니다.");
            }

            if (parameterId.IntegerValue < 0 && Enum.IsDefined(typeof(BuiltInParameter), parameterId.IntegerValue))
            {
                string bipName = ((BuiltInParameter)parameterId.IntegerValue).ToString();
                return new ParameterDescriptor
                {
                    ParameterToken = "BIP:" + bipName,
                    ParameterName = bipName,
                    ParameterGuid = string.Empty,
                    ParameterKind = "BuiltIn"
                };
            }

            Element parameterElement = doc.GetElement(parameterId);
            SharedParameterElement sharedParameter = parameterElement as SharedParameterElement;
            if (sharedParameter != null)
            {
                string sharedName = GetParameterElementName(sharedParameter);
                return new ParameterDescriptor
                {
                    ParameterToken = !string.IsNullOrWhiteSpace(sharedName) ? sharedName : sharedParameter.GuidValue.ToString(),
                    ParameterName = sharedName ?? string.Empty,
                    ParameterGuid = sharedParameter.GuidValue.ToString(),
                    ParameterKind = "Shared"
                };
            }

            ParameterElement parameterElementBase = parameterElement as ParameterElement;
            if (parameterElementBase != null)
            {
                string parameterName = GetParameterElementName(parameterElementBase);
                return new ParameterDescriptor
                {
                    ParameterToken = !string.IsNullOrWhiteSpace(parameterName) ? parameterName : parameterElementBase.Id.IntegerValue.ToString(CultureInfo.InvariantCulture),
                    ParameterName = parameterName ?? string.Empty,
                    ParameterGuid = string.Empty,
                    ParameterKind = "ParameterElement"
                };
            }

            string definitionName = TryGetParameterDefinitionName(parameterElement);
            return new ParameterDescriptor
            {
                ParameterToken = !string.IsNullOrWhiteSpace(definitionName) ? definitionName : parameterId.IntegerValue.ToString(CultureInfo.InvariantCulture),
                ParameterName = definitionName ?? string.Empty,
                ParameterGuid = string.Empty,
                ParameterKind = parameterElement != null ? parameterElement.GetType().Name : "Unknown"
            };
        }

        private static string GetParameterElementName(Element parameterElement)
        {
            string definitionName = TryGetParameterDefinitionName(parameterElement);
            if (!string.IsNullOrWhiteSpace(definitionName)) return definitionName;
            return parameterElement?.Name;
        }

        private static string TryGetParameterDefinitionName(Element parameterElement)
        {
            if (parameterElement == null) return null;

            MethodInfo getDefinition = parameterElement.GetType().GetMethod("GetDefinition", Type.EmptyTypes);
            if (getDefinition != null)
            {
                Definition def = getDefinition.Invoke(parameterElement, null) as Definition;
                if (def != null && !string.IsNullOrWhiteSpace(def.Name)) return def.Name;
            }

            PropertyInfo definitionProperty = parameterElement.GetType().GetProperty("Definition");
            if (definitionProperty != null)
            {
                Definition def = definitionProperty.GetValue(parameterElement, null) as Definition;
                if (def != null && !string.IsNullOrWhiteSpace(def.Name)) return def.Name;
            }

            return null;
        }

        private static string ToRuleValue(FilterRule rule)
        {
            if (rule == null) return string.Empty;
            FilterRule effectiveRule = UnwrapInverseRule(rule, out _);

            FilterStringRule stringRule = effectiveRule as FilterStringRule;
            if (stringRule != null)
            {
                return stringRule.RuleString ?? string.Empty;
            }

            foreach (string propertyName in new[] { "RuleValue", "Value", "RuleString" })
            {
                PropertyInfo property = effectiveRule.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
                if (property == null) continue;

                object value = property.GetValue(effectiveRule, null);
                string textValue = ConvertRuleValueToString(value);
                if (!string.IsNullOrWhiteSpace(textValue) || value != null)
                {
                    return textValue;
                }
            }

            foreach (string methodName in new[] { "GetRuleString", "GetRuleValue" })
            {
                MethodInfo method = effectiveRule.GetType().GetMethod(methodName, BindingFlags.Public | BindingFlags.Instance, null, Type.EmptyTypes, null);
                if (method == null) continue;

                object value = method.Invoke(effectiveRule, null);
                string textValue = ConvertRuleValueToString(value);
                if (!string.IsNullOrWhiteSpace(textValue) || value != null)
                {
                    return textValue;
                }
            }

            return string.Empty;
        }

        private static string ConvertRuleValueToString(object value)
        {
            if (value == null) return string.Empty;
            if (value is ElementId elementId)
            {
                return elementId.IntegerValue.ToString(CultureInfo.InvariantCulture);
            }

            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private static ICollection<ElementId> ResolveCategories(Document doc, IEnumerable<string> tokens, Action<string> log)
        {
            var result = new List<ElementId>();
            Categories categories = doc.Settings.Categories;

            foreach (string token in tokens)
            {
                if (string.IsNullOrWhiteSpace(token)) continue;

                BuiltInCategory bic;
                if (Enum.TryParse(token, true, out bic))
                {
                    try
                    {
                        Category cat = categories.get_Item(bic);
                        if (cat != null)
                        {
                            result.Add(cat.Id);
                            continue;
                        }
                    }
                    catch
                    {
                    }
                }

                Category byName = categories.Cast<Category>()
                    .FirstOrDefault(x => string.Equals(x.Name, token, StringComparison.OrdinalIgnoreCase));

                if (byName != null)
                {
                    result.Add(byName.Id);
                }
                else
                {
                    log?.Invoke("필터 카테고리 미해결: " + token);
                }
            }

            return result.Distinct(new ElementIdComparer()).ToList();
        }

        private static ParameterResolution ResolveParameterIdentityCached(Document doc, ICollection<ElementId> categoryIds, string parameterToken, string parameterName, string parameterGuid, IDictionary<string, ParameterResolution> cache)
        {
            string cacheKey = BuildParameterCacheKey(parameterToken, parameterName, parameterGuid);
            ParameterResolution cached;
            if (cache.TryGetValue(cacheKey, out cached)) return cached;

            ElementId parameterId;
            StorageType storageType;
            ResolveParameterIdentity(doc, categoryIds, parameterToken, parameterName, parameterGuid, out parameterId, out storageType);
            cached = new ParameterResolution { ParameterId = parameterId, StorageType = storageType };
            cache[cacheKey] = cached;
            return cached;
        }

        private static void ResolveParameterIdentity(Document doc, ICollection<ElementId> categoryIds, string parameterToken, out ElementId parameterId, out StorageType storageType)
        {
            ResolveParameterIdentity(doc, categoryIds, parameterToken, parameterToken, null, out parameterId, out storageType);
        }

        private static void ResolveParameterIdentity(Document doc, ICollection<ElementId> categoryIds, string parameterToken, string parameterName, string parameterGuid, out ElementId parameterId, out StorageType storageType)
        {
            parameterToken = (parameterToken ?? string.Empty).Trim();
            parameterName = (parameterName ?? string.Empty).Trim();
            parameterGuid = (parameterGuid ?? string.Empty).Trim();

            if (TryResolveBuiltInParameter(parameterToken, out parameterId))
            {
                storageType = GuessBuiltInParameterStorageType(doc, categoryIds, parameterId);
                return;
            }

            if (TryResolveBuiltInParameter(parameterName, out parameterId))
            {
                storageType = GuessBuiltInParameterStorageType(doc, categoryIds, parameterId);
                return;
            }

            Guid guidValue;
            if (!string.IsNullOrWhiteSpace(parameterGuid) && Guid.TryParse(parameterGuid, out guidValue))
            {
                SharedParameterElement byGuidAttr = new FilteredElementCollector(doc)
                    .OfClass(typeof(SharedParameterElement))
                    .Cast<SharedParameterElement>()
                    .FirstOrDefault(x => x.GuidValue == guidValue);

                if (byGuidAttr != null)
                {
                    parameterId = byGuidAttr.Id;
                    storageType = GuessParameterStorageTypeFromElements(doc, categoryIds, parameterId, parameterName);
                    return;
                }
            }

            if (Guid.TryParse(parameterToken, out guidValue))
            {
                SharedParameterElement byGuidToken = new FilteredElementCollector(doc)
                    .OfClass(typeof(SharedParameterElement))
                    .Cast<SharedParameterElement>()
                    .FirstOrDefault(x => x.GuidValue == guidValue);

                if (byGuidToken != null)
                {
                    parameterId = byGuidToken.Id;
                    storageType = GuessParameterStorageTypeFromElements(doc, categoryIds, parameterId, parameterName);
                    return;
                }
            }

            string preferredName = !string.IsNullOrWhiteSpace(parameterName) ? parameterName : parameterToken;
            if (!string.IsNullOrWhiteSpace(preferredName))
            {
                ParameterElement parameterElementByName = new FilteredElementCollector(doc)
                    .OfClass(typeof(ParameterElement))
                    .Cast<ParameterElement>()
                    .FirstOrDefault(x => string.Equals(GetParameterElementName(x), preferredName, StringComparison.OrdinalIgnoreCase));

                if (parameterElementByName != null)
                {
                    parameterId = parameterElementByName.Id;
                    storageType = GuessParameterStorageTypeFromElements(doc, categoryIds, parameterId, preferredName);
                    return;
                }
            }

            Parameter parameterByName = FindParameterOnSampleElements(doc, categoryIds, preferredName);
            if (parameterByName != null)
            {
                parameterId = parameterByName.Id;
                storageType = parameterByName.StorageType;
                return;
            }

            if (Guid.TryParse(parameterToken, out guidValue))
            {
                throw new InvalidOperationException("공유 파라미터 GUID를 찾지 못했습니다: " + parameterToken);
            }

            throw new InvalidOperationException("필터 파라미터를 찾지 못했습니다: " + preferredName);
        }

        private static string BuildParameterCacheKey(string parameterToken, string parameterName, string parameterGuid)
        {
            return (parameterToken ?? string.Empty).Trim() + "|" + (parameterName ?? string.Empty).Trim() + "|" + (parameterGuid ?? string.Empty).Trim();
        }

        private static Parameter FindParameterOnSampleElements(Document doc, ICollection<ElementId> categoryIds, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(parameterName)) return null;

            IEnumerable<Element> sampleElements = FindSampleElements(doc, categoryIds, 10);
            foreach (Element element in sampleElements)
            {
                if (element == null) continue;

                foreach (Parameter parameter in element.Parameters)
                {
                    if (parameter?.Definition == null) continue;
                    if (string.Equals(parameter.Definition.Name, parameterName, StringComparison.OrdinalIgnoreCase))
                    {
                        return parameter;
                    }
                }
            }

            return null;
        }

        private static IEnumerable<Element> FindSampleElements(Document doc, ICollection<ElementId> categoryIds, int maxPerCategory)
        {
            foreach (ElementId categoryId in categoryIds)
            {
                FilteredElementCollector collector = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .WherePasses(new ElementCategoryFilter(categoryId));

                int count = 0;
                foreach (Element element in collector)
                {
                    if (element == null) continue;
                    yield return element;
                    count++;
                    if (count >= maxPerCategory) break;
                }
            }
        }

        private static bool TryResolveBuiltInParameter(string token, out ElementId parameterId)
        {
            parameterId = ElementId.InvalidElementId;
            if (string.IsNullOrWhiteSpace(token)) return false;

            string work = token.Trim();
            if (work.StartsWith("BIP:", StringComparison.OrdinalIgnoreCase))
            {
                work = work.Substring(4).Trim();
            }

            BuiltInParameter bip;
            if (Enum.TryParse(work, true, out bip))
            {
                parameterId = new ElementId((int)bip);
                return true;
            }

            return false;
        }

        private static StorageType GuessBuiltInParameterStorageType(Document doc, ICollection<ElementId> categoryIds, ElementId parameterId)
        {
            Element sampleElement = FindSampleElement(doc, categoryIds);
            if (sampleElement == null) return StorageType.String;

            foreach (Parameter parameter in sampleElement.Parameters)
            {
                if (parameter == null) continue;
                if (parameter.Id.IntegerValue == parameterId.IntegerValue) return parameter.StorageType;
            }

            return StorageType.String;
        }

        private static StorageType GuessParameterStorageTypeFromElements(Document doc, ICollection<ElementId> categoryIds, ElementId parameterId, string parameterToken)
        {
            Element sampleElement = FindSampleElement(doc, categoryIds);
            if (sampleElement == null) return StorageType.String;

            foreach (Parameter parameter in sampleElement.Parameters)
            {
                if (parameter?.Definition == null) continue;
                if (parameter.Id.IntegerValue == parameterId.IntegerValue) return parameter.StorageType;
                if (string.Equals(parameter.Definition.Name, parameterToken, StringComparison.OrdinalIgnoreCase)) return parameter.StorageType;
            }

            return StorageType.String;
        }

        private static Element FindSampleElement(Document doc, ICollection<ElementId> categoryIds)
        {
            foreach (ElementId categoryId in categoryIds)
            {
                Element element = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .WherePasses(new ElementCategoryFilter(categoryId))
                    .FirstElement();

                if (element != null) return element;
            }

            return null;
        }

        private static FilterRule CreateRule(ElementId parameterId, StorageType storageType, FilterRuleOperator op, string rawValue)
        {
            switch (storageType)
            {
                case StorageType.Integer:
                    if (op == FilterRuleOperator.NotEquals) return CreateInverseRuleFallback(parameterId, rawValue, false, 0.0001, new[] { "CreateNotEqualsRule" }, () => CreateRule(parameterId, StorageType.Integer, FilterRuleOperator.Equals, rawValue));
                    if (op == FilterRuleOperator.Greater) return InvokeRuleFactoryRequired(new[] { "CreateGreaterRule" }, parameterId, rawValue, false, 0.0001);
                    if (op == FilterRuleOperator.GreaterOrEqual) return InvokeRuleFactoryRequired(new[] { "CreateGreaterOrEqualRule" }, parameterId, rawValue, false, 0.0001);
                    if (op == FilterRuleOperator.Less) return InvokeRuleFactoryRequired(new[] { "CreateLessRule" }, parameterId, rawValue, false, 0.0001);
                    if (op == FilterRuleOperator.LessOrEqual) return InvokeRuleFactoryRequired(new[] { "CreateLessOrEqualRule" }, parameterId, rawValue, false, 0.0001);

                    int intValue;
                    if (!int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out intValue))
                    {
                        throw new InvalidOperationException("정수형 필터 값으로 변환할 수 없습니다: " + rawValue);
                    }
                    return ParameterFilterRuleFactory.CreateEqualsRule(parameterId, intValue);

                case StorageType.Double:
                    if (op == FilterRuleOperator.NotEquals) return CreateInverseRuleFallback(parameterId, rawValue, false, 0.0001, new[] { "CreateNotEqualsRule" }, () => CreateRule(parameterId, StorageType.Double, FilterRuleOperator.Equals, rawValue));
                    if (op == FilterRuleOperator.Greater) return InvokeRuleFactoryRequired(new[] { "CreateGreaterRule" }, parameterId, rawValue, false, 0.0001);
                    if (op == FilterRuleOperator.GreaterOrEqual) return InvokeRuleFactoryRequired(new[] { "CreateGreaterOrEqualRule" }, parameterId, rawValue, false, 0.0001);
                    if (op == FilterRuleOperator.Less) return InvokeRuleFactoryRequired(new[] { "CreateLessRule" }, parameterId, rawValue, false, 0.0001);
                    if (op == FilterRuleOperator.LessOrEqual) return InvokeRuleFactoryRequired(new[] { "CreateLessOrEqualRule" }, parameterId, rawValue, false, 0.0001);

                    double doubleValue;
                    if (!double.TryParse(rawValue, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out doubleValue))
                    {
                        throw new InvalidOperationException("실수형 필터 값으로 변환할 수 없습니다: " + rawValue);
                    }
                    return ParameterFilterRuleFactory.CreateEqualsRule(parameterId, doubleValue, 0.0001);

                case StorageType.ElementId:
                    if (op == FilterRuleOperator.NotEquals) return CreateInverseRuleFallback(parameterId, rawValue, false, 0.0001, new[] { "CreateNotEqualsRule" }, () => CreateRule(parameterId, StorageType.ElementId, FilterRuleOperator.Equals, rawValue));

                    int elementIdValue;
                    if (!int.TryParse(rawValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out elementIdValue))
                    {
                        throw new InvalidOperationException("ElementId 필터 값으로 변환할 수 없습니다: " + rawValue);
                    }
                    return ParameterFilterRuleFactory.CreateEqualsRule(parameterId, new ElementId(elementIdValue));

                case StorageType.String:
                default:
                    switch (op)
                    {
                        case FilterRuleOperator.Contains:
                            return ParameterFilterRuleFactory.CreateContainsRule(parameterId, rawValue ?? string.Empty, false);
                        case FilterRuleOperator.NotContains:
                            return InvokeRuleFactoryRequired(new[] { "CreateNotContainsRule" }, parameterId, rawValue, false, 0.0001);
                        case FilterRuleOperator.BeginsWith:
                            return ParameterFilterRuleFactory.CreateBeginsWithRule(parameterId, rawValue ?? string.Empty, false);
                        case FilterRuleOperator.NotBeginsWith:
                            return InvokeRuleFactoryRequired(new[] { "CreateNotBeginsWithRule" }, parameterId, rawValue, false, 0.0001);
                        case FilterRuleOperator.EndsWith:
                            return ParameterFilterRuleFactory.CreateEndsWithRule(parameterId, rawValue ?? string.Empty, false);
                        case FilterRuleOperator.NotEndsWith:
                            return InvokeRuleFactoryRequired(new[] { "CreateNotEndsWithRule" }, parameterId, rawValue, false, 0.0001);
                        case FilterRuleOperator.NotEquals:
                            return CreateInverseRuleFallback(parameterId, rawValue, false, 0.0001, new[] { "CreateNotEqualsRule" }, () => CreateRule(parameterId, StorageType.String, FilterRuleOperator.Equals, rawValue));
                        case FilterRuleOperator.Greater:
                            return InvokeRuleFactoryRequired(new[] { "CreateGreaterRule" }, parameterId, rawValue, false, 0.0001);
                        case FilterRuleOperator.GreaterOrEqual:
                            return InvokeRuleFactoryRequired(new[] { "CreateGreaterOrEqualRule" }, parameterId, rawValue, false, 0.0001);
                        case FilterRuleOperator.Less:
                            return InvokeRuleFactoryRequired(new[] { "CreateLessRule" }, parameterId, rawValue, false, 0.0001);
                        case FilterRuleOperator.LessOrEqual:
                            return InvokeRuleFactoryRequired(new[] { "CreateLessOrEqualRule" }, parameterId, rawValue, false, 0.0001);
                        case FilterRuleOperator.HasValue:
                            return InvokeRuleFactoryRequired(new[] { "CreateHasValueParameterRule" }, parameterId, rawValue, false, 0.0001);
                        case FilterRuleOperator.HasNoValue:
                            return InvokeRuleFactoryRequired(new[] { "CreateHasNoValueParameterRule" }, parameterId, rawValue, false, 0.0001);
                        default:
                            return ParameterFilterRuleFactory.CreateEqualsRule(parameterId, rawValue ?? string.Empty, false);
                    }
            }
        }

        private sealed class ElementIdComparer : IEqualityComparer<ElementId>
        {
            public bool Equals(ElementId x, ElementId y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(x, null) || ReferenceEquals(y, null)) return false;
                return x.IntegerValue == y.IntegerValue;
            }

            public int GetHashCode(ElementId obj)
            {
                return obj?.IntegerValue ?? 0;
            }
        }

        private sealed class ParameterDescriptor
        {
            public string ParameterToken { get; set; }
            public string ParameterName { get; set; }
            public string ParameterGuid { get; set; }
            public string ParameterKind { get; set; }
        }

        private sealed class ParameterResolution
        {
            public ElementId ParameterId { get; set; }
            public StorageType StorageType { get; set; }
        }

        private sealed class RulePreview
        {
            public string ParameterToken { get; set; }
            public FilterRuleOperator Operator { get; set; }
            public string RuleValue { get; set; }
        }
    }
}
