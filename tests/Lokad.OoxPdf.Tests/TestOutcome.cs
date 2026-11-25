using System.Reflection;
namespace Lokad.OoxPdf.Tests;
internal sealed record TestResult(string Group, string Name, string Outcome, long Milliseconds, string Detail);
internal sealed class TestSkippedException(string message) : Exception(message);
internal static class TestCatalogValidator
{
    public static IReadOnlyList<string> Validate(IReadOnlyList<TestCase> catalog, IReadOnlySet<string>? exactNames,
        Assembly? coverageAssembly)
    {
        var problems = new List<string>();
        var registered = new Dictionary<string, MethodInfo>(catalog.Count, StringComparer.Ordinal);
        foreach (var testCase in catalog)
        {
            MethodInfo method = testCase.Action.Method;
            if (!registered.TryAdd(method.Name, method))
                problems.Add($"Duplicate test registration '{method.Name}'.");
        }
        if (exactNames is not null)
        {
            foreach (string name in exactNames)
                if (!registered.ContainsKey(name))
                    problems.Add($"Unknown exact test name '{name}'.");
        }
        if (problems.Count != 0 || coverageAssembly is null)
            return problems;

        foreach (Type type in coverageAssembly.GetTypes())
        {
            if (!type.IsClass || !type.Name.EndsWith("Tests", StringComparison.Ordinal))
            {
                continue;
            }
            foreach (MethodInfo method in type.GetMethods(BindingFlags.Public | BindingFlags.Static))
            {
                if (method.GetParameters().Length == 0 && method.ReturnType == typeof(void))
                {
                    if (!registered.TryGetValue(method.Name, out var registeredMethod) || registeredMethod != method)
                    {
                        problems.Add("Unregistered test " + type.Name + "." + method.Name + " would never run.");
                    }
                }
                else if (!method.IsSpecialName)
                {
                    problems.Add("Suspicious member " + type.Name + "." + method.Name + " is public static but not a runnable test.");
                }
            }
        }
        return problems;
    }
}
