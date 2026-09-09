using System.Reflection;
namespace Lokad.OoxPdf.Tests;
internal sealed record TestResult(string Group, string Name, string Outcome, long Milliseconds, string Detail);
internal sealed class TestSkippedException(string message) : Exception(message);
internal static class TestCatalogValidator
{
    public static IReadOnlyList<string> Validate(IReadOnlyList<TestCase> catalog)
    {
        var problems = new List<string>();
        var registered = new HashSet<MethodInfo>(catalog.Select(testCase => testCase.Action.Method));
        Assembly assembly = typeof(TestCatalogValidator).Assembly;
        foreach (Type type in assembly.GetTypes())
        {
            if (!type.IsClass || !type.Name.EndsWith("Tests", StringComparison.Ordinal))
            {
                continue;
            }
            foreach (MethodInfo method in type.GetMethods(BindingFlags.Public | BindingFlags.Static))
            {
                if (method.GetParameters().Length == 0 && method.ReturnType == typeof(void))
                {
                    if (!registered.Contains(method))
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
