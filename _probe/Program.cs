using Mscc.GenerativeAI;
using System.Reflection;
// Show exact parameter names for GenerateContent(List<IPart> ...)
foreach (var m in typeof(GenerativeModel).GetMethods(BindingFlags.Public | BindingFlags.Instance)
    .Where(m => m.Name == "GenerateContent" && m.GetParameters().Length > 0))
{
    var ps = m.GetParameters();
    if (ps[0].ParameterType.Name.Contains("List") || ps[0].ParameterType.Name.Contains("IList"))
        Console.WriteLine($"  ({string.Join(", ", ps.Select(p => p.ParameterType.Name + " " + p.Name))})");
}
