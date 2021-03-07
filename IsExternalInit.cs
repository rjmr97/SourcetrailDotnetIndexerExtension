// Without this class we get the following error:
// "Predefined type 'System.Runtime.CompilerServices.IsExternalInit' is not defined or imported"
// This solution is taken from here: https://stackoverflow.com/a/64749403/6651287
namespace System.Runtime.CompilerServices
{
    internal static class IsExternalInit { }
}
