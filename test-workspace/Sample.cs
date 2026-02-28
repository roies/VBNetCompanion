namespace SampleWorkspace;

public class CSharpGreeter
{
    public string GetMessage(string name)
    {
        return $"Hello from C#: {name}";
    }

    public void Run()
    {
        var message = GetMessage("Developer");
        int localValue = message.Length;
        _ = localValue.ToString();
    }
}
