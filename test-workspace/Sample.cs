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
        _ = message.Length;
    }
}
