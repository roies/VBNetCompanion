namespace CSharpLib;

public class GreeterService
{
	public string FormatMessage(string name)
	{
		return $"Hello, {name} from C#";
	}

	public int Add(int left, int right)
	{
		return left + right;
	}
}
