using System.Text.RegularExpressions;

namespace CodesysProtocols.Model.Type;

public class WagoTime : IEquatable<WagoTime>
{
    private int _timeValue;
    public int TimeValue
    {
        get { return _timeValue; }
        set
        {
            if (value >= 0)
            {
                this._timeValue = value;
            }
        }
    }

    private string _timeBase;
    public string TimeBase
    {
        get { return _timeBase; }
        set
        {
            if (WagoTimeBaseType.TypesDict.ContainsValue(value))
            {
                this._timeBase = value;
            }
        }
    }

    public WagoTime()
    {
        TimeBase = WagoTimeBaseType.TypesDict[WagoTimeBaseTypes.Millisecond.ToString()];
    }

    public void SetTimeValue(string input)
    {
        var exceptionInfo = $"Invalid input string format: {input}";

        if (input.Length < 4)
            throw new ArgumentException(exceptionInfo);
        if(!string.Equals(input.Substring(0,2), @"t#", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException(exceptionInfo);

        var subStr = input.Substring(2);
        var matches = Regex.Matches(subStr, @"([0-9]+)");

        if (matches.Count != 1)
            throw new ArgumentException(exceptionInfo);
        var timeValue = matches[0].Value;
        var timeBase = subStr.Substring(timeValue.Length);
        
        if((timeValue.Length + timeBase.Length) != subStr.Length)
            throw new ArgumentException(exceptionInfo);
        this.TimeBase = timeBase;
        if (this.TimeBase != timeBase)
            throw new ArgumentException(exceptionInfo);
        this.TimeValue = int.Parse(timeValue);
    }

    public override string ToString()
    {
        return $"t#{this.TimeValue}{this.TimeBase}";
    }

    public bool Equals(WagoTime other)
    {
        if (other == null)
            return false;
        return this.TimeValue == other.TimeValue
            && this.TimeBase == other.TimeBase;
    }
}
