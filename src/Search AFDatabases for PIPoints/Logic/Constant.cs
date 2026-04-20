namespace Search_AFDatabases_for_PIPoints.Logic;

internal static class Constant
{
    public static string DashLine { get; } = new string('-', 80);
    public static string Pad { get; } = new string(' ', 3);

    public static string[] PointAttributes => new string[]
    {
        PICommonPointAttributes.Descriptor,
        PICommonPointAttributes.EngineeringUnits,
        PICommonPointAttributes.CreationDate
    };

}
