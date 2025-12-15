namespace OfficeHelperOpenXml.Utils
{
    /// <summary>
    /// 单位转换工具类
    /// </summary>
    public static class UnitConverter
    {
        private const double EMU_PER_INCH = 914400.0;
        private const double CM_PER_INCH = 2.54;
        private const double POINTS_PER_INCH = 72.0;
        private const double EMU_PER_CM = EMU_PER_INCH / CM_PER_INCH;
        private const double POINTS_PER_CM = POINTS_PER_INCH / CM_PER_INCH;

        public static double EmuToCm(long emu) => emu / EMU_PER_CM;
        public static long CmToEmu(double cm) => (long)(cm * EMU_PER_CM);
        public static double EmuToInches(long emu) => emu / EMU_PER_INCH;
        public static long InchesToEmu(double inches) => (long)(inches * EMU_PER_INCH);
        public static double PointsToCm(double points) => points / POINTS_PER_CM;
        public static double CmToPoints(double cm) => cm * POINTS_PER_CM;
        public static double PointsToInches(double points) => points / POINTS_PER_INCH;
        public static double InchesToPoints(double inches) => inches * POINTS_PER_INCH;
        public static double InchesToCm(double inches) => inches * CM_PER_INCH;
        public static double CmToInches(double cm) => cm / CM_PER_INCH;
        public static double EmuToPoints(long emu) => emu / (EMU_PER_INCH / POINTS_PER_INCH);
        public static long PointsToEmu(double points) => (long)(points * (EMU_PER_INCH / POINTS_PER_INCH));

        public static string EmuBoxToCmString(long x, long y, long width, long height)
        {
            return $"{EmuToCm(x):F2},{EmuToCm(y):F2},{EmuToCm(width):F2},{EmuToCm(height):F2}";
        }

        public static string PointsBoxToCmString(double x, double y, double width, double height)
        {
            return $"{PointsToCm(x):F2},{PointsToCm(y):F2},{PointsToCm(width):F2},{PointsToCm(height):F2}";
        }
    }
}
