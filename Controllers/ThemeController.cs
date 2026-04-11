namespace win9xplorer
{
    internal sealed class ThemeController
    {
        internal sealed record ThemeProfile(
            string Name,
            string ButtonStyle,
            int BevelSize,
            string FontName,
            float FontSize,
            Color FontColor,
            Color BaseColor,
            Color LightColor,
            Color DarkColor,
            Color MenuSelectionColor);

        public static readonly ThemeProfile Win95 = new(
            "Win95",
            "Win98",
            1,
            "MS Sans Serif",
            8.25f,
            Color.Black,
            Color.FromArgb(192, 192, 192),
            Color.White,
            Color.FromArgb(128, 128, 128),
            Color.FromArgb(0, 0, 128));

        public static readonly ThemeProfile Win98Classic = new(
            "Win98 Classic",
            "Win98",
            2,
            "MS Sans Serif",
            8.25f,
            Color.Black,
            Color.FromArgb(192, 192, 192),
            Color.White,
            Color.FromArgb(128, 128, 128),
            Color.FromArgb(0, 0, 128));

        public static readonly ThemeProfile Win98PlusLike = new(
            "Win98 Plus-like",
            "Win98",
            2,
            "MS Sans Serif",
            8.25f,
            Color.FromArgb(10, 10, 10),
            Color.FromArgb(172, 183, 200),
            Color.FromArgb(238, 243, 250),
            Color.FromArgb(88, 100, 124),
            Color.FromArgb(20, 60, 160));

        private static readonly ThemeProfile[] Profiles =
        {
            Win95,
            Win98Classic,
            Win98PlusLike
        };

        public bool TryGetProfile(string? name, out ThemeProfile profile)
        {
            profile = Profiles.FirstOrDefault(p => p.Name.Equals(name, StringComparison.OrdinalIgnoreCase))!;
            return profile != null;
        }
    }
}
