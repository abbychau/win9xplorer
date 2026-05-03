namespace win9xplorer
{
    internal enum ToolbarPart
    {
        Panel,
        Grip
    }

    internal enum ScaffoldSlot
    {
        ToolbarsSeparatorPanel,
        SpotifyToolbarSeparatorPanel,
        StartToolbarSeparatorPanel,
        StartHostPanel,
        AddressToolbarSeparatorPanel,
        AddressRowGripPanel,
        TaskButtonsHostPanel,
        NotificationAreaPanel,
        TaskButtonsSeparatorPanel
    }

    internal readonly record struct ToolbarSlot(ToolbarId ToolbarId, ToolbarPart Part);

    internal readonly record struct LayoutSlot(ToolbarSlot? Toolbar, ScaffoldSlot? Scaffold)
    {
        public static LayoutSlot ForToolbar(ToolbarId toolbarId, ToolbarPart part) => new(new ToolbarSlot(toolbarId, part), null);
        public static LayoutSlot ForScaffold(ScaffoldSlot scaffold) => new(null, scaffold);
    }

    internal sealed record ToolbarLayoutPlan(
        bool UseDedicatedAddressToolbarRow,
        bool ShowAddressInLeftToolbarRail,
        bool ShowVolumeInDedicatedAddressRow,
        IReadOnlyList<LayoutSlot> MainRowSlots,
        IReadOnlyList<LayoutSlot> AddressRowSlots);

    internal static class ToolbarLayoutService
    {
        public static ToolbarLayoutPlan BuildPlan(
            bool showQuickLaunchToolbar,
            bool showAddressToolbar,
            bool showVolumeToolbar,
            bool showSpotifyToolbar,
            bool addressToolbarBeforeQuickLaunch,
            bool useDedicatedAddressToolbarRow)
        {
            var showAddressInLeftToolbarRail = showAddressToolbar && !useDedicatedAddressToolbarRow;
            var showVolumeInDedicatedAddressRow = useDedicatedAddressToolbarRow && showVolumeToolbar;
            var mainSlots = new List<LayoutSlot>();

            if (showQuickLaunchToolbar && showAddressInLeftToolbarRail)
            {
                if (addressToolbarBeforeQuickLaunch)
                {
                    mainSlots.Add(LayoutSlot.ForToolbar(ToolbarId.QuickLaunch, ToolbarPart.Panel));
                    mainSlots.Add(LayoutSlot.ForToolbar(ToolbarId.QuickLaunch, ToolbarPart.Grip));
                    mainSlots.Add(LayoutSlot.ForScaffold(ScaffoldSlot.ToolbarsSeparatorPanel));
                    mainSlots.Add(LayoutSlot.ForToolbar(ToolbarId.Address, ToolbarPart.Panel));
                    mainSlots.Add(LayoutSlot.ForToolbar(ToolbarId.Address, ToolbarPart.Grip));
                }
                else
                {
                    mainSlots.Add(LayoutSlot.ForToolbar(ToolbarId.Address, ToolbarPart.Panel));
                    mainSlots.Add(LayoutSlot.ForToolbar(ToolbarId.Address, ToolbarPart.Grip));
                    mainSlots.Add(LayoutSlot.ForScaffold(ScaffoldSlot.ToolbarsSeparatorPanel));
                    mainSlots.Add(LayoutSlot.ForToolbar(ToolbarId.QuickLaunch, ToolbarPart.Panel));
                    mainSlots.Add(LayoutSlot.ForToolbar(ToolbarId.QuickLaunch, ToolbarPart.Grip));
                }
            }
            else
            {
                if (showQuickLaunchToolbar)
                {
                    mainSlots.Add(LayoutSlot.ForToolbar(ToolbarId.QuickLaunch, ToolbarPart.Panel));
                    mainSlots.Add(LayoutSlot.ForToolbar(ToolbarId.QuickLaunch, ToolbarPart.Grip));
                }

                if (showAddressInLeftToolbarRail)
                {
                    mainSlots.Add(LayoutSlot.ForToolbar(ToolbarId.Address, ToolbarPart.Panel));
                    mainSlots.Add(LayoutSlot.ForToolbar(ToolbarId.Address, ToolbarPart.Grip));
                }
            }

            if (!showVolumeInDedicatedAddressRow && showVolumeToolbar)
            {
                mainSlots.Add(LayoutSlot.ForToolbar(ToolbarId.Volume, ToolbarPart.Grip));
                mainSlots.Add(LayoutSlot.ForToolbar(ToolbarId.Volume, ToolbarPart.Panel));
            }

            if (showSpotifyToolbar)
            {
                if (mainSlots.Count > 0)
                {
                    mainSlots.Add(LayoutSlot.ForScaffold(ScaffoldSlot.SpotifyToolbarSeparatorPanel));
                }
                mainSlots.Add(LayoutSlot.ForToolbar(ToolbarId.Spotify, ToolbarPart.Panel));
                mainSlots.Add(LayoutSlot.ForToolbar(ToolbarId.Spotify, ToolbarPart.Grip));
            }

            if (mainSlots.Count > 0)
            {
                mainSlots.Add(LayoutSlot.ForScaffold(ScaffoldSlot.StartToolbarSeparatorPanel));
            }
            mainSlots.Add(LayoutSlot.ForScaffold(ScaffoldSlot.StartHostPanel));

            var addressRowSlots = new List<LayoutSlot>
            {
                LayoutSlot.ForToolbar(ToolbarId.Address, ToolbarPart.Panel),
                LayoutSlot.ForScaffold(ScaffoldSlot.AddressRowGripPanel),
                LayoutSlot.ForScaffold(ScaffoldSlot.AddressToolbarSeparatorPanel)
            };
            if (showVolumeInDedicatedAddressRow)
            {
                addressRowSlots.Add(LayoutSlot.ForToolbar(ToolbarId.Volume, ToolbarPart.Panel));
                addressRowSlots.Add(LayoutSlot.ForToolbar(ToolbarId.Volume, ToolbarPart.Grip));
            }

            return new ToolbarLayoutPlan(
                UseDedicatedAddressToolbarRow: useDedicatedAddressToolbarRow,
                ShowAddressInLeftToolbarRail: showAddressInLeftToolbarRail,
                ShowVolumeInDedicatedAddressRow: showVolumeInDedicatedAddressRow,
                MainRowSlots: mainSlots,
                AddressRowSlots: addressRowSlots);
        }
    }
}
