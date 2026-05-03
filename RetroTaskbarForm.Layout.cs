namespace win9xplorer
{
    internal sealed partial class RetroTaskbarForm
    {
        private ToolbarHostApplier GetToolbarHostApplier()
        {
            toolbarHostApplier ??= new ToolbarHostApplier(new ToolbarHostContext(
                ContentPanel: contentPanel,
                TaskButtonsHostPanel: taskButtonsHostPanel,
                AddressRowHostPanel: addressRowHostPanel,
                Toolbars: toolbarsById,
                ScaffoldControls: new Dictionary<ScaffoldSlot, Control>
                {
                    [ScaffoldSlot.TaskButtonsHostPanel] = taskButtonsHostPanel,
                    [ScaffoldSlot.NotificationAreaPanel] = notificationAreaPanel,
                    [ScaffoldSlot.TaskButtonsSeparatorPanel] = taskButtonsSeparatorPanel,
                    [ScaffoldSlot.ToolbarsSeparatorPanel] = toolbarsSeparatorPanel,
                    [ScaffoldSlot.SpotifyToolbarSeparatorPanel] = spotifyToolbarSeparatorPanel,
                    [ScaffoldSlot.StartToolbarSeparatorPanel] = startToolbarSeparatorPanel,
                    [ScaffoldSlot.StartHostPanel] = startHostPanel,
                    [ScaffoldSlot.AddressToolbarSeparatorPanel] = addressToolbarSeparatorPanel,
                    [ScaffoldSlot.AddressRowGripPanel] = addressRowGripPanel
                }));
            return toolbarHostApplier;
        }

        private bool ShouldUseDedicatedAddressToolbarRow()
        {
            return taskbarRows > 1 && showAddressToolbar;
        }

        private int GetDedicatedAddressToolbarHeight()
        {
            return TaskbarButtonHeight;
        }

        private int GetAddressInputHostHeight()
        {
            var textHeight = Math.Max(18, addressToolbarComboBox.Height);
            return Math.Max(textHeight + 4, TaskbarButtonHeight - 8);
        }

        private void ApplyToolbarLayout(bool refreshBounds = true)
        {
            var useDedicatedAddressToolbarRow = ShouldUseDedicatedAddressToolbarRow();
            var plan = ToolbarLayoutService.BuildPlan(
                showQuickLaunchToolbar,
                showAddressToolbar,
                showVolumeToolbar,
                showSpotifyToolbar,
                addressToolbarBeforeQuickLaunch,
                useDedicatedAddressToolbarRow);

            quickLaunchPanel.Visible = showQuickLaunchToolbar;
            addressToolbarPanel.Visible = showAddressToolbar;
            volumeToolbarPanel.Visible = showVolumeToolbar;
            var layoutState = new ToolbarLayoutState(
                ShowQuickLaunch: showQuickLaunchToolbar,
                ShowAddressInLeftRail: plan.ShowAddressInLeftToolbarRail,
                UseDedicatedAddressRow: plan.UseDedicatedAddressToolbarRow,
                ShowVolume: showVolumeToolbar,
                ShowSpotify: showSpotifyToolbar);
            foreach (var relation in toolbarGripRelations)
            {
                relation.Grip.SetToolbarVisible(relation.IsVisible(layoutState));
            }
            foreach (var relation in toolbarSeparatorRelations)
            {
                relation.Separator.Visible = relation.IsVisible(layoutState);
            }
            spotifyToolbarPanel.Visible = showSpotifyToolbar;
            addressRowSeparatorPanel.Visible = plan.UseDedicatedAddressToolbarRow;
            addressRowHostPanel.Visible = plan.UseDedicatedAddressToolbarRow;

            var hostApplier = GetToolbarHostApplier();

            taskButtonsHostPanel.SuspendLayout();
            hostApplier.RemoveFloatingToolbarControls();

            addressToolbarPanel.Margin = Padding.Empty;
            addressToolbarPanel.Dock = DockStyle.Left;
            addressToolbarPanel.Width = 260;
            addressToolbarPanel.Height = GetDedicatedAddressToolbarHeight();
            volumeToolbarPanel.Margin = Padding.Empty;
            volumeToolbarPanel.Dock = DockStyle.Left;
            volumeToolbarPanel.Width = 140;
            volumeToolbarPanel.Height = GetDedicatedAddressToolbarHeight();
            addressInputHostPanel.Height = GetAddressInputHostHeight();
            CenterAddressTextBoxVertically();

            hostApplier.ApplyAddressRow(plan);
            if (plan.UseDedicatedAddressToolbarRow)
            {
                addressRowHostPanel.Height = GetDedicatedAddressToolbarHeight() + AddressRowSeparatorHeight;
                for (var i = 0; i < addressRowHostPanel.Controls.Count; i++)
                {
                    var child = addressRowHostPanel.Controls[i];
                    var childId =
                        ReferenceEquals(child, volumeToolbarGripPanel) ? nameof(volumeToolbarGripPanel) :
                        ReferenceEquals(child, volumeToolbarPanel) ? nameof(volumeToolbarPanel) :
                        ReferenceEquals(child, addressToolbarSeparatorPanel) ? nameof(addressToolbarSeparatorPanel) :
                        ReferenceEquals(child, addressRowGripPanel) ? nameof(addressRowGripPanel) :
                        ReferenceEquals(child, addressToolbarPanel) ? nameof(addressToolbarPanel) :
                        ReferenceEquals(child, addressRowSeparatorPanel) ? nameof(addressRowSeparatorPanel) :
                        child.GetType().Name;
                    Console.WriteLine(
                        $"[AddressRowOrder] Index={i}, Control={childId}, Dock={child.Dock}, Visible={child.Visible}, Left={child.Left}, Width={child.Width}, Right={child.Right}");
                }
            }

            hostApplier.ApplyMainRow(plan);

            taskButtonsHostPanel.ResumeLayout();

            ResizeQuickLaunchPanel();
            UpdateSpotifyToolbarLayout();
            if (refreshBounds && IsHandleCreated)
            {
                SetTaskbarBounds();
                RefreshTaskbar();
            }
        }
    }
}
