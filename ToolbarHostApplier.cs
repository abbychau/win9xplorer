using System.Windows.Forms;

namespace win9xplorer
{
    internal sealed record ToolbarHostContext(
        Panel ContentPanel,
        Panel TaskButtonsHostPanel,
        Panel AddressRowHostPanel,
        IReadOnlyDictionary<ToolbarId, ToolbarComponents> Toolbars,
        IReadOnlyDictionary<ScaffoldSlot, Control> ScaffoldControls);

    internal sealed class ToolbarHostApplier
    {
        private readonly ToolbarHostContext context;

        public ToolbarHostApplier(ToolbarHostContext context)
        {
            this.context = context;
        }

        public void RemoveFloatingToolbarControls()
        {
            RemoveFromAllHosts(ResolveToolbarControl(ToolbarId.Address, ToolbarPart.Panel));
            RemoveFromAllHosts(ResolveToolbarControl(ToolbarId.Volume, ToolbarPart.Panel));
            RemoveFromAllHosts(ResolveToolbarControl(ToolbarId.Volume, ToolbarPart.Grip));
        }

        public void ApplyAddressRow(ToolbarLayoutPlan plan)
        {
            if (!plan.UseDedicatedAddressToolbarRow)
            {
                return;
            }

            if (plan.ShowVolumeInDedicatedAddressRow)
            {
                context.AddressRowHostPanel.Controls.Add(ResolveToolbarControl(ToolbarId.Volume, ToolbarPart.Grip));
                context.AddressRowHostPanel.Controls.Add(ResolveToolbarControl(ToolbarId.Volume, ToolbarPart.Panel));
                ResolveToolbar(ToolbarId.Volume).Panel.Dock = DockStyle.Left;
            }
            context.AddressRowHostPanel.Controls.Add(ResolveToolbarControl(ToolbarId.Address, ToolbarPart.Panel));
            context.AddressRowHostPanel.Controls.Add(GetScaffoldControl(ScaffoldSlot.AddressRowGripPanel));
            ResolveToolbar(ToolbarId.Address).Panel.Dock = DockStyle.Fill;

            for (var i = 0; i < plan.AddressRowSlots.Count; i++)
            {
                if (TryResolveLayoutSlot(plan.AddressRowSlots[i], out var control) &&
                    context.AddressRowHostPanel.Controls.Contains(control))
                {
                    context.AddressRowHostPanel.Controls.SetChildIndex(control, i);
                }
            }
        }

        public void ApplyMainRow(ToolbarLayoutPlan plan)
        {
            context.ContentPanel.SuspendLayout();
            context.ContentPanel.Controls.Clear();
            context.ContentPanel.Controls.Add(GetScaffoldControl(ScaffoldSlot.TaskButtonsHostPanel));
            context.ContentPanel.Controls.Add(GetScaffoldControl(ScaffoldSlot.NotificationAreaPanel));
            context.ContentPanel.Controls.Add(GetScaffoldControl(ScaffoldSlot.TaskButtonsSeparatorPanel));

            foreach (var slot in plan.MainRowSlots)
            {
                if (!TryResolveLayoutSlot(slot, out var control))
                {
                    continue;
                }

                context.ContentPanel.Controls.Add(control);
                if (slot.Toolbar is ToolbarSlot toolbarSlot &&
                    toolbarSlot.ToolbarId == ToolbarId.Volume &&
                    toolbarSlot.Part == ToolbarPart.Panel &&
                    context.ContentPanel.Controls.Contains(ResolveToolbarControl(ToolbarId.Volume, ToolbarPart.Grip)))
                {
                    context.ContentPanel.Controls.SetChildIndex(
                        ResolveToolbarControl(ToolbarId.Volume, ToolbarPart.Grip),
                        context.ContentPanel.Controls.GetChildIndex(ResolveToolbarControl(ToolbarId.Volume, ToolbarPart.Panel)));
                }
            }
            context.ContentPanel.ResumeLayout();
        }

        private bool TryResolveLayoutSlot(LayoutSlot slot, out Control control)
        {
            if (slot.Toolbar is ToolbarSlot toolbarSlot)
            {
                control = ResolveToolbarControl(toolbarSlot.ToolbarId, toolbarSlot.Part);
                return true;
            }

            if (slot.Scaffold is ScaffoldSlot scaffoldSlot &&
                context.ScaffoldControls.TryGetValue(scaffoldSlot, out control!))
            {
                return true;
            }

            control = context.TaskButtonsHostPanel;
            return false;
        }

        private ToolbarComponents ResolveToolbar(ToolbarId toolbarId) => context.Toolbars[toolbarId];

        private Control ResolveToolbarControl(ToolbarId toolbarId, ToolbarPart part)
            => part == ToolbarPart.Grip ? ResolveToolbar(toolbarId).Grip : ResolveToolbar(toolbarId).Panel;

        private Control GetScaffoldControl(ScaffoldSlot slot) => context.ScaffoldControls[slot];

        private void RemoveFromAllHosts(Control control)
        {
            RemoveIfPresent(context.TaskButtonsHostPanel, control);
            RemoveIfPresent(context.AddressRowHostPanel, control);
        }

        private static void RemoveIfPresent(Control parent, Control child)
        {
            if (parent.Controls.Contains(child))
            {
                parent.Controls.Remove(child);
            }
        }
    }
}
