namespace win9xplorer
{
    internal sealed class StartMenuController
    {
        private readonly TaskbarInteractionStateMachine stateMachine;
        private readonly Func<bool> isMenuVisible;
        private readonly Action openMenu;
        private readonly Action closeMenu;

        public StartMenuController(
            TaskbarInteractionStateMachine stateMachine,
            Func<bool> isMenuVisible,
            Action openMenu,
            Action closeMenu)
        {
            this.stateMachine = stateMachine;
            this.isMenuVisible = isMenuVisible;
            this.openMenu = openMenu;
            this.closeMenu = closeMenu;
        }

        public void OnStartButtonMouseDown()
        {
            stateMachine.RecordStartButtonMouseDown(isMenuVisible());
        }

        public void OnStartButtonClick()
        {
            if (stateMachine.WasStartMenuVisibleOnMouseDown())
            {
                closeMenu();
                stateMachine.ClearStartButtonMouseState();
                return;
            }

            openMenu();
            stateMachine.ClearStartButtonMouseState();
        }

        public void OnStartMenuClosed()
        {
            stateMachine.RecordStartMenuClosed(DateTime.UtcNow);
        }
    }
}
