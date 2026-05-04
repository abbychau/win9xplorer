namespace win9xplorer
{
    internal sealed class TaskbarInteractionStateMachine
    {
        private bool startMenuVisibleOnMouseDown;
        private IntPtr activeWindowHandle = IntPtr.Zero;
        private IntPtr foregroundWindowBeforeTaskClick = IntPtr.Zero;

        public TaskbarInteractionStateMachine(TimeSpan _) { }

        public IntPtr ActiveWindowHandle => activeWindowHandle;

        public IntPtr ForegroundWindowBeforeTaskClick => foregroundWindowBeforeTaskClick;

        public void RecordStartButtonMouseDown(bool menuVisible)
        {
            startMenuVisibleOnMouseDown = menuVisible;
        }

        public bool WasStartMenuVisibleOnMouseDown() => startMenuVisibleOnMouseDown;

        public void RecordStartMenuClosed(DateTime _)
        {
            startMenuVisibleOnMouseDown = false;
        }

        public void ClearStartButtonMouseState()
        {
            startMenuVisibleOnMouseDown = false;
        }

        public void RecordTaskButtonMouseDown(IntPtr foregroundHandle)
        {
            foregroundWindowBeforeTaskClick = foregroundHandle;
        }

        public bool ShouldMinimizeTaskWindow(IntPtr handle)
        {
            return handle == foregroundWindowBeforeTaskClick || handle == activeWindowHandle;
        }

        public void SetActiveWindowHandle(IntPtr handle)
        {
            activeWindowHandle = handle;
        }

        public void ClearActiveWindowHandleIfMatches(IntPtr handle)
        {
            if (activeWindowHandle == handle)
            {
                activeWindowHandle = IntPtr.Zero;
            }
        }

        public void ClearTaskButtonMouseState()
        {
            foregroundWindowBeforeTaskClick = IntPtr.Zero;
        }
    }
}
