TODO
## A. Taskbar 佈局與互動（高優先）
- [ ] 多列 taskbar 真正生效（2/3 rows 時 running programs 能換行分佈，不只是容器變高）
- [ ] 鎖定/解鎖狀態更貼 Win9x：
  - lock 時上沿純亮條
  - unlock 時可拖曳 grip（視覺與 hit-test 完全一致）
- [ ] Start/Task/Quick Launch 三區高度與 baseline 做「像素級對齊」配置（避免每次手調 padding）
- [ ] Button 間距、上下距離改為集中配置（theme token），避免散落常數
## B. Start Menu 完整 Win9x 行為（高優先）
- [ ] Start menu keyboard navigation 全面化：
  - 左右鍵跨層 submenu 導航
  - accelerator key（例如按 P 跳 Programs）
  - Home/End/PageUp/PageDown
- [ ] 搜尋輸入框改成可編輯且不影響 hover（目前是 readonly + hook）
- [ ] 搜尋結果與原 menu 項目融合排序（前綴匹配 > 完整匹配 > apps）
- [ ] Programs / Windows Apps 右鍵 context menu 行為一致化（包括 submenu node）
## C. Tray 相容性（高優先）
- [ ] 針對「Windows Security / 某些特殊 icon」建立可配置 fallback 規則表
- [ ] Tray callback 封裝成獨立 service（減少 form 內事件邏輯耦合）
- [ ] Tooltip 行為與 Win9x 更接近（延遲、持續時間、標題/內容優先級）
- [ ] Tray icon refresh/placement 健壯化（多螢幕、DPI 切換、taskbar rows 變化）
## D. 視覺主題系統（高優先）
- [ ] 把目前顏色 + bevel + button style 抽成 Theme Profile
- [ ] 預設主題：
  - Win95
  - Win98 Classic
  - Win98 Plus-like
- [ ] Context menu renderer 與 taskbar renderer 統一使用同一份 theme token
- [ ] 新增「一鍵還原 Win98 預設外觀」按鈕（含顏色/字型/bevel）
## E. Options 體驗（中優先）
- [ ] Options 分頁化（Appearance / Behavior / Start Menu / Tray）
- [ ] 顏色設定加入「輸入 HEX」與「複製/貼上主題」
- [ ] Apply 後即時預覽，Cancel 回滾（目前 Apply 是 commit）
- [ ] 增加「匯出/匯入設定」JSON（便於分享主題）
## F. Quick Launch（中優先）
- [ ] 支援拖放順序重排（目前主要是資料夾排序）
- [ ] 支援分隔符捷徑（像 Win9x 的視覺分組）
- [ ] 右鍵選單補齊：
  - Rename
  - Move left/right
  - Open containing folder
- [ ] 啟動失敗時 UI 提示（目前主要是 debug log）
## G. 程式清單來源（中優先）
- [ ] StartApps（UWP）索引改為快取服務 + 背景刷新（避免偶發慢/重複）
- [ ] Windows Apps icon 取圖雙路 fallback（PIDL + Shell COM）
- [ ] 移除重複項策略可配置（同名不同來源顯示規則）
## H. 穩定性與維護（高價值）
- [ ] 拆分 `RetroTaskbarForm`：
  - `StartMenuController`
  - `TrayController`
  - `QuickLaunchController`
  - `ThemeController`
- [ ] 建立 UI 回歸測試清單（手動 smoke + 自動最小腳本）
- [ ] 所有 timing/race（start toggle、active handle）集中到狀態機
- [ ] 加入診斷面板（顯示 active handle、tray icon counts、cache state）
## I. RetroBar 對齊項目（建議優先）
- [ ] 菜單繪製細節（字距、icon margin、separator 凹陷深度）做逐項對照
- [ ] Task button 按下/激活/閃爍狀態機對齊
- [ ] Notification area 排版與 hit-test 行為對齊
- [ ] DPI 行為（100/125/150）逐像素驗證