namespace DrawioPpt.WordAddIn.Ribbon
{
    public static class RibbonXmlProvider
    {
        public static string GetCustomUi()
        {
            return
@"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab id='Greensoft.DrawioWord.Tab' label='Draw.io'>
        <group id='Greensoft.DrawioWord.CreateGroup' label='创建'>
          <button id='Greensoft.DrawioWord.New' label='新建' size='large' getImage='GetButtonImage' onAction='OnNewDiagram' screentip='新建 Draw.io 图形' supertip='在当前光标位置插入一个新的 Draw.io 图形，并立即进入编辑。' />
        </group>
        <group id='Greensoft.DrawioWord.CurrentGroup' label='当前图形'>
          <button id='Greensoft.DrawioWord.Edit' getLabel='GetEditButtonLabel' size='large' getImage='GetButtonImage' onAction='OnEditSelectedDiagram' getEnabled='GetEditEnabled' screentip='重新编辑当前图形' supertip='选中图片后点击此按钮。只有点击时才读取 Draw.io 元数据，选中和拖动过程不执行插件处理。' />
          <button id='Greensoft.DrawioWord.Refresh' label='刷新' getImage='GetButtonImage' onAction='OnRefreshSelectedDiagram' getEnabled='GetManagedPictureEnabled' screentip='刷新当前图形' supertip='选中图片后点击此按钮，从绑定的 Draw.io 内容重新生成并替换当前图形。' />
          <button id='Greensoft.DrawioWord.Bind' label='绑定' getImage='GetButtonImage' onAction='OnBindSelectedPicture' getEnabled='GetBindEnabled' screentip='绑定选中图片' supertip='选中图片后点击此按钮，把普通图片纳入 Draw.io 管理。' />
          <button id='Greensoft.DrawioWord.Clear' label='清除绑定' getImage='GetButtonImage' onAction='OnClearSelectedPicture' getEnabled='GetManagedPictureEnabled' screentip='清除当前绑定' supertip='选中图片后点击此按钮，移除 Draw.io 绑定信息但保留图片。' />
        </group>
        <group id='Greensoft.DrawioWord.WorkflowGroup' label='工作流'>
          <button id='Greensoft.DrawioWord.Settings' label='设置' getImage='GetButtonImage' onAction='OnOpenSettings' screentip='插件设置' supertip='配置桌面编辑器和 URL 模式等选项。' />
        </group>
        <group id='Greensoft.DrawioWord.InfoGroup' label='信息'>
          <labelControl id='Greensoft.DrawioWord.Status' getLabel='GetSelectionStatus' />
          <labelControl id='Greensoft.DrawioWord.Detail' getLabel='GetSelectionDetailSummary' />
          <labelControl id='Greensoft.DrawioWord.Mode' getLabel='GetEditorModeSummary' />
          <labelControl id='Greensoft.DrawioWord.Version' getLabel='GetVersionSummary' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }
    }
}
