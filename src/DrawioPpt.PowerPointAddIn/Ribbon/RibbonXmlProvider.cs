namespace DrawioPpt.PowerPointAddIn.Ribbon
{
    public static class RibbonXmlProvider
    {
        public static string GetCustomUi()
        {
            return
@"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='Greensoft.DrawioPpt.Tab' label='Draw.io'>
        <group id='Greensoft.DrawioPpt.Group' label='Draw.io 图形'>
          <button id='Greensoft.DrawioPpt.New' label='新建图形' size='large' imageMso='InsertSmartArtGraphic' onAction='OnNewDiagram' />
          <button id='Greensoft.DrawioPpt.Edit' label='编辑选中图形' size='large' imageMso='DiagramCycle' onAction='OnEditSelectedDiagram' getEnabled='GetEditEnabled' />
          <button id='Greensoft.DrawioPpt.Settings' label='设置' imageMso='FileProperties' onAction='OnOpenSettings' />
          <labelControl id='Greensoft.DrawioPpt.Status' getLabel='GetSelectionStatus' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }
    }
}

