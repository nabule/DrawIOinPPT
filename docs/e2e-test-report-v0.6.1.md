# Full E2E Test Report (v0.6.1-stable)

本轮完整实机测试已覆盖：

- 注册并加载 PowerPoint COM Add-in
- URL 模式新建 Draw.io 图形
- 保存 PPT、关闭、重新打开
- 对同一图形再次编辑并回写
- 验证 `Shape.Tags` / `AlternativeText` / `Presentation.CustomXMLParts` 持久化可继续工作

| Check | Result | Detail |
| --- | --- | --- |
| ReleasePackage | PASS | C:\Users\nabul\Desktop\greensoft\code\drawioppt\artifacts\\releases\\v0.6.1-stable\package |
| InstallRelease | PASS | C:\Users\nabul\AppData\Local\Temp\DrawioPpt\\installed-497fa794f7c749aa903291453c2f6319 |
| PowerPointAddInLoad | PASS | Connect=True |
| InstalledUrlSmoke | PASS | C:\Users\nabul\AppData\Local\Temp\DrawioPpt\\installed-497fa794f7c749aa903291453c2f6319\scripts\\url-editor-smoke.ps1 |
| PowerPointUrlE2E | PASS | ExitCode=0 |
| DesktopExporter | PASS | C:\\Program Files\\draw.io\\draw.io.exe |
