# Word 复杂图形拖动性能优化设计

## 目标

修复 Word 插件插入复杂 Draw.io 图形后，移动、缩放或重新定位图片时出现的明显卡顿。优化必须保留完整 Draw.io 源数据、现有编辑能力和旧文档兼容能力，并交付为本机可安装验证的 `v1.0.8 Release x64` 版本。

## 已确认事实

- `v1.0.6` 已删除 Word 宿主每 500ms 读取选区的定时器，但 `WindowSelectionChange` 回调仍会同步读取图片 `AlternativeText`。
- `WordPictureMetadataService.Save` 当前把包含完整 `DrawioXml` 的 envelope 同时写入 `Document.CustomXMLParts` 和图片 `AlternativeText`。
- `WordPictureSelectionReader.Read` 会在选区事件中读取并反序列化完整 `AlternativeText`；反序列化过程包含 XML 解析、Base64 解码和 GZip 解压。
- 真实 Word 红灯使用 1,200 个图元构造约 271,361 字符的 Draw.io XML，得到约 75,997 字符的 `AlternativeText`。完整源数据已同时存在于 `Document.CustomXMLParts`。
- 当前 Release DLL 的只读基准显示，复杂 envelope 的同步识别约为每次 8 至 16ms，尚未计入 Word COM 和界面重绘；轻量 envelope 的成本接近可忽略。
- Word 的 `WindowSelectionChange` 只保证在选区发生变化时触发，当前代码和测试尚未量化一次真实鼠标拖动期间的事件次数。

## 设计决策

采用“文档级完整存储 + 图片轻量引用”方案：

- `Document.CustomXMLParts` 是 Word Draw.io 源数据的唯一完整主存储。
- 图片 `AlternativeText` 只保存可识别的轻量 envelope，保留 `diagramId`、名称、编辑模式、编辑目标、sidecar path 和更新时间，但 `DrawioXml` 必须为空。
- 不按图形大小设置分流阈值。所有新保存或成功迁移的 Word 受管图片使用同一种轻量格式，避免复杂度临界点和不一致行为。
- 不通过选区事件缓存或节流掩盖问题。选区读取仍保持同步和事件驱动，但只处理轻量元数据。

## 数据流

### 新建、绑定和编辑后保存

1. 生成或更新包含完整 `DrawioXml` 的 envelope。
2. 调用 `DocumentDiagramStore.Upsert`，把完整 envelope 写入活动 Word 文档的 `CustomXMLParts`。
3. 仅当 `Upsert` 返回有效部件 ID 时，将图片 `AlternativeText` 写成不含 `DrawioXml` 的轻量 envelope。
4. 如果文档级写入失败，将完整 envelope 保留或写回图片 `AlternativeText`，避免丢失源数据，并记录可诊断日志。
5. 图片 `Title`、尺寸、位置、环绕方式、锁定宽高比和 SVG 内容保持现有行为。

### 选区变化

1. `SelectionMonitor` 继续接收 `WindowSelectionChange`。
2. `WordPictureSelectionReader` 读取图片基本信息和轻量 `AlternativeText`。
3. 只解析 `diagramId` 并判断是否为受管图片，不读取 `CustomXMLParts`，也不解压完整 Draw.io XML。
4. `AddInHost` 继续使用选区上下文更新功能区和“自动打开”判断。

### 显式编辑和刷新

1. 从图片轻量 envelope 取得 `diagramId`。
2. 通过 `DocumentDiagramStore.TryRead` 从 `CustomXMLParts` 读取完整 envelope。
3. 打开编辑器或刷新 SVG。
4. 保存结果时重新执行“先完整主存储、后轻量引用”的顺序。

### 旧文档迁移

1. 如果图片 `AlternativeText` 中存在旧版完整 envelope，仍按现有格式正常反序列化。
2. 如果 `CustomXMLParts` 中已有同一 `diagramId` 的完整 envelope，以文档级数据为准，并把图片改写成轻量引用。
3. 如果文档级数据缺失，以旧版图片 envelope 恢复完整数据。
4. 只有补写 `CustomXMLParts` 成功后，才将旧版图片元数据瘦身。
5. 补写失败时保留旧版完整 `AlternativeText`，本次仍可编辑，不进行破坏性迁移。

## 组件变更

### `WordPictureMetadataService`

- 将常规保存语义改为生成并写入轻量图片 envelope。
- 增加明确的完整回退写入入口，仅用于文档级存储失败时的数据保护。
- 轻量 envelope 必须复制业务字段，不得修改调用方持有的完整 envelope。
- `TryRead` 同时兼容旧版完整 envelope 和新版轻量 envelope。

### `AddInHost`

- `SaveManagedEnvelope` 检查 `DocumentDiagramStore.Upsert` 的返回结果。
- 文档级写入成功时保存轻量引用，失败时保存完整回退。
- `TryReadManagedEnvelope` 从文档级存储读到完整数据后，只回写轻量引用，禁止重新膨胀 `AlternativeText`。
- 保留旧版回退恢复和孤儿清理逻辑；图片中的轻量 `diagramId` 仍用于引用扫描。

### 测试脚本

- 将 `scripts/word-complex-metadata-e2e.ps1` 纳入回归范围并扩展为真实 Word 复杂元数据测试。
- 更新直接从 `AlternativeText.DrawioXml` 取完整数据的 Word URL E2E，改为按 `diagramId` 从 `Document.CustomXMLParts` 验证。
- 保留 PowerPoint 现有完整 `AlternativeText` 策略，本次不修改其代码和测试。

## 失败处理

- 文档级存储失败时，正确性优先于性能：保留完整图片回退，不丢失 Draw.io XML。
- 轻量 envelope 构造失败时，不修改原图片元数据。
- 读取到轻量引用但找不到对应 `CustomXMLParts` 时，返回“无法恢复完整源数据”，不得用空 XML 覆盖图片或文档。
- 迁移和保存继续使用现有受保护执行入口记录异常；用户显式操作失败时显示可理解的错误信息。
- 不在本次改动中吞并相邻重构，也不修改核心序列化格式版本。

## 兼容性边界

- 整份 `.docx` 保存、复制到其他机器和重新打开时，完整 `CustomXMLParts` 会随文档保存，编辑能力保持不变。
- 旧版仅在图片 `AlternativeText` 中保存完整数据的文档可自动迁移。
- 用户已确认接受：单独复制一张 Word 图片到另一份文档时，目标文档不一定带有对应 `CustomXMLParts`，因此不再保证能从 `AlternativeText` 单独恢复完整 Draw.io XML。
- SVG 根元素中的 `content` 数据继续保留，但本次不新增从 Word 内部 SVG 自动恢复源 XML 的能力。
- 不修改 PowerPoint 的 `Tags + CustomXMLParts + AlternativeText` 策略。

## 测试与验收

### TDD 红灯

- 使用当前实现运行复杂元数据 E2E，确认：
  - 完整 XML 已写入 `CustomXMLParts`；
  - `AlternativeText` 大于 2,048 字符；
  - 图片 envelope 中仍含完整 `DrawioXml`；
  - 测试按预期失败。

### 自动化绿灯

- 1,200 图元场景中：
  - `AlternativeText` 不超过 2,048 字符；
  - 图片 envelope 的 `DrawioXml` 为空；
  - 图片和文档级 envelope 的 `diagramId` 一致；
  - 文档级 `DrawioXml` 与输入逐字符一致。
- 保存、关闭、重新打开 Word 文档后，仍可从 `CustomXMLParts` 读取完整 XML。
- 旧版完整图片 envelope 在文档级写入成功后自动变为轻量引用。
- 模拟文档级写入失败时，图片仍保留完整 XML。
- Word URL 新建、保存、重开编辑和二次回写 E2E 全部通过。
- Word 选区事件、无定时轮询约束、Release 构建和完整 Office E2E 全部通过。

### 性能与真实 Word 验收

- 轻量 envelope 的选区识别基准平均耗时不超过 2ms，且复杂图与简单图之间不再随 Draw.io XML 大小增长。
- 在安装后的真实 Word 中插入复杂浮动图形，连续移动、缩放和重新定位，与同尺寸普通图片进行对照。
- 拖动期间图片位置应连续跟随指针，不应再出现由插件元数据读取造成的明显周期性停顿。
- 若禁用插件后同一 SVG 仍有等量卡顿，应单独记录为 Word 原生 SVG 或浮动布局开销，不把它误报为本次修复失败。

## 文档、版本与本地安装

- 更新 `docs/architecture.md`、`docs/architecture.en.md`。
- 更新 `docs/word-addin.md`、`docs/word-addin.en.md`。
- 更新 `docs/user-guide.md`、`docs/user-guide.en.md` 和回归清单。
- 新增 `v1.0.8` 中英文发布说明、E2E 测试报告和发布证据。
- 将复杂元数据测试复制进发布包，并由 `scripts/full-e2e-test.ps1` 串行执行。
- 生成 `v1.0.8 Release x64` 发布包，保持现有 COM 类标识和程序集身份稳定。
- 使用当前用户注册脚本安装本地版本，重启 Word 后核对加载路径、`LoadBehavior` 和真实功能。

## 非目标

- 不改变 SVG、PNG 或其他图片格式选择。
- 不简化 Draw.io 图形内容，也不栅格化复杂 SVG。
- 不修改 Word 图片环绕方式和排版策略。
- 不增加选区事件节流、异步 COM 调用或后台线程。
- 不修改 PowerPoint 元数据策略。
- 不实现单图片跨文档复制后的完整源数据自动恢复。
