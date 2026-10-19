# Smart Bundle Downloader GUI v1.12

## v1.12 — 修复动作类型与关键帧模板同步

- **升级主界面视觉**：统一采用浅色工作区、卡片式布局、层级更清晰的标签页、现代化控件颜色与更宽松的间距；下载、资源表和导出核心逻辑保持不变。
- **恢复关键帧编辑的直接操作能力**：关键帧表支持双击直接编辑时间、X、Y；表内显示时间进度、相邻位移变化量（ΔX/ΔY）、区间时长和阶段，便于理解整个动作走势。
- **保留完整关键帧工具**：添加末尾、插入中点、复制、删除、均匀分布时间、单帧 / 全部 X/Y 归零、时间与位置步进、重新生成预设；至少保留两个关键帧，时间必须非负且严格递增。
- **关键帧成为真实动作数据**：保存的手动关键帧会直接交给导出引擎，不会再因动作类型是内置预设而被重新计算覆盖；旧配置中没有保存关键帧的内置预设仍会按原参数自动生成，维持兼容。
- **动作模板批量管理**：保留新建、更新、复制、删除、按类型批量套用预设、恢复默认动作与保存模板。
- Spine 功能范围仍限定为明确识别到 SD / CH 标识的 `.skel`，自动执行默认关闭；下载、资源表和角色分类规则不变。

## 当前导出目录规则（v1.9）

- 普通高质量角色资源不建立 `HighQuality/`；`ch_<服饰ID>`、纯服饰 ID 子目录直接放在对应服饰 ID 目录下。
- 低质量资源建立 `LowQuality/`，位置在它所归属的角色家族目录内。
- `Adults/` 与 `General/` 等实际存在的不同内容版本保留独立目录。
- 旧版本已经导出的目录不会因为更新程序自动改变。要修整旧文件，请在 GUI 设置里把「导出目录」指向旧版本的 `exported_assets`，载入对应资源表后执行「修复历史导出分类 / 清理残留」。程序会优先读取当前设置的日志目录，并自动查找旧导出目录同级的 `logs`；修复仅移动 Manifest 登记的文件，未登记文件保留。

面向 Windows 的独立 GUI 下载与 Unity Bundle 导出工具。保留 V7 核心的 Catalog 解析、资源表生成、Bundle 下载、依赖展开、断点续传、Manifest、UnityPy 导出和临时清理能力；同时优化角色导出路径，并集成可选的 SD/CH Spine 全局动作工具。

## 启动

双击 `启动GUI.bat` 或 `start.bat`。需要 Python 3.10 或更新版本，Tkinter 应随官方 Windows Python 安装提供。批处理使用无 BOM 的 ASCII 编码和 Windows CRLF 换行，避免首行出现 `锘緻echo` 等乱码。启动器优先使用 `pyw` / `pythonw`，正常运行时不会额外打开命令行窗口；若 GUI 初始化或回调异常，会显示错误提示并将详细 traceback 写入 `logs/gui_error.log`。

资源表读取与 Bundle 下载只依赖 Python 标准库。Unity 资源导出需要 UnityPy 和 Pillow，可从“设置 → 安装 / 检查导出依赖”安装。Spine 全局动作处理另需游戏/资源版本兼容的 `SpineSkeletonDataConverter.exe`，在“Spine 全局动作”页选择即可。

## 角色分类导出（精简路径）

“设置 → Export 目录布局”默认是 `character`，按七位角色 ID 的前六位 / 末位拆分为特征 ID 与服饰 ID。例如 `1077004` 对应 `107700/1077004`。

```text
exported_assets/
├─ 111800/
│  └─ 1118001/
│     └─ Thumb_004.png
└─ 107700/
   └─ 1077004/
      ├─ ch_1077004/
      │  └─ ch_1077004.skel
      └─ 1077004/
         └─ 1077004.prefab
```

角色资源遵循原有 ID 归档思路，只在最外层增加“六位特征 ID”目录：`特征ID/服饰ID/文件`。会移除 `Chara_Thumb_004_a/AssetBundles/Sprites/Chara/` 等包装目录；已明确存在的 `ch_<服饰ID>` / `<服饰ID>` 分组会保留，但不再额外添加 `Sprites`、`Spine`、`Prefabs` 等类型子目录。若同一路径发生重名，会自动给文件名增加 Bundle 标识，避免覆盖。

无可识别角色 ID 的共享资源才按 `Sprites`、`Spine`、`Audio`、`Textures`、`Prefabs`、`Meshes`、`Fonts` 等类型目录归档，并去掉深层 Unity 包装路径。`original` 和 `id` 两种布局保留不变。

保留两种可选布局：

- `original`：保留原始资源路径。
- `id`：按资源 ID 归档；同一 ID 存在多个版本时会建立版本子目录。

路径冲突时会用 Bundle 标识生成区分后缀，避免覆盖不同资源。更换布局不会整体移动以前导出的文件；但重新导出同一个 Bundle 且本轮全部成功时，程序会参考历史导出 Manifest 清理该 Bundle 先前生成的过期路径。

## LowQuality 默认策略

设置中默认启用“跳过 LowQuality 版本”。下载与下载前预检查时，会排除 `version`、资源路径或 Bundle 名明确标记为 `LowQuality` 的记录；未标记高 / 低质量的资源不会被排除，因此不区分质量的角色资源仍会正常下载。该策略也用于依赖展开，避免明确标记为 LowQuality 的依赖被重新加入队列。资源表仍显示所有版本；关闭该选项后，可手动下载 LowQuality。

## SD / CH Spine 全局动作

新增“Spine 全局动作”页，将 `spine_global_action_tool` 的全局动作注入能力并入主程序：

- **默认关闭自动执行**：设置“导出完成后自动处理 SD/CH Spine”后，下载并导出或仅导出任务完成时才自动处理。
- **可独立执行**：不需要启动下载任务；扫描导出目录或手动添加 `.skel`，点击“执行列表中的 SD/CH Spine（独立运行）”即可运行。
- **目标范围严格限定**：只接收 `.skel` 文件，且文件名或路径中含明确的 `sd_` / `ch_`、`sd-` / `ch-` 等目标标识；会排除 `chara` 等相似但非目标前缀和其他 Spine。
- **源文件安全**：原始 `.skel` 不会被覆盖。修改后 JSON 与可选的 `_with_actions.skel` 输出至独立的动作输出目录，并镜像来源相对路径。
- 可配置 Converter、Root 骨骼名、动作模板、`--remove-curve`、是否输出 SKEL、保留中间 JSON，以及覆盖动作输出时备份 `.bak`。
- 动作模板支持跳跃、左右摇摆、上下弹动、左右抖动；动作参数与 Converter 设置写入 `config.json`。

自动处理需有效的 `SpineSkeletonDataConverter.exe`。如果没有配置 Converter，普通下载与导出仍会完成，只会在日志中提示并跳过 Spine 动作阶段。

## 资源表与记忆功能

- 支持 CSV、XLSX、JSON、TXT 分隔文本；可分析本地 Catalog 或在线更新 Catalog，并生成 CSV/XLSX。
- 支持关键字、分类、版本筛选及无版本记录筛选。
- 点击资源表任意列标题即可切换升序 / 降序；ID 列尽量按数字顺序排列。
- 支持 Ctrl/Shift 多选、勾选、全选筛选结果与反选筛选结果；勾选通过 Bundle 名称记忆。
- 默认记住资源表、搜索词、分类与版本筛选、排序列与方向、勾选项和窗口大小。
- `config.json` 保存常用设置；更改后可点“保存设置”，关闭程序时也会自动保存。

## 下载 / 任务管理

- 支持并行下载、下载线程数设置、已下载跳过、强制覆盖、自动依赖展开、失败重试、HTTP Range 断点续传、Manifest 和日志。
- 下载队列展示每个 Bundle 的状态、进度、已处理大小与速度。
- 支持暂停、继续、取消；取消会保留已完成文件，非空 `.part` 文件会留待后续续传。
- 支持预检查、下载并导出、仅导出本地资源。
- Export 使用隔离子进程 worker，保留完整 Sprite Texture 导出、Spine `.atlas.txt → .atlas` / `.skel.bytes → .skel` 处理、Audio/Video/Font/Mesh/Shader 等专用导出和 PPtr 解析。

## 项目文件

- `app.py`：GUI 入口、资源库、任务管理及 Spine 全局动作页。
- `core/smart_bundle_downloader_export_v7.py`：下载、Catalog 和 Unity 导出核心，包含精简的路径归档逻辑。
- `core/spine_global_actions.py`：SD/CH Spine 过滤、Converter 调用及全局动作注入引擎。
- `start.bat` / `启动GUI.bat`：一键启动入口（无 BOM、CRLF 换行）。
- `launch_gui.pyw`：无控制台启动包装器，记录静默启动异常与 GUI 回调异常。
- `config.json`：设置与界面状态记忆。
- `legacy/`：原版 V7 脚本和辅助启动文件副本。

## 规则边界

- 角色分类依据七位 ID 的“前六位 = 特征 ID、最后一位 = 服饰 ID”执行。
- LowQuality 过滤依据资源表质量字段与资源路径判断；未明确标记 LowQuality 的资源保留。
- Spine 动作处理只针对明确识别到 SD/CH 标识的 `.skel`，默认不会在导出后自动执行。


## v1.7 — 回归原有角色 ID 归档

- **SD / TF / ST 不再作为独立大分类**。像 `sd_1043002_SkeletonData`、`st_1043002_SkeletonData` 这样的名称会先提取七位角色 ID `1043002`，再归入 `104300/1043002`，不会生成 `st_1043002_SkeletonData/` 等额外目录。
- **相关文件保持在同一角色归档内**。Spine 文件名中的 `sd_`、`tf_`、`st_` 只保留在文件名中，不参与顶层目录命名；资源名相同发生冲突时只对文件名添加冲突标识，不会再按每个资源名建文件夹。
- **质量仍独立区分，但置于角色 ID 里面**。示例：`104300/1043002/HighQuality/...` 与 `104300/1043002/LowQuality/...`；同一角色不同前缀且质量相同的 Bundle 使用同一个角色/质量分组。
- **遵循原脚本的 ID 模式**。`id` 布局严格按 Catalog 的 `id/[version]/filename` 保存，不再用文件名前缀覆盖 Catalog ID；`original` 布局保留原始路径。
- **兼容未使用七位角色 ID 的 CH 资源**。类似 `ch_100021` 的旧式 CH Spine ID 继续按独立 ID 归档；没有可识别角色编号的 SD/TF/ST 文件按 `Spine` 等资源类型归档，不会按每个名称创建独立目录。
- **历史路径可修复**。在设置页点击“修复历史导出分类 / 清理残留”，程序依据历史 Export Manifest 与资源表，将旧版散落在前缀目录下的可追踪文件移到新的特征 ID / 服饰 ID 目录。无 Manifest 记录的文件保留，不会擅自删除。
- 质量优先读取明确的版本字段，再查 `asset_root` 与 Bundle 名；`HighQuality_` 归一为 `HighQuality`，LowQuality 默认过滤逻辑不变。

在项目目录执行 `python -m unittest discover -s tests -v` 可运行角色归档、质量分组、历史路径迁移等回归测试。


## v1.8 — SD / TF / ST Spine 目录与版本规则修正

- SD/TF/ST Spine 仍先归入 `特征ID/服饰ID`，然后在服饰 ID 下按规范化家族 ID 分组，例如 `tf_1043001_m0.skel` → `104300/1043001/tf_1043001/tf_1043001_m0.skel`。不会为 `_m0`、`SkeletonData` 等单个资源名创建目录。
- `HighQuality` 为默认版本，不建立 `HighQuality` 目录；`LowQuality` 版本单独进入 `LowQuality/`。
- 同一 Spine 家族存在 `Adults` 与 `General` 时，分别进入 `Adults/`、`General/`，避免两个版本混放；如果某版本同时是 LowQuality，则在相应版本目录内再建立 `LowQuality/`。
- `original` 与 `id` 导出布局不变；普通非 SD/TF/ST 资源沿用原有角色归档逻辑。
- 新增回归测试覆盖 `tf_1043001_m0.skel`、高低质量路径、Adults/General 分离和历史导出修复。


## v1.9 — 修正角色子目录与默认质量目录

- `HighQuality` 是默认版本，不建立 `HighQuality/` 目录；普通角色资源中的 `ch_<服饰ID>`、纯服饰 ID 子目录直接位于对应服饰 ID 下。
- `LowQuality` 只在实际导出低质量版本时建立，并放在相应角色子目录内部；显式 `Adults` / `General` 等内容版本仍分别归档。
- 历史导出修复会结合 Manifest 的 Unity 路径和现有物理输出路径识别 `ch_<服饰ID>` / 纯服饰 ID 子目录，将它们从旧 `HighQuality/` 包装层中提出来；未登记文件仍保留。


## v1.10 — 启动器编码与错误诊断修复

- 两个 GUI 启动批处理改为无 BOM 的 ASCII + Windows CRLF，解决 `锘緻echo` 被当作命令的问题。
- 启动顺序改为先检测 `pyw.exe`、再检测 `pythonw.exe`，最后回退到可显示错误的控制台 Python。
- 新增 `launch_gui.pyw`：GUI 启动异常写入 `logs/gui_error.log` 并弹窗提示；Tk 回调异常也会写入日志，避免窗口进程静默退出而无法定位原因。


### v1.12 更新

- 在动作编辑器中切换内置动作类型时，自动载入该类型对应的关键帧模板，按当前时长、幅度、次数和方向生成实际动作数据。
- 切换到“自定义”只更改模式并保留当前关键帧，方便继续手工编辑。
- 同类型下编辑过的关键帧不会被添加动作流程无条件覆盖；“套用预设”仍可主动重新生成。
- 修复更新已选动作时，类型标签可能变化但关键帧仍沿用旧类型的问题。
- 增加类型模板同步回归测试。
