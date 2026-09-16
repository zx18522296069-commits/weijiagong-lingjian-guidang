# 未加工零件自动归档系统

系统采用：**增量读取，完整核验，批量入账，全量重建视图，幂等归档。**

## 事实源

- Google Drive“正在加工”中的订单原始汇总表：当前订单事实源。
- `累计加工台账.xlsx`：永久加工事实，保存累计状态、加工流水、板材入账记录、异常和已退出完成订单历史。
- `当前待加工零件.xlsx`：动态视图，由当前订单事实 + 永久累计事实完整重建，不作为永久去重依据。
- `runtime_state/state.json`：仅为可随时重建的性能缓存，不能替代 Google Drive 正式业务数据。

## 增量流程

每次运行：

1. 只读取 Drive metadata，按 `file_id + modifiedTime + md5Checksum + parent_folder_id + order_container_name` 分类 `unchanged / added / modified / removed`。
2. unchanged 直接使用 runtime state 中的标准化 records，不下载 Excel；added/modified 只下载变化文件并替换对应缓存。
3. removed 从活动 source cache 移除；已完成历史永久保留，未完成历史冻结并记录异常。
4. 累计台账先比对正式文件 ID + MD5；与 runtime state 一致时直接使用缓存中的 existing state/flows/board records/anomalies，不下载累计 Excel。
5. 只扫描“拆图结果”根目录的完成文件，不递归扫描“已录入数量”。每张板独立整板预校验，任意一行阻断则整板 0 入账。
6. 所有 accepted 板材一次性合并累计变化；`当前待加工零件.xlsx`始终由完整当前状态重新生成，不局部改单元格。
7. 本地用 openpyxl 检查工作表、数量、剩余、状态、颜色和板材入账记录。
8. 先上传累计台账并用 Drive metadata `md5Checksum + size` 对比本地 MD5；再上传当前待加工零件并做相同确认。禁止写回后重新下载整个 Excel 验证。
9. 两份正式文件都确认成功后，才统一归档 accepted + reconciled 文件；随后正式生产运行才保存新的 runtime state。
10. 如果累计台账上传成功而当前待加工文件失败：抛错、不归档、不提交本轮 runtime state。下一轮因累计 MD5 变化会重新读取正式累计台账，并依靠“完整板材号 + 内容指纹”避免重复扣减。

## 板材与异常规则

- 完整板材号相同 + 内容指纹相同：已入账重复文件，只补归档，不重复累计。
- 完整板材号相同 + 内容指纹不同：`同板材号内容冲突`，阻断、记录异常、文件留根目录。
- `#2323`、`#2323-1`、`#2323-2` 为不同完整板材号，分别处理。
- 订单/图号/厚度/坡口无法唯一确认、基础数量或基础重量不符、文件损坏等都属于整板阻断。
- 基础资料均正确但累计后 `当前剩余 < 0` 不阻断，照实入账；状态为`超加工/待核查`，负数保留，整行红色。
- 状态统一为：`未加工`、`部分完成`、`已完成`、`超加工/待核查`。未加工/部分完成黄色，已完成绿色，超加工红色。
- 订单未整体完成时，该订单全部零件继续保留，包括当前剩余为 0 的已完成行；只有全订单所有行 `当前剩余 == 0` 且无阻止完成的异常时，整个订单才从当前待加工视图退出。

## runtime state

`runtime_state/state.json` 至少保存：

- `state_schema_version`
- 每个订单源的 file id、name、modifiedTime、md5Checksum、parent folder、order container、标准化 source records/错误状态
- 正式累计台账 file id、modifiedTime、md5Checksum、size
- 已解析累计 `state / historical_rows / flows / board_records / anomalies / posted_boards / posted_board_keys / legacy_posted_boards`

cache 缺失、损坏或 schema 不兼容时自动回退完整初始化。dry-run 可以读取正式 cache 加速，但不得保存/覆盖 production cache。

## 定时与手动执行

当前 `.github/workflows/update_parts.yml` 保留两次 GitHub Actions 定时：

- `0 4 * * *` → 北京时间 12:00
- `0 9 * * *` → 北京时间 17:00

旧的 17:25、22:00 不再使用。

`workflow_dispatch` 继续保留。手动 `test` 是 **dry-run / 只读测试**：读取正式数据和正式 runtime cache，但不得写正式 Drive、不得移动拆图结果、不得覆盖 production cache。`production` 用于正式补跑或故障恢复。

## 并发

```yaml
concurrency:
  group: weijiagong-lingjian-guidang
  cancel-in-progress: false
```

本次不改为 true。先依靠幂等、事务恢复、增量和无变化快速退出缩短运行时间。

## 性能日志

每次运行记录：订单 metadata 的 unchanged/added/modified/removed、实际订单正文下载数、累计 cache 命中/MD5变化、拆图结果新入账/补归档/阻断、本轮 Drive 正文下载总数，以及 metadata 扫描、cache 恢复、Drive 下载、Excel 解析、整板核验、状态重建、Excel 生成、本地验证、Drive 上传、metadata/hash 验证和总耗时。
