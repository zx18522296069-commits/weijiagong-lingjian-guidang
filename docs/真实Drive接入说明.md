# 真实 Google Drive 接入

## 已确认目录

- 正在加工
- 拆图结果
- 已录入数量

## 输出文件

- 当前待加工零件.xlsx
- 累计加工台账.xlsx

## 接入原则

1. GitHub Actions 只负责执行程序。
2. 生产数据继续保留在 Google Drive。
3. 不上传生产文件到 GitHub。
4. 使用授权凭证访问指定 Drive 目录。
5. 执行前读取，执行后更新并归档。

## 已完成配置

- [x] Google Drive API 服务账号
- [x] GitHub Secret
- [x] 云端只读测试
- [x] 正式生产写入与回读验证
- [x] 北京时间17:25、22:00定时生产任务

服务账号密钥只保存在 GitHub Secret `GOOGLE_SERVICE_ACCOUNT_JSON` 中，不写入仓库。
