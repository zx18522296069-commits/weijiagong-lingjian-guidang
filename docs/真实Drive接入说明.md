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

## 后续配置

- Google Drive API 服务账号
- GitHub Secrets
- 第一次云端只读测试
- 第一次完整写入测试
