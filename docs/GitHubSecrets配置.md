# GitHub Actions 云端授权配置

## 必需 Secrets

在 GitHub Repository Settings → Secrets and variables → Actions 中配置：

### GOOGLE_SERVICE_ACCOUNT_JSON

内容：Google Cloud 服务账号 JSON 密钥。

用途：
- GitHub Actions 无人值守访问 Google Drive。

### DRIVE_ROOT_FOLDER_ID

值：生产文件根目录ID。

当前已确认：
- 正在加工文件夹ID：1PGM1ZEbHk2v3CAZPfwHdqC2EV_kHNzJH

## 安全规则

- 不上传生产Excel到GitHub。
- 不把账号密码写入代码。
- 仅通过Secrets访问。

## 完成标准

1. 云端可以读取Drive。
2. 云端可以识别订单和完成文件。
3. 测试通过后开启自动写入。
