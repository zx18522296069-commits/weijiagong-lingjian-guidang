# GitHub Actions 云端授权配置

## 必需 Secrets

在 GitHub Repository Settings → Secrets and variables → Actions 中配置：

### GOOGLE_SERVICE_ACCOUNT_JSON

内容：Google Cloud 服务账号 JSON 密钥。

用途：
- GitHub Actions 无人值守访问 Google Drive。

目录和正式文件 ID 在 `.github/workflows/update_parts.yml` 中作为非密钥环境变量配置；不得把服务账号 JSON 或私钥写入这些变量。

## 安全规则

- 不上传生产Excel到GitHub。
- 不把账号密码写入代码。
- 仅通过Secrets访问。

## 完成标准

1. 云端可以读取Drive。
2. 云端可以识别订单和完成文件。
3. 正式写回后可以回读并验证两个固定Excel。
4. 定时生产任务只在回归测试通过后执行。
