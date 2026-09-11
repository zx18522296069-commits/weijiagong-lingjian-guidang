# 未加工零件自动归档系统

本项目已将未加工零件统计流程迁移到 GitHub Actions 云端自动执行，Mac 关机不影响运行。

## 目标功能

- Google Drive 自动读取订单数据
- 拆图结果自动识别
- 按板材编号累计加工数量
- 保留“废”字作为独立编号
- 防止重复录入
- 自动更新：
  - 当前待加工零件.xlsx
  - 累计加工台账.xlsx
- 自动归档已录入完成文件
- 自动生成执行日志

## 自动执行

计划：

- 北京时间 17:25
- 北京时间 22:00

定时任务以正式生产模式运行；手动运行默认是只读测试，也可显式选择正式生产。

## 当前状态

- GitHub Secret 与 Google Drive 服务账号已接通。
- 正式 Drive 读取、写回及回读验证已通过。
- 代码自检和业务回归测试已启用。
- 生产 Excel 只保存在 Google Drive，不保存为 GitHub Actions artifact；GitHub 仅保存代码与执行日志。
- 无法唯一匹配或原始资料缺失的板材会阻断入账并保留在“拆图结果”根目录，禁止猜测处理。

## Google Drive 凭据

个人网盘写入推荐使用用户 OAuth，两个自动化仓库配置同一组三项 Actions secrets：

- `GOOGLE_OAUTH_CLIENT_ID`
- `GOOGLE_OAUTH_CLIENT_SECRET`
- `GOOGLE_OAUTH_REFRESH_TOKEN`

三项齐全时程序优先使用 OAuth；未设置时兼容原 `GOOGLE_SERVICE_ACCOUNT_JSON`。服务账号没有个人 Drive 存储额度，因此只能把它作为 Shared Drive 或只更新既有文件的后备方式。
