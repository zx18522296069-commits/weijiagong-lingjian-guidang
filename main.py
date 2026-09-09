"""
未加工零件自动归档系统主入口

测试阶段：
1. 加载配置
2. 连接 Google Drive
3. 扫描目录结构
4. 输出扫描结果
5. 不修改生产文件
"""

import os
from modules.logger import get_logger
from modules.drive_manager import DriveManager

logger = get_logger()


def run():
    logger.info("开始执行未加工零件自动更新任务")

    test_mode = os.getenv("TEST_MODE", "false").lower() == "true"

    if not test_mode:
        logger.warning("当前仍建议先使用TEST_MODE进行首次真实连接验证")

    drive = DriveManager()

    if not drive.check_config():
        raise RuntimeError("缺少 Google Drive 配置，请设置 GitHub Secrets")

    logger.info("Google Drive配置检查通过")

    try:
        files = drive.list_drive_files()
        logger.info(f"扫描到文件数量: {len(files)}")

        for item in files:
            logger.info(f"文件: {item.get('name')} | ID: {item.get('id')}")

    except Exception as e:
        logger.error(f"Drive扫描失败: {e}")
        raise

    if test_mode:
        logger.info("只读测试完成，未执行写入操作")
    else:
        logger.info("扫描完成，等待正式处理流程接入")


if __name__ == "__main__":
    run()
