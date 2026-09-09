"""
未加工零件自动归档系统主入口

测试阶段：
1. 加载配置
2. 连接 Google Drive
3. 扫描目录结构
4. 下载Excel测试读取
5. 输出扫描结果
6. 不修改生产文件
"""

import os
import tempfile
from pathlib import Path

from modules.logger import get_logger
from modules.drive_manager import DriveManager
from modules.excel_reader import read_excel

logger = get_logger()


def run():
    logger.info("开始执行未加工零件自动更新任务")

    test_mode = os.getenv("TEST_MODE", "false").lower() == "true"

    drive = DriveManager()

    if not drive.check_config():
        raise RuntimeError("缺少 Google Drive 配置，请设置 GitHub Secrets")

    logger.info("Google Drive配置检查通过")

    try:
        files = drive.scan_folder_recursive()
        logger.info(f"递归扫描文件数量: {len(files)}")

        summary_files = []
        complete_files = []

        for item in files:
            name = item.get("name", "")
            path = item.get("path", "")

            if "完成" in name:
                complete_files.append(item)
            if "汇总表" in name:
                summary_files.append(item)

            logger.info(f"文件: {name} | 路径: {path}")

        logger.info(f"订单汇总表数量: {len(summary_files)}")
        logger.info(f"完成文件数量: {len(complete_files)}")

        # 测试下载第一份Excel，验证Drive到Excel读取链路
        test_files = summary_files + complete_files
        if test_files:
            with tempfile.TemporaryDirectory() as temp_dir:
                target = Path(temp_dir) / test_files[0]["name"]
                drive.download_file(test_files[0]["id"], str(target))
                data = read_excel(target)
                logger.info(f"Excel读取测试成功: {len(data) if data is not None else 0} 行")

    except Exception as e:
        logger.error(f"Drive/Excel测试失败: {e}")
        raise

    if test_mode:
        logger.info("只读测试完成，未执行写入操作")
    else:
        logger.info("扫描和读取完成，等待正式处理流程接入")


if __name__ == "__main__":
    run()
