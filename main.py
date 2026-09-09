"""
未加工零件自动归档系统主入口

当前阶段：
1. 加载配置
2. 连接 Google Drive
3. 扫描目录结构
4. 下载并读取Excel
5. 建立累计加工索引
6. 生成当前剩余明细
7. 上传结果文件（正式模式）
8. 测试模式不修改生产文件
"""

import os
import tempfile
from pathlib import Path

from modules.logger import get_logger
from modules.drive_manager import DriveManager
from modules.excel_reader import read_excel
from modules.process_parts import build_processed_index, build_remaining_records
from modules.excel_generator import generate_report

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

        summary_files = []
        complete_files = []

        for item in files:
            name = item.get("name", "")
            if "完成" in name:
                complete_files.append(item)
            if "汇总表" in name:
                summary_files.append(item)

        completion_records = []
        summary_records = []

        with tempfile.TemporaryDirectory() as temp_dir:
            for item in complete_files:
                target = Path(temp_dir) / item["name"]
                drive.download_file(item["id"], str(target))
                rows = read_excel(target)
                completion_records.extend(rows or [])

            for item in summary_files:
                target = Path(temp_dir) / item["name"]
                drive.download_file(item["id"], str(target))
                rows = read_excel(target)
                summary_records.extend(rows or [])

            processed_index = build_processed_index(completion_records)
            remaining = build_remaining_records(summary_records, processed_index)

            output = Path(temp_dir) / "当前待加工零件.xlsx"
            generate_report(remaining, output)

            logger.info(f"生成结果: {output}")

            if not test_mode:
                target_folder = os.getenv("RESULT_FOLDER_ID", "")
                if target_folder:
                    drive.upload_file(str(output), target_folder)
                    logger.info("结果文件已上传Drive")
                else:
                    logger.warning("未设置 RESULT_FOLDER_ID，跳过上传")

    except Exception as e:
        logger.error(f"处理失败: {e}")
        raise

    if test_mode:
        logger.info("只读测试完成，未执行写入")


if __name__ == "__main__":
    run()
