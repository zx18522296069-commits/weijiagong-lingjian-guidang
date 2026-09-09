"""
未加工零件自动归档系统主入口

流程：
1. 加载配置
2. 连接 Google Drive
3. 扫描目录结构
4. 下载并读取Excel
5. 建立累计加工索引
6. 生成当前剩余明细
7. 上传结果文件
8. 整单完成后执行归档
"""

import os
import tempfile
from pathlib import Path

from modules.logger import get_logger
from modules.drive_manager import DriveManager
from modules.excel_reader import read_excel
from modules.process_parts import build_processed_index, build_remaining_records, is_order_completed
from modules.excel_generator import generate_report
from modules.archive_manager import archive_completed_file

logger = get_logger()


def run():
    logger.info("开始执行未加工零件自动更新任务")

    test_mode = os.getenv("TEST_MODE", "false").lower() == "true"
    drive = DriveManager()

    if not drive.check_config():
        raise RuntimeError("缺少 Google Drive 配置")

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
            completion_records.extend(read_excel(target) or [])

        for item in summary_files:
            target = Path(temp_dir) / item["name"]
            drive.download_file(item["id"], str(target))
            summary_records.extend(read_excel(target) or [])

        processed_index = build_processed_index(completion_records)
        remaining = build_remaining_records(summary_records, processed_index)

        output = Path(temp_dir) / "当前待加工零件.xlsx"
        generate_report(remaining, output)

        if not test_mode:
            result_folder = os.getenv("RESULT_FOLDER_ID", "")
            if result_folder:
                drive.upload_file(str(output), result_folder)

            archive_folder = os.getenv("ARCHIVE_FOLDER_ID", "")
            if archive_folder:
                if is_order_completed(remaining):
                    for item in complete_files:
                        record = archive_completed_file(item["id"], archive_folder)
                        if record:
                            drive.move_file(item["id"], archive_folder)
                            logger.info(f"归档完成: {item['name']}")
                else:
                    logger.info("订单未全部完成，跳过归档")

    if test_mode:
        logger.info("测试模式结束，未执行上传和归档")


if __name__ == "__main__":
    run()
