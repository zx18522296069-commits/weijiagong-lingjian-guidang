"""
未加工零件自动归档系统主入口

执行流程：
1. 加载配置
2. 连接 Google Drive
3. 扫描订单和完成文件
4. 读取 Excel
5. 数据校验
6. 累计加工计算
7. 生成结果文件
8. 上传结果
9. 归档已录入文件
10. 输出日志
"""

import os
from modules.logger import get_logger
from modules.drive_manager import DriveManager

logger = get_logger()


def run():
    logger.info("开始执行未加工零件自动更新任务")

    test_mode = os.getenv("TEST_MODE", "false").lower() == "true"
    if test_mode:
        logger.info("当前运行模式：只读测试模式，不修改Drive文件")

    drive = DriveManager()

    if not drive.check_config():
        raise RuntimeError("缺少 Google Drive 配置，请设置 GitHub Secrets")

    logger.info("Google Drive配置检查通过")

    # 下一阶段接入真实Drive读取：
    # 1. 获取订单文件
    # 2. 获取拆图完成文件
    # 3. 读取Excel
    # 4. 累计加工计算
    # 5. 生成结果
    # 6. 写回Drive
    # 7. 归档完成文件

    if test_mode:
        logger.info("只读测试完成，未执行写入操作")
    else:
        logger.info("生产模式准备执行")


if __name__ == "__main__":
    run()
