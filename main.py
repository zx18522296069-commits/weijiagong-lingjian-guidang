"""
未加工零件自动归档系统主入口

执行流程：
1. 加载配置
2. 连接Google Drive
3. 扫描正在加工和拆图结果
4. 读取完成文件
5. 更新累计加工台账
6. 更新当前待加工零件
7. 完成文件归档
8. 输出执行日志
"""

from modules.logger import get_logger

logger = get_logger()


def run():
    logger.info("开始执行未加工零件自动更新任务")

    # 正式业务流程将在此调用：
    # drive_manager
    # process_parts
    # excel_generator
    # archive_manager

    logger.info("自动执行主流程框架完成")


if __name__ == "__main__":
    run()
