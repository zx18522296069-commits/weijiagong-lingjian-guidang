"""
未加工零件自动归档系统主入口

正式执行流程：
1. 加载配置
2. 连接 Google Drive
3. 扫描正在加工和拆图结果
4. 读取完成文件
5. 数据校验
6. 更新累计加工台账
7. 更新当前待加工零件
8. 归档已录入文件
9. 输出日志
"""

from modules.logger import get_logger

logger = get_logger()


def run():
    logger.info("开始执行未加工零件自动更新任务")

    # 模块调用入口
    # drive_manager:
    #   获取Google Drive文件
    # excel_reader:
    #   读取订单和完成数据
    # validator:
    #   数据核对
    # process_parts:
    #   累计加工计算
    # excel_generator:
    #   生成输出表
    # archive_manager:
    #   完成归档

    logger.info("流程模块加载完成，等待生产环境配置")


if __name__ == "__main__":
    run()
