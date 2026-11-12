class ConfigError(Exception):
    """工作流配置有误（运行前就能发现）"""


class StepError(Exception):
    """执行某一步时出错（列找不到、文件打不开等），会中止整个工作流"""
