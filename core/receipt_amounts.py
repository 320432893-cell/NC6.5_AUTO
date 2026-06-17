# 职责：提供收款单金额口径计算，隔离 Excel 到账净额、手续费和 NC 原币金额的转换
# 不做什么：不读取 Excel，不写 NC，不解析金额字符串，不决定匹配规则
# 允许依赖层：收款单数据模型或具备 raw_amount/fee 属性的计划行对象
# 谁不应该 import：无收款单金额语义的通用基础模块不应 import


def receipt_nc_amount(row):
    return row.raw_amount + row.fee


def receipt_net_amount(row):
    return row.raw_amount
