# 职责：收款单明细主行/手续费行固定业务常量的单一来源(SSOT)——科目、业务类型、结算方式
# 不做什么：不读 Excel/配置、不构造 business dict、不录入 NC、不做 JAB/GUI 动作
# 允许依赖层：纯常量模块，无内部依赖；可被 core 与 tools 任意 import
# 谁不应该 import：本模块不 import 任何项目内模块，避免循环依赖

# 口径来自 README「明细表当前已验证的填写顺序」与「收款单自制录入明细表」：
#   主行：收款业务类型=货款、科目=1002、结算方式=网银
#   手续费行：手续费、660305、网银
# 这些是当前授权主体真实保存 T0 已验证的固定值，不随主体/币种变化。

# 主行：科目 1002(银行存款)、收款业务类型 货款。
RECEIPT_MAIN_SUBJECT = "1002"
RECEIPT_MAIN_BUSINESS_TYPE = "货款"

# 手续费行：科目 660305(手续费)、收款业务类型 手续费。
RECEIPT_FEE_SUBJECT = "660305"
RECEIPT_FEE_BUSINESS_TYPE = "手续费"

# 表头与明细统一的结算方式：网银。
RECEIPT_SETTLEMENT = "网银"
