# 职责：凭证制单工作流编排壳——组合 save/match/verify 三个 mixin,代理到 processor
# 不做什么：业务方法已拆到 nc_voucher_{save,match,verify};此处只留组合与 processor 代理
# 允许依赖层：core.nc_voucher_{save,match,verify}
# 谁不应该 import：其它 nc_*_workflow 不应 import(import-linter 独立性约束)

from core.nc_voucher_match import NCVoucherMatchMixin
from core.nc_voucher_save import NCVoucherSaveMixin
from core.nc_voucher_verify import NCVoucherVerifyMixin


class NCVoucherWorkflow(NCVoucherSaveMixin, NCVoucherMatchMixin, NCVoucherVerifyMixin):
    def __init__(self, processor):
        super().__setattr__("processor", processor)

    def __getattr__(self, name):
        return getattr(self.processor, name)

    def __setattr__(self, name, value):
        if name == "processor":
            super().__setattr__(name, value)
            return
        setattr(self.processor, name, value)
