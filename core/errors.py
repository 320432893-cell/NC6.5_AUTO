class NCWorkflowError(RuntimeError):
    """Base error for NC workflow failures."""


class WorkflowStateError(NCWorkflowError):
    """Raised when the current NC page/window state blocks the requested action."""


class TableMatchError(NCWorkflowError):
    """Raised when Excel rows cannot be matched to NC table rows safely."""


class ContractViolation(NCWorkflowError):
    """Raised when a workflow post-condition is not met."""


class JABControlNotFound(NCWorkflowError):
    """Raised when a required JAB control or window cannot be found."""


class JABActionError(NCWorkflowError):
    """Raised when a JAB action returns failure."""


class ExcelLockedError(NCWorkflowError):
    """Raised when Excel output cannot be written because the workbook is locked."""


class ExcelPreflightError(NCWorkflowError):
    """Raised when local Excel validation fails before any NC/JAB access."""
