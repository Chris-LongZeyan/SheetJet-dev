class SheetJetError(Exception):
    """An actionable workbook or operation error."""


class UnsupportedOperation(SheetJetError):
    """The requested operation cannot be performed with the fidelity contract."""


class BudgetExceeded(SheetJetError):
    """Narrow the request or explicitly raise its budget."""
