from __future__ import annotations

import sys


def main() -> int:
    from billing_app.ui.app import BillingApp
    app = BillingApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
